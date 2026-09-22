using System;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Scribble.Configuration;

namespace Scribble.Testing
{
    public static class TestLabStressBudget
    {
        // Final native validation uses the remaining balance on the same
        // no-reset key. A final targeted validation may use the remaining
        // balance while retaining $0.05 below the $30 hard cap. Every model
        // request rechecks the live key before dispatch.
        public const decimal MaximumTotalUsd = 30m;
        public const decimal MaximumCheckpointUsageUsd = 30m;
        public static async Task GuardRequestAsync(AppSettings actual, string requestedModel, CancellationToken cancel)
        {
            var state = TestLabSuite.Active();
            if (TestLab.Status()?.suite_id != "scribble-stress-v1" && state?.fixtureSuiteId != "scribble-stress-v1") return;
            if (state == null || TestLab.ActiveRunId() == null)
                throw new InvalidOperationException("The stress request no longer owns an active suite and case.");
            ValidateRequestSettings(actual, new SettingsStore().Load(), requestedModel);
            await CheckAsync(state, cancel).ConfigureAwait(true);
        }
        public static void ValidateRequestSettings(AppSettings actual, AppSettings expected, string requestedModel)
        {
            if (actual == null || expected == null || string.IsNullOrWhiteSpace(actual.ApiKey) ||
                !string.Equals(actual.ApiKey, expected.ApiKey, StringComparison.Ordinal) ||
                !string.Equals(actual.BaseUrl, expected.BaseUrl, StringComparison.Ordinal) ||
                actual.Model != "qwen/qwen3.8-27b" || requestedModel != actual.Model || expected.Model != actual.Model)
                throw new InvalidOperationException("The Office pane's model connection differs from the verified stress-test configuration. Reload the pane before retrying. No model request was submitted.");
        }
        public static async Task CheckAsync(SuiteState state, CancellationToken cancel)
        {
            if (state.fixtureSuiteId != "scribble-stress-v1") return;
            ThrowIfProviderStopped(state);
            var settings = new SettingsStore().Load();
            Uri endpoint;
            if (!settings.IsConfigured || !Uri.TryCreate(settings.BaseUrl, UriKind.Absolute, out endpoint) ||
                endpoint.Scheme != "https" || endpoint.Host != "openrouter.ai" || !endpoint.IsDefaultPort ||
                !string.IsNullOrEmpty(endpoint.UserInfo) || !string.IsNullOrEmpty(endpoint.Query) ||
                endpoint.AbsolutePath.TrimEnd('/') != "/api/v1" || settings.Model != "qwen/qwen3.8-27b")
                throw new InvalidOperationException("Stress tests require the configured Qwen3.8 27B OpenRouter endpoint and the approved no-reset key capped at $30 total. No model request was submitted.");
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancel))
            using (var handler = new HttpClientHandler { AllowAutoRedirect = false })
            using (var client = new HttpClient(handler))
            using (var request = new HttpRequestMessage(HttpMethod.Get, "https://openrouter.ai/api/v1/key"))
            {
                timeout.CancelAfter(TimeSpan.FromSeconds(15));
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", settings.ApiKey);
                using (var response = await client.SendAsync(request, timeout.Token).ConfigureAwait(true))
                {
                    if (!response.IsSuccessStatusCode)
                        throw new InvalidOperationException("Cannot verify the API spending cap (HTTP " + (int)response.StatusCode + "). No next case was submitted.");
                    var text = await response.Content.ReadAsStringAsync().ConfigureAwait(true);
                    var key = Validate(text);
                    // Only billing totals are retained; never the key, label, or account details.
                    File.AppendAllText(Path.Combine(state.folder, "usage.jsonl"), TestLab.Serialize(new {
                        utc = DateTime.UtcNow.ToString("O"), case_id = state.caseId, model = settings.Model,
                        limit_usd = key.limit, remaining_usd = key.limit_remaining, usage_usd = key.usage,
                        source = "OpenRouter /api/v1/key", total_limit_no_reset = true,
                        checkpoint_stop_usage_usd = MaximumCheckpointUsageUsd
                    }) + Environment.NewLine, new UTF8Encoding(false));
                }
            }
        }
        public static void RecordProviderResponse(int httpStatus)
        {
            if (!StopsSuite(httpStatus)) return;
            var state = TestLabSuite.Active();
            if (state?.fixtureSuiteId != "scribble-stress-v1" || string.IsNullOrEmpty(state.runId) ||
                state.runId != TestLab.ActiveRunId()) return;
            RecordProviderResponse(state, httpStatus);
        }
        public static void RecordProviderResponse(SuiteState state, int httpStatus)
        {
            if (state?.fixtureSuiteId != "scribble-stress-v1" || !StopsSuite(httpStatus)) return;
            var path = ProviderStopPath(state);
            // The first provider rejection is immutable for this attempt. A
            // funded key's limit_remaining does not prove account credit exists.
            if (File.Exists(path)) return;
            try
            {
                using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
                using (var writer = new StreamWriter(file, new UTF8Encoding(false)))
                    writer.Write(TestLab.Serialize(new StressProviderStop {
                        suite_id = state.id, case_id = state.caseId, http_status = httpStatus,
                        utc = DateTime.UtcNow.ToString("O", CultureInfo.InvariantCulture)
                    }));
            }
            catch (IOException) when (File.Exists(path)) { /* Another response already latched this attempt. */ }
        }
        private static bool StopsSuite(int status) => status == 401 || status == 402 || status == 403;
        private static string ProviderStopPath(SuiteState state)
        {
            if (string.IsNullOrWhiteSpace(state.id) || string.IsNullOrWhiteSpace(state.folder))
                throw new InvalidOperationException("The stress attempt has no valid provider-stop location.");
            return Path.Combine(state.folder, "provider-stop.json");
        }
        private static void ThrowIfProviderStopped(SuiteState state)
        {
            var path = ProviderStopPath(state);
            if (!File.Exists(path)) return;
            StressProviderStop stop;
            try
            {
                if (new FileInfo(path).Length > 4096) throw new InvalidDataException("Oversized provider-stop record.");
                stop = TestLabSuite.Read<StressProviderStop>(path);
                if (stop == null || stop.suite_id != state.id || !StopsSuite(stop.http_status))
                    throw new InvalidDataException("Invalid provider-stop record.");
            }
            catch (Exception error) when (error is IOException || error is UnauthorizedAccessException || error is ArgumentException || error is InvalidOperationException)
            {
                throw new InvalidOperationException("This stress attempt has an unreadable provider-stop record. Remaining tests were not submitted.", error);
            }
            throw new InvalidOperationException("The provider rejected inference with HTTP " + stop.http_status +
                " in case " + stop.case_id + ". This attempt is stopped; remaining tests were not submitted. Verify account credit or key access before starting a new attempt.");
        }
        public static StressKeyBudget Validate(string response)
        {
            if (string.IsNullOrWhiteSpace(response) || response.Length > 65536)
                throw new InvalidDataException("Missing or oversized API budget response.");
            var envelope = new JavaScriptSerializer().Deserialize<StressKeyEnvelope>(response);
            var key = envelope?.data;
            if (key == null || !key.limit.HasValue || key.limit <= 0 || key.limit > MaximumTotalUsd ||
                !string.IsNullOrEmpty(key.limit_reset) || !key.limit_remaining.HasValue || !key.usage.HasValue ||
                key.limit_remaining < 0 || key.limit_remaining > key.limit || key.usage < 0 ||
                key.is_management_key || key.is_provisioning_key)
                throw new InvalidOperationException("The API key must be an inference key with a positive total limit of at most $30 and no periodic reset. Unverified limits cannot start stress tests.");
            var reserve = key.limit >= 20m ? .05m : .25m;
            if (key.limit_remaining <= reserve)
                throw new InvalidOperationException("API budget nearly exhausted: $" + key.limit_remaining.Value.ToString("0.000", CultureInfo.InvariantCulture) + " remains. Remaining tests were not submitted.");
            if (key.usage >= MaximumCheckpointUsageUsd - reserve)
                throw new InvalidOperationException("Golden Showcase checkpoint budget nearly exhausted: total key usage is $" + key.usage.Value.ToString("0.000", CultureInfo.InvariantCulture) + ". The checkpoint stops before $" + MaximumCheckpointUsageUsd.ToString("0.00", CultureInfo.InvariantCulture) + " total usage; remaining tests were not submitted.");
            return key;
        }
    }
    public sealed class StressProviderStop
    {
        public string suite_id { get; set; }
        public string case_id { get; set; }
        public int http_status { get; set; }
        public string utc { get; set; }
    }
    public sealed class StressKeyEnvelope { public StressKeyBudget data { get; set; } }
    public sealed class StressKeyBudget
    {
        public decimal? limit { get; set; }
        public decimal? limit_remaining { get; set; }
        public decimal? usage { get; set; }
        public string limit_reset { get; set; }
        public bool is_management_key { get; set; }
        public bool is_provisioning_key { get; set; }
    }
}
