using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Scribble.Configuration;
using Scribble.Security;

namespace Scribble.Chat
{
    public sealed class OpenAiCompatibleClient : IDisposable
    {
        internal static readonly TimeSpan CompletionRequestTimeout =
            TimeSpan.FromMinutes(3);
        internal static readonly TimeSpan OpenRouterQwenCompletionRequestTimeout =
            TimeSpan.FromMinutes(5);
        private readonly TimeSpan _completionRequestTimeout;
        private readonly HttpClient _httpClient;
        // Vision requests carry multi-megabyte base64 image parts; the
        // serializer's 2 MB default would reject them. Responses stay
        // bounded separately by ReadBoundedAsync.
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer
            {
                MaxJsonLength = int.MaxValue
            };

        private readonly GeminiCodeAssistGateway _gemini =
            new GeminiCodeAssistGateway();
        private readonly object _optionalToolControlSync = new object();
        private readonly HashSet<string> _optionalToolControlUnsupported =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, DateTime> _emptyResponseCircuits = new Dictionary<string, DateTime>(StringComparer.Ordinal);

        public OpenAiCompatibleClient()
            : this(CompletionRequestTimeout)
        {
        }

        internal OpenAiCompatibleClient(TimeSpan completionRequestTimeout)
        {
            if (completionRequestTimeout <= TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(
                    nameof(completionRequestTimeout));
            _completionRequestTimeout = completionRequestTimeout;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            // .NET Framework HTTP latency defaults hurt every
            // request: Expect: 100-continue adds a round trip per
            // POST (and some proxies stall on it), Nagle delays
            // small writes, and the 100-second idle timeout forces
            // a fresh TLS handshake on the first message after a
            // short pause.
            ServicePointManager.Expect100Continue = false;
            ServicePointManager.UseNagleAlgorithm = false;
            ServicePointManager.MaxServicePointIdleTime = 300000;
            _httpClient = new HttpClient
            {
                Timeout = Timeout.InfiniteTimeSpan
            };
        }

        // Lets the settings window prime the Gemini token cache
        // right after a browser sign-in completes.
        public GeminiCodeAssistGateway GeminiGateway
        {
            get { return _gemini; }
        }

        // Streaming completion: Gemini streams text deltas through
        // onTextDelta as they arrive; the OpenAI-compatible path
        // falls back to the buffered call (local servers vary too
        // much in SSE support to stream them safely). The returned
        // message is always the complete response.
        public async Task<ChatCompletionResponseMessage>
            CompleteStreamingAsync(
                AppSettings settings,
                ChatCompletionRequest requestModel,
                Action<string> onTextDelta,
                CancellationToken cancellationToken)
        {
            if (RoutesToGemini(settings, requestModel))
            {
                await Scribble.Testing.TestLabStressBudget.GuardRequestAsync(settings, requestModel?.model, cancellationToken).ConfigureAwait(true);
                return await _gemini.GenerateStreamAsync(
                    _httpClient,
                    settings,
                    requestModel,
                    onTextDelta,
                    cancellationToken).ConfigureAwait(true);
            }

            return await CompleteAsync(
                settings,
                requestModel,
                cancellationToken).ConfigureAwait(true);
        }

        // Transport follows the model the user picked, never the
        // Gemini tick: that switch only decides whether Gemini
        // models are offered at all. Picking qwen while Gemini
        // sign-in is on must reach the local endpoint, and picking
        // a gemini model must reach Google - the mismatch was what
        // produced HTTP 400s from a local server being handed a
        // gemini id.
        private static bool RoutesToGemini(
            AppSettings settings,
            ChatCompletionRequest requestModel)
        {
            var model = requestModel?.model;
            if (!GeminiCodeAssistGateway.IsGeminiModel(model))
            {
                return false;
            }

            if (AdminPolicy.GeminiDisabled)
            {
                throw new AiEndpointException(
                    "GEMINI_DISABLED_BY_POLICY",
                    "Google Gemini is unavailable in this build. " +
                    "Choose a model served by your endpoint.");
            }

            if (settings == null || !settings.UseGeminiSignIn)
            {
                throw new AiEndpointException(
                    "GEMINI_MODEL_NOT_ENABLED",
                    "The selected model '" +
                    TextBoundary.SingleLine(model, 80) +
                    "' is a Google Gemini model, but Gemini " +
                    "sign-in is off. Turn on Gemini sign-in in " +
                    "Settings, or pick one of your endpoint's own " +
                    "models.");
            }

            return true;
        }

        public async Task<ChatCompletionResponseMessage> CompleteAsync(
            AppSettings settings,
            ChatCompletionRequest requestModel,
            CancellationToken cancellationToken)
        {
            await Scribble.Testing.TestLabStressBudget.GuardRequestAsync(settings, requestModel?.model, cancellationToken).ConfigureAwait(true);
            if (settings == null || !settings.IsConfigured)
            {
                throw new AiEndpointException(
                    "CONFIGURATION_INCOMPLETE",
                    "Open Settings and configure the endpoint, model, and API key.");
            }

            if (requestModel == null)
            {
                throw new ArgumentNullException(nameof(requestModel));
            }

            if (RoutesToGemini(settings, requestModel))
            {
                return await _gemini.GenerateAsync(
                    _httpClient,
                    settings,
                    requestModel,
                    cancellationToken).ConfigureAwait(true);
            }

            Uri endpoint;
            if (!AppSettings.TryGetChatCompletionsUri(
                settings.BaseUrl,
                settings.AllowInsecureHttp,
                out endpoint))
            {
                throw new AiEndpointException(
                    "ENDPOINT_INVALID",
                    "The configured endpoint is invalid.");
            }

            var capabilityKey = endpoint.AbsoluteUri + "\n" + requestModel.model;
            var openRouterQwenPolicy = UsesOpenRouterQwenPolicy(
                endpoint,
                requestModel.model);
            var hasOptionalToolControls =
                requestModel.temperature.HasValue ||
                requestModel.parallel_tool_calls.HasValue ||
                openRouterQwenPolicy;
            var includeOptionalToolControls = hasOptionalToolControls &&
                !OptionalToolControlsUnsupported(capabilityKey);
            try
            {
                return await CompleteOpenAiAsync(
                    settings,
                    endpoint,
                    requestModel,
                    includeOptionalToolControls,
                    cancellationToken).ConfigureAwait(true);
            }
            catch (AiEndpointException exception)
                when (includeOptionalToolControls &&
                      OptionalToolControlsRejected(exception))
            {
                MarkOptionalToolControlsUnsupported(capabilityKey);
                return await CompleteOpenAiAsync(
                    settings,
                    endpoint,
                    requestModel,
                    false,
                    cancellationToken).ConfigureAwait(true);
            }
        }

        private async Task<ChatCompletionResponseMessage>
            CompleteOpenAiAsync(
                AppSettings settings,
                Uri endpoint,
                ChatCompletionRequest requestModel,
                bool includeOptionalToolControls,
                CancellationToken cancellationToken,
                bool retryEmptyResponse = true,
                bool retryTransientResponse = true,
                string ignoredProvider = null,
                int providerRetriesRemaining = 2)
        {
            var circuitKey = endpoint.AbsoluteUri + "\n" + requestModel.model;
            lock (_optionalToolControlSync)
            {
                DateTime until;
                if (retryEmptyResponse && _emptyResponseCircuits.TryGetValue(circuitKey, out until) && until > DateTime.UtcNow)
                    throw new AiEndpointException("MODEL_CIRCUIT_OPEN", "This endpoint/model repeatedly returned empty completions. The task is retained. Wait 30 seconds or select another model before resuming.");
            }
            var payload = SerializablePayload(
                requestModel,
                endpoint,
                includeOptionalToolControls);
            ApplyTransientProviderExclusion(
                payload,
                endpoint,
                ignoredProvider);
            var requestJson = _serializer.Serialize(payload);
            requestModel.Diagnostics?.Record("inference_request", new { endpoint = endpoint.GetLeftPart(UriPartial.Path),
                model = requestModel.model, request = requestJson });

            using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
            using (var requestDeadline =
                CancellationTokenSource.CreateLinkedTokenSource(
                    cancellationToken))
            {
                // ResponseHeadersRead is intentionally used for bounded body
                // handling below, so HttpClient.Timeout alone would not cover
                // a provider that stalls before or during the response. Give
                // every inference attempt its own deadline while preserving
                // the caller's Stop/cancellation token.
                object requestedTokens;
                payload.TryGetValue("max_tokens", out requestedTokens);
                var completionRequestTimeout = CompletionDeadlineFor(
                    endpoint,
                    requestModel.model,
                    _completionRequestTimeout,
                    requestedTokens is int ? (int?)requestedTokens : null);
                requestDeadline.CancelAfter(completionRequestTimeout);
                request.Headers.Authorization =
                    new AuthenticationHeaderValue("Bearer", settings.ApiKey);
                request.Headers.Accept.Add(
                    new MediaTypeWithQualityHeaderValue("application/json"));
                request.Content = new StringContent(
                    requestJson,
                    Encoding.UTF8,
                    "application/json");

                HttpResponseMessage response;
                try
                {
                    response = await _httpClient
                        .SendAsync(
                            request,
                            HttpCompletionOption.ResponseHeadersRead,
                            requestDeadline.Token)
                        .ConfigureAwait(true);
                }
                catch (OperationCanceledException exception)
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        throw;
                    }

                    throw new AiEndpointException(
                        "AI_TIMEOUT",
                        "The AI endpoint did not respond before the request was cancelled.",
                        exception);
                }
                catch (HttpRequestException exception)
                {
                    throw CreateNetworkException(
                        exception,
                        endpoint);
                }

                using (response)
                {
                    Scribble.Testing.TestLabStressBudget.RecordProviderResponse((int)response.StatusCode);
                    string responseText;
                    try
                    {
                        responseText = await ReadBoundedAsync(
                            response.Content,
                            requestDeadline.Token).ConfigureAwait(true);
                    }
                    catch (AiEndpointException)
                    {
                        throw;
                    }
                    catch (OperationCanceledException exception)
                    {
                        if (cancellationToken.IsCancellationRequested)
                        {
                            throw;
                        }

                        throw new AiEndpointException(
                            "AI_TIMEOUT",
                            "The AI endpoint did not complete the response " +
                            "within " + FormatTimeout(completionRequestTimeout) +
                            ". No partial tool action ran; " +
                            "the task is preserved and can be resumed.",
                            exception);
                    }
                    catch (Exception exception)
                    {
                        throw new AiEndpointException(
                            "RESPONSE_READ_FAILED",
                            "The AI endpoint response could not be read.",
                            exception);
                    }

                    var requestId = GetRequestId(response);
                    requestModel.Diagnostics?.Record("inference_response", new { endpoint = endpoint.GetLeftPart(UriPartial.Path),
                        model = requestModel.model, http_status = (int)response.StatusCode, request_id = requestId,
                        server = response.Headers.Server.ToString(), response = responseText,
                        request_hash = TaskCheckpointStore.Fingerprint(requestJson) });
                    if (!response.IsSuccessStatusCode)
                    {
                        var error = TryReadError(responseText);
                        var status = (int)response.StatusCode;
                        if (retryTransientResponse && (status == 429 || status == 502 || status == 503 || status == 504))
                        {
                            var hint = response.Headers.RetryAfter;
                            var retryAfter = hint?.Delta ?? (hint?.Date.HasValue == true ? hint.Date.Value - DateTimeOffset.UtcNow : TimeSpan.FromSeconds(1));
                            if (retryAfter >= TimeSpan.Zero && retryAfter <= TimeSpan.FromSeconds(2))
                            {
                                await Scribble.Testing.TestLabStressBudget
                                    .GuardRequestAsync(
                                        settings,
                                        requestModel?.model,
                                        cancellationToken).ConfigureAwait(true);
                                await Task.Delay(retryAfter, cancellationToken).ConfigureAwait(true);
                                return await CompleteOpenAiAsync(
                                    settings,
                                    endpoint,
                                    requestModel,
                                    includeOptionalToolControls,
                                    cancellationToken,
                                    retryEmptyResponse,
                                    false,
                                    ignoredProvider,
                                    providerRetriesRemaining).ConfigureAwait(true);
                            }
                        }
                        var reason = string.IsNullOrWhiteSpace(response.ReasonPhrase)
                            ? response.StatusCode.ToString()
                            : response.ReasonPhrase;
                        throw new AiEndpointException(
                            BuildHttpCode(status, reason),
                            "The AI endpoint rejected the request: " +
                            status + " " + reason + "." +
                            RecoveryHint(status),
                            httpStatus: status,
                            providerCode: error?.code ?? error?.type,
                            requestId: requestId,
                            responseSnippet: error?.message ?? responseText);
                    }

                    ChatCompletionResponse completion;
                    try
                    {
                        completion =
                            _serializer.Deserialize<ChatCompletionResponse>(
                                responseText);
                    }
                    catch (Exception exception)
                    {
                        throw new AiEndpointException(
                            "RESPONSE_INVALID_JSON",
                            "The AI endpoint returned invalid JSON.",
                            exception,
                            (int)response.StatusCode,
                            requestId: requestId,
                            responseSnippet: responseText);
                    }

                    var choice =
                        completion?.choices != null &&
                        completion.choices.Count > 0
                            ? completion.choices[0]
                            : null;
                    if (choice?.error != null ||
                        string.Equals(
                            choice?.finish_reason,
                            "error",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        // OpenRouter can return HTTP 200 with a provider-side
                        // 5xx embedded in the first choice. Its message may
                        // contain partial text or a truncated tool call, none
                        // of which is safe to execute. Retry the identical
                        // inference once on generic endpoints. OpenRouter can
                        // safely route across up to two other providers before
                        // surfacing a resumable failure. Every retry is still
                        // checked by the Test Lab's hard spend guard.
                        var excludedProviders = SplitProviders(ignoredProvider);
                        if (!string.IsNullOrWhiteSpace(completion?.provider) &&
                            !excludedProviders.Contains(completion.provider, StringComparer.OrdinalIgnoreCase))
                            excludedProviders.Add(completion.provider);
                        var openRouter = endpoint != null && string.Equals(
                            endpoint.Host,
                            "openrouter.ai",
                            StringComparison.OrdinalIgnoreCase);
                        if (retryTransientResponse ||
                            (openRouter && providerRetriesRemaining > 0))
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            await Scribble.Testing.TestLabStressBudget
                                .GuardRequestAsync(
                                    settings,
                                    requestModel?.model,
                                    cancellationToken).ConfigureAwait(true);
                            await Task.Delay(
                                TimeSpan.FromSeconds(1),
                                cancellationToken).ConfigureAwait(true);
                            return await CompleteOpenAiAsync(
                                settings,
                                endpoint,
                                requestModel,
                                includeOptionalToolControls,
                                cancellationToken,
                                retryEmptyResponse,
                                false,
                                string.Join("\n", excludedProviders),
                                providerRetriesRemaining - 1).ConfigureAwait(true);
                        }

                        var providerError = choice?.error;
                        throw new AiEndpointException(
                            "PROVIDER_RESPONSE_ERROR",
                            "The selected AI provider interrupted the response. " +
                            "No partial tool action ran. The task is preserved; " +
                            "resume or choose another provider.",
                            httpStatus: (int)response.StatusCode,
                            providerCode:
                                providerError?.code ?? providerError?.type,
                            requestId: requestId,
                            responseSnippet:
                                providerError?.message ?? responseText);
                    }

                    var message = choice?.message;

                    var hasToolCalls =
                        message?.tool_calls != null &&
                        message.tool_calls.Count > 0;
                    if (message == null ||
                        (!hasToolCalls &&
                         string.IsNullOrWhiteSpace(message.content)))
                    {
                        // No assistant message or tool action was delivered, so
                        // retry this inference once without replaying any tools.
                        // Persistent empty responses remain a resumable failure.
                        if (retryEmptyResponse)
                        {
                            var excludedProviders =
                                SplitProviders(ignoredProvider);
                            var openRouter = endpoint != null &&
                                string.Equals(
                                    endpoint.Host,
                                    "openrouter.ai",
                                    StringComparison.OrdinalIgnoreCase);
                            if (openRouter &&
                                !string.IsNullOrWhiteSpace(
                                    completion?.provider) &&
                                !excludedProviders.Contains(
                                    completion.provider,
                                    StringComparer.OrdinalIgnoreCase))
                            {
                                excludedProviders.Add(completion.provider);
                            }
                            cancellationToken.ThrowIfCancellationRequested();
                            await Scribble.Testing.TestLabStressBudget
                                .GuardRequestAsync(
                                    settings,
                                    requestModel?.model,
                                    cancellationToken).ConfigureAwait(true);
                            return await CompleteOpenAiAsync(
                                settings,
                                endpoint,
                                requestModel,
                                includeOptionalToolControls,
                                cancellationToken,
                                false,
                                retryTransientResponse,
                                string.Join("\n", excludedProviders),
                                providerRetriesRemaining).ConfigureAwait(true);
                        }
                        lock (_optionalToolControlSync) _emptyResponseCircuits[circuitKey] = DateTime.UtcNow.AddSeconds(30);
                        throw new AiEndpointException(
                            "RESPONSE_MISSING_CONTENT",
                            "The AI endpoint returned an empty response twice. No new tool actions ran. The task is preserved; resume or choose another model.",
                            httpStatus: (int)response.StatusCode,
                            requestId: requestId,
                            responseSnippet: responseText);
                    }

                    message.RawContent = TextBoundary.PlainText(message.content, TextBoundary.MaxHttpResponseCharacters);
                    lock (_optionalToolControlSync) _emptyResponseCircuits.Remove(circuitKey);
                    message.content = TextBoundary.PlainText(
                        message.content,
                        TextBoundary.MaxAssistantCharacters);
                    NormalizeToolCalls(message.tool_calls);
                    return message;
                }
            }
        }

        private static void ApplyTransientProviderExclusion(
            Dictionary<string, object> payload,
            Uri endpoint,
            string provider)
        {
            if (payload == null ||
                endpoint == null ||
                !string.Equals(
                    endpoint.Host,
                    "openrouter.ai",
                    StringComparison.OrdinalIgnoreCase) ||
                string.IsNullOrWhiteSpace(provider))
            {
                return;
            }

            var ignored = SplitProviders(provider).ToArray();
            if (ignored.Length == 0) return;

            Dictionary<string, object> preferences;
            object existing;
            if (payload.TryGetValue("provider", out existing))
            {
                preferences = existing as Dictionary<string, object>;
            }
            else
            {
                preferences = null;
            }
            if (preferences == null)
            {
                preferences = new Dictionary<string, object>();
                payload["provider"] = preferences;
            }
            preferences["ignore"] = ignored;
        }

        private static List<string> SplitProviders(string providers)
        {
            return (providers ?? "").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(value => value.Trim()).Where(value => value.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        // The Gemini tick only decides whether Google models are
        // OFFERED. With it on, the picker lists the Gemini models
        // AND the endpoint's own models together, so switching to a
        // local model is a picker choice rather than a settings
        // change.
        public async Task<IReadOnlyList<string>> GetModelsAsync(
            AppSettings settings,
            CancellationToken cancellationToken)
        {
            if (settings == null)
            {
                throw new ArgumentNullException(nameof(settings));
            }

            if (AdminPolicy.GeminiDisabled ||
                !settings.UseGeminiSignIn)
            {
                return await FetchEndpointModelsAsync(
                    settings,
                    cancellationToken).ConfigureAwait(true);
            }

            var offered = new List<string>();
            Exception geminiFailure = null;
            try
            {
                offered.AddRange(
                    await _gemini.VerifySignInAsync(
                        _httpClient,
                        settings,
                        cancellationToken).ConfigureAwait(true));
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception)
            {
                // Not signed in yet, or Google is unreachable: the
                // endpoint's own models must still be offered.
                geminiFailure = exception;
            }

            if (settings.HasEndpointCredentials)
            {
                try
                {
                    offered.AddRange(
                        await FetchEndpointModelsAsync(
                            settings,
                            cancellationToken)
                            .ConfigureAwait(true));
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception)
                {
                    // An unreachable local endpoint must never hide
                    // the Gemini models that do work.
                }
            }

            if (offered.Count == 0 && geminiFailure != null)
            {
                throw geminiFailure;
            }

            return offered
                .Where(item =>
                    !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private async Task<IReadOnlyList<string>>
            FetchEndpointModelsAsync(
                AppSettings settings,
                CancellationToken cancellationToken)
        {
            Uri endpoint;
            if (!AppSettings.TryGetModelsUri(
                settings.BaseUrl,
                settings.AllowInsecureHttp,
                out endpoint))
            {
                throw new AiEndpointException(
                    "ENDPOINT_INVALID",
                    "The configured endpoint is invalid.");
            }

            if (string.IsNullOrWhiteSpace(settings.ApiKey))
            {
                throw new AiEndpointException(
                    "CONFIGURATION_INCOMPLETE",
                    "Enter an API key before checking the endpoint.");
            }

            using (var request =
                new HttpRequestMessage(HttpMethod.Get, endpoint))
            {
                request.Headers.Authorization =
                    new AuthenticationHeaderValue(
                        "Bearer",
                        settings.ApiKey);
                request.Headers.Accept.Add(
                    new MediaTypeWithQualityHeaderValue(
                        "application/json"));

                HttpResponseMessage response;
                try
                {
                    response = await _httpClient
                        .SendAsync(
                            request,
                            HttpCompletionOption.ResponseHeadersRead,
                            cancellationToken)
                        .ConfigureAwait(true);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (HttpRequestException exception)
                {
                    throw CreateNetworkException(
                        exception,
                        endpoint);
                }

                using (response)
                {
                    var responseText = await ReadBoundedAsync(
                        response.Content,
                        cancellationToken).ConfigureAwait(true);
                    var requestId = GetRequestId(response);

                    if (!response.IsSuccessStatusCode)
                    {
                        var error = TryReadError(responseText);
                        var status = (int)response.StatusCode;
                        var reason =
                            string.IsNullOrWhiteSpace(
                                response.ReasonPhrase)
                                ? response.StatusCode.ToString()
                                : response.ReasonPhrase;
                        throw new AiEndpointException(
                            BuildHttpCode(status, reason),
                            "The endpoint model list request failed: " +
                            status + " " + reason + "." +
                            RecoveryHint(status),
                            httpStatus: status,
                            providerCode:
                                error?.code ?? error?.type,
                            requestId: requestId,
                            responseSnippet:
                                error?.message ?? responseText);
                    }

                    ModelListResponse list;
                    try
                    {
                        list = _serializer
                            .Deserialize<ModelListResponse>(
                                responseText);
                    }
                    catch (Exception exception)
                    {
                        throw new AiEndpointException(
                            "MODELS_INVALID_JSON",
                            "The endpoint returned an invalid model list.",
                            exception,
                            (int)response.StatusCode,
                            requestId: requestId,
                            responseSnippet: responseText);
                    }

                    var models = (list?.data ??
                        new List<ModelListItem>())
                        .Where(item =>
                            item != null &&
                            ModelSelectionPolicy
                                .IsGenerativeModel(item.id))
                        .Select(item =>
                            TextBoundary.PlainText(item.id, 200))
                        .Where(item => item.Length > 0)
                        .Distinct(
                            StringComparer.OrdinalIgnoreCase)
                        .OrderBy(item => item)
                        .ToList();
                    if (models.Count == 0)
                    {
                        throw new AiEndpointException(
                            "MODELS_EMPTY",
                            "The endpoint returned no generative model identifiers.",
                            httpStatus:
                                (int)response.StatusCode,
                            requestId: requestId,
                            responseSnippet: responseText);
                    }

                    return models;
                }
            }
        }

        public void Dispose()
        {
            _httpClient.Dispose();
        }

        private static async Task<string> ReadBoundedAsync(
            HttpContent content,
            CancellationToken cancellationToken)
        {
            if (content.Headers.ContentLength.HasValue &&
                content.Headers.ContentLength.Value >
                TextBoundary.MaxHttpResponseCharacters)
            {
                throw new AiEndpointException(
                    "RESPONSE_TOO_LARGE",
                    "The AI endpoint response was too large.");
            }

            using (var stream = await content.ReadAsStreamAsync()
                .ConfigureAwait(true))
            // On .NET Framework ReadAsync takes no token, so a read
            // blocked on a stalled connection would ignore Stop.
            // Closing the stream on cancellation faults the pending
            // read, which is rethrown below as a cancellation.
            using (cancellationToken.Register(stream.Close))
            using (var reader = new StreamReader(
                stream,
                Encoding.UTF8,
                true,
                4096,
                false))
            {
                var builder = new StringBuilder();
                var buffer = new char[4096];
                while (true)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int read;
                    try
                    {
                        read = await reader.ReadAsync(
                            buffer,
                            0,
                            buffer.Length).ConfigureAwait(true);
                    }
                    catch (Exception exception)
                        when (cancellationToken
                            .IsCancellationRequested)
                    {
                        throw new OperationCanceledException(
                            "The response read was cancelled.",
                            exception,
                            cancellationToken);
                    }

                    if (read == 0)
                    {
                        break;
                    }

                    if (builder.Length + read >
                        TextBoundary.MaxHttpResponseCharacters)
                    {
                        throw new AiEndpointException(
                            "RESPONSE_TOO_LARGE",
                            "The AI endpoint response was too large.");
                    }

                    builder.Append(buffer, 0, read);
                }

                return builder.ToString();
            }
        }

        private static void NormalizeToolCalls(
            IList<ChatToolCall> toolCalls)
        {
            if (toolCalls == null)
            {
                return;
            }

            for (var index = 0;
                 index < toolCalls.Count;
                 index++)
            {
                var call = toolCalls[index];
                if (call == null)
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(call.id))
                {
                    call.id = "call_" +
                        (index + 1).ToString();
                }

                if (string.IsNullOrWhiteSpace(call.type))
                {
                    call.type = "function";
                }
            }
        }

        private static AiEndpointException CreateNetworkException(
            HttpRequestException exception,
            Uri endpoint)
        {
            var webException = FindWebException(exception);
            var code = "NETWORK_REQUEST_FAILED";
            if (webException != null)
            {
                switch (webException.Status)
                {
                    case WebExceptionStatus.NameResolutionFailure:
                        code = "NETWORK_NAME_RESOLUTION";
                        break;
                    case WebExceptionStatus.ConnectFailure:
                        code = "NETWORK_CONNECT_FAILURE";
                        break;
                    case WebExceptionStatus.TrustFailure:
                    case WebExceptionStatus.SecureChannelFailure:
                        code = "TLS_SECURE_CHANNEL_FAILURE";
                        break;
                    case WebExceptionStatus.ProxyNameResolutionFailure:
                        code = "NETWORK_PROXY_NAME_RESOLUTION";
                        break;
                    case WebExceptionStatus.Timeout:
                        code = "AI_TIMEOUT";
                        break;
                    default:
                        code = "NETWORK_" +
                            webException.Status.ToString().ToUpperInvariant();
                        break;
                }
            }

            return new AiEndpointException(
                code,
                "The AI endpoint could not be reached. Check its URL, " +
                "network access, TLS certificate, and whether the local server is running.",
                exception,
                transportDetails: BuildTransportDetails(
                    exception,
                    endpoint));
        }

        private static string BuildTransportDetails(
            Exception exception,
            Uri endpoint)
        {
            var details = new List<string>();
            if (endpoint != null)
            {
                details.Add(
                    "Target " +
                    endpoint.Scheme +
                    "://" +
                    endpoint.Host +
                    ":" +
                    endpoint.Port);
            }

            var current = exception;
            var depth = 0;
            while (current != null && depth < 8)
            {
                var item =
                    current.GetType().Name +
                    " HRESULT 0x" +
                    current.HResult.ToString("X8");

                var webException = current as WebException;
                if (webException != null)
                {
                    item +=
                        " WebExceptionStatus " +
                        webException.Status;
                }

                var socketException = current as SocketException;
                if (socketException != null)
                {
                    item +=
                        " SocketError " +
                        socketException.SocketErrorCode +
                        " NativeError " +
                        socketException.NativeErrorCode;
                }

                var message = TextBoundary.PlainText(
                    current.Message,
                    600)
                    .Replace("\r", " ")
                    .Replace("\n", " ")
                    .Trim();
                if (message.Length > 0)
                {
                    item += ": " + message;
                }

                details.Add(item);
                current = current.InnerException;
                depth++;
            }

            return string.Join(" | ", details);
        }

        private ChatCompletionError TryReadError(string responseText)
        {
            try
            {
                return _serializer
                    .Deserialize<ChatCompletionResponse>(responseText)
                    ?.error;
            }
            catch
            {
                return null;
            }
        }

        private static WebException FindWebException(Exception exception)
        {
            var current = exception;
            while (current != null)
            {
                var webException = current as WebException;
                if (webException != null)
                {
                    return webException;
                }

                current = current.InnerException;
            }

            return null;
        }

        private static string GetRequestId(HttpResponseMessage response)
        {
            foreach (var name in new[]
            {
                "x-request-id",
                "request-id",
                "x-correlation-id",
                "traceparent"
            })
            {
                IEnumerable<string> values;
                if (response.Headers.TryGetValues(name, out values))
                {
                    foreach (var value in values)
                    {
                        return TextBoundary.PlainText(value, 200);
                    }
                }
            }

            return string.Empty;
        }

        // Strict OpenAI-compatible endpoints reject requests that
        // carry "tools": [] or null tool fields with 400, so the
        // optional fields are included only when they carry a value.
        private static Dictionary<string, object> SerializablePayload(
            ChatCompletionRequest requestModel,
            Uri endpoint,
            bool includeOptionalToolControls = true)
        {
            var payload = new Dictionary<string, object>
            {
                { "model", requestModel.model },
                { "messages", requestModel.messages },
                { "stream", requestModel.stream }
            };
            if (requestModel.tools != null &&
                requestModel.tools.Count > 0)
            {
                payload["tools"] = requestModel.tools;
                if (requestModel.tool_choice != null)
                {
                    payload["tool_choice"] = requestModel.tool_choice;
                }
            }

            if (requestModel.max_tokens.HasValue)
            {
                payload["max_tokens"] = requestModel.max_tokens.Value;
            }

            if (includeOptionalToolControls &&
                requestModel.temperature.HasValue)
            {
                payload["temperature"] = requestModel.temperature.Value;
            }

            if (includeOptionalToolControls &&
                requestModel.parallel_tool_calls.HasValue)
            {
                payload["parallel_tool_calls"] =
                    requestModel.parallel_tool_calls.Value;
            }

            // OpenRouter exposes reasoning as provider metadata rather than
            // part of Scribble's endpoint-neutral request contract. Qwen 3.8
            // defaults to xhigh reasoning there, which can consume the entire
            // response allowance before a tool call or answer is emitted.
            // Office tools also operate on one COM apartment and must not be
            // dispatched in parallel. Keep both overrides narrowly bound to
            // the exact stress-suite endpoint/model pair.
            if (UsesOpenRouterQwenPolicy(endpoint, requestModel.model))
            {
                // Dense Office authoring calls carry native table/chart JSON.
                // Qwen can otherwise truncate a syntactically valid tool call at
                // the generic 4K draft ceiling and spend more on retries. A full
                // six-to-eight-slide payload can exceed 8K, while the model route
                // supports 32K completions. Keep that ceiling exclusive to the
                // PowerPoint draft tool; other draft calls get 8K and compact
                // reviewers/summarizers retain their original smaller limits.
                var isDraftRequest = requestModel.max_tokens ==
                    DocumentChatRequestFactory.DraftResponseTokens;
                if (isDraftRequest)
                {
                    var hasPresentationDraftTool = requestModel.tools != null &&
                        requestModel.tools.Any(tool => tool?.function != null &&
                            (string.Equals(tool.function.name,
                                 PresentationToolCatalog.AddDraftSlides,
                                 StringComparison.Ordinal) ||
                             string.Equals(tool.function.name,
                                 CrossAppToolCatalog.SendToPowerPoint,
                                 StringComparison.Ordinal)));
                    payload["max_tokens"] = hasPresentationDraftTool
                        ? 32768
                        : 8192;
                }
                // Compact reviewers and summarizers need a verdict, not a
                // hidden chain of thought. Some OpenRouter providers have
                // spent the entire 2K response allowance on reasoning,
                // returning no content. Disable reasoning for those bounded
                // internal calls. Qwen 3.8 advertises low as its smallest
                // supported reasoning effort; sending the unsupported minimal
                // value can fall back to the model's xhigh default and consume
                // the entire response allowance before a tool call is emitted.
                // Keep low reasoning on normal task turns so Qwen can still
                // reconcile source material while leaving room for complete
                // tool-call JSON.
                var compactInternalCall =
                    (requestModel.tools == null ||
                     requestModel.tools.Count == 0) &&
                    requestModel.max_tokens.HasValue &&
                    requestModel.max_tokens.Value <= 2048;
                payload["reasoning"] = new Dictionary<string, object>
                {
                    {
                        "effort",
                        compactInternalCall ? "none" : "low"
                    }
                };
                if (includeOptionalToolControls &&
                    requestModel.tools != null &&
                    requestModel.tools.Count > 0)
                {
                    payload["parallel_tool_calls"] = false;

                    // OpenRouter's default price-weighted routing can select
                    // endpoints that advertise generic tool support but do not
                    // reliably honor Qwen's bounded reasoning/tool-choice
                    // contract. Keep tool-bearing requests on the endpoints
                    // observed to support the required tool parameters,
                    // ordered by successful Scribble tool-turn latency.
                    // OpenRouter's endpoint metadata does
                    // not advertise the optional parallel_tool_calls switch for
                    // any Qwen 3.8 route, so require_parameters cannot be used
                    // even though false is the serial-safe value we need. The
                    // allow-list prevents an outside fallback from reintroducing
                    // the same empty/timeout failure mode.
                    var reliableToolProviders = new[]
                    {
                        "reka",
                        "mancer",
                        "phala",
                        "coreweave",
                        "dekallm",
                        "chutes"
                    };
                    // A long authoring completion is bound by generation
                    // speed, not first-token latency: an 11 token/s route
                    // needs eight minutes for a deck payload that an
                    // 80 token/s route returns in one. Keep the allow-list
                    // but let OpenRouter choose its fastest member there.
                    object authoringTokens;
                    var longAuthoringCall =
                        payload.TryGetValue("max_tokens", out authoringTokens) &&
                        authoringTokens is int &&
                        (int)authoringTokens >= 8192;
                    payload["provider"] = longAuthoringCall
                        ? new Dictionary<string, object>
                        {
                            { "only", reliableToolProviders },
                            { "sort", "throughput" },
                            { "allow_fallbacks", true }
                        }
                        : new Dictionary<string, object>
                        {
                            { "order", reliableToolProviders },
                            { "only", reliableToolProviders },
                            { "allow_fallbacks", true }
                        };
                }
            }

            return payload;
        }

        private static bool UsesOpenRouterQwenPolicy(
            Uri endpoint,
            string model)
        {
            return endpoint != null &&
                string.Equals(
                    endpoint.Scheme,
                    Uri.UriSchemeHttps,
                    StringComparison.OrdinalIgnoreCase) &&
                string.Equals(
                    endpoint.Host,
                    "openrouter.ai",
                    StringComparison.OrdinalIgnoreCase) &&
                endpoint.IsDefaultPort &&
                string.IsNullOrEmpty(endpoint.UserInfo) &&
                string.IsNullOrEmpty(endpoint.Query) &&
                string.Equals(
                    endpoint.AbsolutePath.TrimEnd('/'),
                    "/api/v1/chat/completions",
                    StringComparison.Ordinal) &&
                string.Equals(
                    model,
                    "qwen/qwen3.8-27b",
                    StringComparison.Ordinal);
        }

        private static TimeSpan CompletionRequestTimeoutFor(
            Uri endpoint,
            string model,
            TimeSpan defaultTimeout)
        {
            return UsesOpenRouterQwenPolicy(endpoint, model) &&
                defaultTimeout == CompletionRequestTimeout
                    ? OpenRouterQwenCompletionRequestTimeout
                    : defaultTimeout;
        }

        internal static readonly TimeSpan MaximumCompletionRequestTimeout =
            TimeSpan.FromMinutes(15);

        // This client buffers the whole completion, so its deadline bounds
        // generation time, not idle time. A multi-slide tool call is several
        // thousand output tokens: a local model at 10-25 tokens per second
        // cannot finish that inside the chat-sized window. Extend the window
        // by the output the request itself allows (10 tokens per second),
        // within a fixed ceiling. An injected test timeout is never widened.
        private static TimeSpan CompletionDeadlineFor(
            Uri endpoint,
            string model,
            TimeSpan defaultTimeout,
            int? maxTokens)
        {
            var window = CompletionRequestTimeoutFor(
                endpoint,
                model,
                defaultTimeout);
            if (defaultTimeout != CompletionRequestTimeout ||
                !maxTokens.HasValue ||
                maxTokens.Value <= 0)
            {
                return window;
            }

            var scaled = window +
                TimeSpan.FromSeconds(maxTokens.Value / 10d);
            return scaled > MaximumCompletionRequestTimeout
                ? MaximumCompletionRequestTimeout
                : scaled;
        }

        private static string FormatTimeout(TimeSpan timeout)
        {
            if (timeout.TotalMinutes == Math.Floor(timeout.TotalMinutes))
            {
                var minutes = (int)timeout.TotalMinutes;
                return minutes.ToString() +
                    (minutes == 1 ? " minute" : " minutes");
            }

            return Math.Ceiling(timeout.TotalSeconds).ToString() + " seconds";
        }

        private bool OptionalToolControlsUnsupported(string capabilityKey)
        {
            lock (_optionalToolControlSync)
            {
                return _optionalToolControlUnsupported.Contains(
                    capabilityKey ?? string.Empty);
            }
        }

        private static bool OptionalToolControlsRejected(
            AiEndpointException exception)
        {
            if (exception == null || exception.HttpStatus != 400)
            {
                return false;
            }

            var detail = (
                (exception.ProviderCode ?? string.Empty) + " " +
                (exception.ResponseSnippet ?? string.Empty) + " " +
                (exception.Message ?? string.Empty))
                .ToLowerInvariant();
            return detail.Contains("temperature") ||
                detail.Contains("parallel_tool_calls");
        }

        private void MarkOptionalToolControlsUnsupported(string capabilityKey)
        {
            lock (_optionalToolControlSync)
            {
                _optionalToolControlUnsupported.Add(
                    capabilityKey ?? string.Empty);
            }
        }

        private static string BuildHttpCode(int status, string reason)
        {
            var builder = new StringBuilder();
            foreach (var character in reason ?? string.Empty)
            {
                if (char.IsLetterOrDigit(character))
                {
                    builder.Append(char.ToUpperInvariant(character));
                }
                else if (builder.Length > 0 &&
                         builder[builder.Length - 1] != '_')
                {
                    builder.Append('_');
                }
            }

            return
                "HTTP_" +
                status +
                "_" +
                builder.ToString().Trim('_');
        }

        private static string RecoveryHint(int status)
        {
            switch (status)
            {
                case 400:
                    return
                        " The endpoint may not support OpenAI-compatible tool calling, " +
                        "or the model name/request format is invalid.";
                case 401:
                    return " Verify the API key.";
                case 403:
                    return " Verify the API key permissions and endpoint policy.";
                case 404:
                    return " Verify the base URL and model name.";
                case 408:
                    return " Retry after checking endpoint load.";
                case 413:
                    return " The endpoint rejected the bounded mailbox context size.";
                case 429:
                    return " The endpoint is rate-limiting requests. Retry later.";
                default:
                    return status >= 500
                        ? " The endpoint failed internally. Check its server logs."
                        : string.Empty;
            }
        }
    }
}
