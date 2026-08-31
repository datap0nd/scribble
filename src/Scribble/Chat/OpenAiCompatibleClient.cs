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

        public OpenAiCompatibleClient()
        {
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

            var requestJson = _serializer.Serialize(
                SerializablePayload(requestModel));

            using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
            {
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
                            cancellationToken)
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
                    string responseText;
                    try
                    {
                        responseText = await ReadBoundedAsync(
                            response.Content,
                            cancellationToken).ConfigureAwait(true);
                    }
                    catch (AiEndpointException)
                    {
                        throw;
                    }
                    catch (Exception exception)
                    {
                        throw new AiEndpointException(
                            "RESPONSE_READ_FAILED",
                            "The AI endpoint response could not be read.",
                            exception);
                    }

                    var requestId = GetRequestId(response);
                    if (!response.IsSuccessStatusCode)
                    {
                        var error = TryReadError(responseText);
                        var status = (int)response.StatusCode;
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

                    var message =
                        completion?.choices != null &&
                        completion.choices.Count > 0
                            ? completion.choices[0]?.message
                            : null;

                    var hasToolCalls =
                        message?.tool_calls != null &&
                        message.tool_calls.Count > 0;
                    if (message == null ||
                        (!hasToolCalls &&
                         string.IsNullOrWhiteSpace(message.content)))
                    {
                        throw new AiEndpointException(
                            "RESPONSE_MISSING_CONTENT",
                            "The AI endpoint returned neither message text nor tool calls.",
                            httpStatus: (int)response.StatusCode,
                            requestId: requestId,
                            responseSnippet: responseText);
                    }

                    message.content = TextBoundary.PlainText(
                        message.content,
                        TextBoundary.MaxAssistantCharacters);
                    NormalizeToolCalls(message.tool_calls);
                    return message;
                }
            }
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
            ChatCompletionRequest requestModel)
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

            return payload;
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
