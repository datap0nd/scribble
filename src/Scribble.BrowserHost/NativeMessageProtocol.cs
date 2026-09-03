using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Security;

namespace Scribble.BrowserHost
{
    internal sealed class BrowserHistoryTurn
    {
        public string role { get; set; }

        public string content { get; set; }
    }

    internal sealed class BrowserPageContext
    {
        public string title { get; set; }

        public string url { get; set; }

        public string selection { get; set; }

        public string pageText { get; set; }

        public string links { get; set; }

        public string screenshotDataUrl { get; set; }
    }

    internal sealed class BrowserNativeToolCall
    {
        public string id { get; set; }

        public string name { get; set; }

        public string arguments { get; set; }
    }

    internal sealed class BrowserNativeToolResult
    {
        public string id { get; set; }

        public string content { get; set; }
    }

    internal sealed class BrowserNativeExchangeTurn
    {
        public string assistantContent { get; set; }

        public List<BrowserNativeToolCall> toolCalls { get; set; }

        public List<BrowserNativeToolResult> results { get; set; }
    }

    internal sealed class BrowserNativeActionDescriptor
    {
        public string action { get; set; }

        public string tagName { get; set; }

        public string inputType { get; set; }

        public string role { get; set; }

        public string name { get; set; }

        public string placeholder { get; set; }

        public string autocomplete { get; set; }

        public string url { get; set; }

        public string value { get; set; }

        public string sourceText { get; set; }

        public string key { get; set; }

        public bool formHasPassword { get; set; }

        public bool formHasPayment { get; set; }

        public bool formHasPersonalData { get; set; }
    }

    internal sealed class BrowserNativeRequest
    {
        public string type { get; set; }

        public string requestId { get; set; }

        public string prompt { get; set; }

        public List<BrowserHistoryTurn> history { get; set; }

        public BrowserPageContext context { get; set; }

        public List<BrowserNativeExchangeTurn> exchange { get; set; }

        public string chatId { get; set; }

        public string turnId { get; set; }

        public string topicId { get; set; }

        public string topicBinding { get; set; }

        public BrowserNativeActionDescriptor action { get; set; }
    }

    internal sealed class BrowserNativeTopic
    {
        public string id { get; set; }

        public string name { get; set; }

        public string binding { get; set; }

        public bool available { get; set; }
    }

    internal sealed class BrowserNativeResponse
    {
        public bool ok { get; set; }

        public string content { get; set; }

        public string error { get; set; }

        public string errorCode { get; set; }

        public string model { get; set; }

        public bool configured { get; set; }

        public bool supportsVision { get; set; }

        public string version { get; set; }

        public string availableExtensionVersion { get; set; }

        public List<BrowserNativeTopic> topics { get; set; }

        // When set, the panel must execute these browser tool calls
        // (navigation, page reads) and continue the chat with an
        // exchange entry carrying assistantContent, toolRequests,
        // and the merged hostResults + browser results.
        public List<BrowserNativeToolCall> toolRequests { get; set; }

        public string assistantContent { get; set; }

        public List<BrowserNativeToolResult> hostResults { get; set; }

        public bool actionAllowed { get; set; }

        public string actionCode { get; set; }
    }

    internal static class NativeMessageProtocol
    {
        internal const int MaxRequestBytes = 16 * 1024 * 1024;
        internal const int MaxResponseBytes = 900 * 1024;
        internal const int MaxHistoryTurns = 12;
        internal const int MaxExchangeResultCharacters = 12 * 1024;
        internal const int MaxExchangeCharacters = 320 * 1024;

        private static readonly Encoding Utf8 =
            new UTF8Encoding(false, true);

        internal static int Run(Stream input, Stream output)
        {
            if (input == null)
            {
                throw new ArgumentNullException(nameof(input));
            }

            if (output == null)
            {
                throw new ArgumentNullException(nameof(output));
            }

            string json;
            try
            {
                if (!TryRead(input, out json))
                {
                    return 0;
                }

                Write(output, Handle(json));
                return 0;
            }
            catch (Exception exception)
            {
                try
                {
                    Write(
                        output,
                        Error(
                            string.Empty,
                            CodeFor(exception),
                            SafeMessage(exception)));
                    return 0;
                }
                catch (Exception writeException)
                {
                    Console.Error.WriteLine(
                        "Scribble native messaging failed: " +
                        writeException.GetType().Name);
                    return 1;
                }
            }
        }

        internal static BrowserNativeResponse Handle(string json)
        {
            BrowserNativeRequest request;
            try
            {
                var serializer = Serializer();
                request = serializer.Deserialize<BrowserNativeRequest>(
                    json ?? string.Empty);
            }
            catch (Exception exception)
            {
                return Error(
                    string.Empty,
                    "REQUEST_INVALID",
                    "The browser sent invalid request data: " +
                    exception.GetType().Name + ".");
            }

            if (request == null)
            {
                return Error(
                    string.Empty,
                    "REQUEST_INVALID",
                    "The browser request was empty.");
            }

            var requestId = TextBoundary.SingleLine(
                request.requestId,
                100);
            BrowserChatService service = null;
            try
            {
                service = new BrowserChatService();
                if (string.Equals(
                    request.type,
                    "ping",
                    StringComparison.Ordinal))
                {
                    return Success(
                        service,
                        requestId,
                        string.Empty,
                        service.Model,
                        false);
                }

                if (string.Equals(
                    request.type,
                    "openSettings",
                    StringComparison.Ordinal))
                {
                    service.Dispose();
                    service = null;
                    SettingsLauncher.ShowSettingsDialog();
                    service = new BrowserChatService();
                    return Success(
                        service,
                        requestId,
                        string.Empty,
                        service.Model,
                        false);
                }

                if (string.Equals(
                    request.type,
                    "clearSession",
                    StringComparison.Ordinal))
                {
                    TopicToolHost.ClearPersistentChat(
                        request.chatId);
                    return Success(
                        service,
                        requestId,
                        string.Empty,
                        service.Model,
                        false);
                }

                if (string.Equals(
                    request.type,
                    "authorizeBrowserAction",
                    StringComparison.Ordinal))
                {
                    return ActionAuthorization(
                        service,
                        requestId,
                        request.action);
                }

                if (!string.Equals(
                    request.type,
                    "chat",
                    StringComparison.Ordinal))
                {
                    return Error(
                        requestId,
                        "REQUEST_TYPE_NOT_ALLOWED",
                        "Only ping, chat, clearSession, openSettings, and authorizeBrowserAction requests are allowed.",
                        service);
                }

                if (!service.IsConfigured)
                {
                    return Error(
                        requestId,
                        "CONFIGURATION_INCOMPLETE",
                        "Open Scribble in Outlook, Excel, PowerPoint, or Word and configure its model connection first.",
                        service);
                }

                var context = request.context ??
                    new BrowserPageContext();
                using (var cancellation =
                    new CancellationTokenSource(
                        TimeSpan.FromSeconds(230)))
                {
                    var result = service.CompleteAsync(
                            NormalizeHistory(request.history),
                            request.prompt,
                            context.title,
                            context.url,
                            context.selection,
                            context.pageText,
                            context.links,
                            context.screenshotDataUrl,
                            NormalizeExchange(request.exchange),
                            request.chatId,
                            request.turnId,
                            request.topicId,
                            request.topicBinding,
                            cancellation.Token)
                        .GetAwaiter()
                        .GetResult();
                    if (result.HasPendingCalls)
                    {
                        return PendingTools(
                            service,
                            requestId,
                            result);
                    }

                    return Success(
                        service,
                        requestId,
                        result.Content,
                        result.Model,
                        result.ScreenshotUsed);
                }
            }
            catch (OperationCanceledException)
            {
                return Error(
                    requestId,
                    "AI_TIMEOUT",
                    "The browser request did not finish within 230 seconds.",
                    service);
            }
            catch (AiEndpointException exception)
            {
                return Error(
                    requestId,
                    exception.Code,
                    exception.Message,
                    service);
            }
            catch (Exception exception)
            {
                return Error(
                    requestId,
                    "BROWSER_HOST_FAILED",
                    SafeMessage(exception),
                    service);
            }
            finally
            {
                if (service != null)
                {
                    service.Dispose();
                }
            }
        }

        internal static bool TryRead(
            Stream input,
            out string json)
        {
            json = string.Empty;
            var lengthBytes = new byte[4];
            var first = input.ReadByte();
            if (first < 0)
            {
                return false;
            }

            lengthBytes[0] = (byte)first;
            ReadExactly(input, lengthBytes, 1, 3);
            if (!BitConverter.IsLittleEndian)
            {
                Array.Reverse(lengthBytes);
            }

            var length = BitConverter.ToInt32(lengthBytes, 0);
            if (length <= 0 || length > MaxRequestBytes)
            {
                throw new InvalidDataException(
                    "The native message length is outside the allowed boundary.");
            }

            var payload = new byte[length];
            ReadExactly(input, payload, 0, payload.Length);
            json = Utf8.GetString(payload);
            return true;
        }

        internal static void Write(
            Stream output,
            BrowserNativeResponse response)
        {
            var serializer = Serializer();
            var bytes = Utf8.GetBytes(
                serializer.Serialize(response ??
                    Error(
                        string.Empty,
                        "RESPONSE_EMPTY",
                        "The native host produced no response.")));
            if (bytes.Length > MaxResponseBytes)
            {
                bytes = Utf8.GetBytes(
                    serializer.Serialize(
                        Error(
                            string.Empty,
                            "RESPONSE_TOO_LARGE",
                            "The bounded native-host response was still too large.")));
            }

            var lengthBytes = BitConverter.GetBytes(bytes.Length);
            if (!BitConverter.IsLittleEndian)
            {
                Array.Reverse(lengthBytes);
            }

            output.Write(lengthBytes, 0, lengthBytes.Length);
            output.Write(bytes, 0, bytes.Length);
            output.Flush();
        }

        private static JavaScriptSerializer Serializer()
        {
            return new JavaScriptSerializer
            {
                MaxJsonLength = MaxRequestBytes,
                RecursionLimit = 100
            };
        }

        private static IReadOnlyList<BrowserExchangeTurn>
            NormalizeExchange(
                IReadOnlyList<BrowserNativeExchangeTurn> exchange)
        {
            var result = new List<BrowserExchangeTurn>();
            var exchangeCharacters = 0;
            if (exchange == null)
            {
                return result;
            }

            var exchangeStart = Math.Max(
                0,
                exchange.Count - BrowserChatRequestFactory.MaxExchangeTurns);
            for (var exchangeIndex = exchangeStart;
                 exchangeIndex < exchange.Count;
                 exchangeIndex++)
            {
                var turn = exchange[exchangeIndex];
                if (turn == null)
                {
                    continue;
                }

                var calls = new List<ChatToolCall>();
                foreach (var call in turn.toolCalls ??
                    new List<BrowserNativeToolCall>())
                {
                    if (call == null)
                    {
                        continue;
                    }

                    var callRemaining = Math.Max(
                        0,
                        MaxExchangeCharacters - exchangeCharacters);
                    var arguments = TextBoundary.PlainText(
                        call.arguments,
                        Math.Min(
                            BrowserChatRequestFactory.MaxToolArgumentCharacters,
                            callRemaining));
                    exchangeCharacters += arguments.Length;
                    calls.Add(new ChatToolCall
                    {
                        id = TextBoundary.SingleLine(call.id, 100),
                        type = "function",
                        function = new ChatToolCallFunction
                        {
                            name = TextBoundary.SingleLine(call.name, 100),
                            arguments = arguments
                        }
                    });
                }

                var results = new List<BrowserExchangeResult>();
                foreach (var toolResult in turn.results ??
                    new List<BrowserNativeToolResult>())
                {
                    if (toolResult == null)
                    {
                        continue;
                    }

                    var remaining = Math.Max(
                        0,
                        MaxExchangeCharacters - exchangeCharacters);
                    var content = TextBoundary.PlainText(
                        toolResult.content,
                        Math.Min(
                            MaxExchangeResultCharacters,
                            remaining));
                    exchangeCharacters += content.Length;
                    results.Add(new BrowserExchangeResult
                    {
                        Id = toolResult.id,
                        Content = content
                    });
                }

                var assistantRemaining = Math.Max(
                    0,
                    MaxExchangeCharacters - exchangeCharacters);
                var assistantContent = TextBoundary.PlainText(
                    turn.assistantContent,
                    Math.Min(
                        BrowserChatRequestFactory.MaxBrowserToolResultCharacters,
                        assistantRemaining));
                exchangeCharacters += assistantContent.Length;
                result.Add(new BrowserExchangeTurn
                {
                    AssistantContent = assistantContent,
                    ToolCalls = calls,
                    Results = results
                });
            }

            return result;
        }

        private static BrowserNativeResponse PendingTools(
            BrowserChatService service,
            string requestId,
            BrowserChatResult result)
        {
            var response = Success(
                service,
                requestId,
                string.Empty,
                result.Model,
                result.ScreenshotUsed);
            response.assistantContent = TextBoundary.PlainText(
                result.PendingAssistantContent,
                LimitOverrides.MaxAssistantCharactersLimit);
            response.toolRequests =
                new List<BrowserNativeToolCall>();
            foreach (var call in result.PendingCalls)
            {
                response.toolRequests.Add(new BrowserNativeToolCall
                {
                    id = TextBoundary.SingleLine(call.id, 100),
                    name = TextBoundary.SingleLine(
                        call.function.name,
                        100),
                    arguments = TextBoundary.PlainText(
                        call.function.arguments,
                        BrowserChatRequestFactory
                            .MaxToolArgumentCharacters)
                });
            }

            response.hostResults =
                new List<BrowserNativeToolResult>();
            foreach (var hostResult in result.HostResults)
            {
                var contentLimit = 120000;
                foreach (var pendingCall in result.PendingCalls)
                {
                    if (string.Equals(
                            pendingCall.id,
                            hostResult.Id,
                            StringComparison.Ordinal) &&
                        TopicToolCatalog.IsTopicTool(
                            pendingCall.function.name))
                    {
                        contentLimit = TopicToolHost
                            .MaxSerializedResultCharacters;
                        break;
                    }
                }

                response.hostResults.Add(
                    new BrowserNativeToolResult
                    {
                        id = hostResult.Id,
                        // Topic document text has a fixed semantic
                        // 120k limit but may be larger on the wire
                        // after JSON escaping. Every other host tool
                        // stays at the established 120k boundary.
                        content = TextBoundary.PlainText(
                            hostResult.Content,
                            contentLimit)
                    });
            }

            return response;
        }

        private static IReadOnlyList<ChatTurn> NormalizeHistory(
            IReadOnlyList<BrowserHistoryTurn> history)
        {
            var result = new List<ChatTurn>();
            if (history == null)
            {
                return result;
            }

            var start = Math.Max(
                0,
                history.Count - MaxHistoryTurns);
            for (var index = start;
                 index < history.Count;
                 index++)
            {
                var turn = history[index];
                if (turn == null ||
                    (turn.role != "user" &&
                     turn.role != "assistant"))
                {
                    continue;
                }

                var limit = turn.role == "user"
                    ? LimitOverrides.MaxPromptCharacters
                    : LimitOverrides.MaxAssistantCharactersLimit;
                result.Add(new ChatTurn(
                    turn.role,
                    TextBoundary.PlainText(
                        turn.content,
                        limit)));
            }

            return result;
        }

        private static BrowserNativeResponse Success(
            BrowserChatService service,
            string requestId,
            string content,
            string model,
            bool screenshotUsed)
        {
            return new BrowserNativeResponse
            {
                ok = true,
                content = TextBoundary.PlainText(
                    content,
                    LimitOverrides.MaxAssistantCharactersLimit),
                error = string.Empty,
                errorCode = string.Empty,
                model = TextBoundary.SingleLine(model, 200),
                configured = service != null && service.IsConfigured,
                supportsVision = service != null &&
                    service.SupportsVision,
                version = VersionText(),
                availableExtensionVersion = BundledExtensionVersion(),
                topics = BuildTopics(service)
            };
        }

        private static BrowserNativeResponse Error(
            string requestId,
            string code,
            string message,
            BrowserChatService service = null)
        {
            return new BrowserNativeResponse
            {
                ok = false,
                content = string.Empty,
                error = TextBoundary.PlainText(message, 1200),
                errorCode = TextBoundary.SingleLine(code, 100),
                model = TextBoundary.SingleLine(
                    service?.Model,
                    200),
                configured = service != null && service.IsConfigured,
                supportsVision = service != null &&
                    service.SupportsVision,
                version = VersionText(),
                availableExtensionVersion = BundledExtensionVersion(),
                topics = BuildTopics(service)
            };
        }

        private static BrowserNativeResponse ActionAuthorization(
            BrowserChatService service,
            string requestId,
            BrowserNativeActionDescriptor action)
        {
            var decision = BrowserActionPolicy.Evaluate(
                action == null
                    ? null
                    : new BrowserActionDescriptor
                    {
                        Action = TextBoundary.SingleLine(
                            action.action,
                            40),
                        TagName = TextBoundary.SingleLine(
                            action.tagName,
                            40),
                        InputType = TextBoundary.SingleLine(
                            action.inputType,
                            40),
                        Role = TextBoundary.SingleLine(
                            action.role,
                            80),
                        Name = TextBoundary.SingleLine(
                            action.name,
                            200),
                        Placeholder = TextBoundary.SingleLine(
                            action.placeholder,
                            200),
                        Autocomplete = TextBoundary.SingleLine(
                            action.autocomplete,
                            300),
                        Url = TextBoundary.SingleLine(
                            action.url,
                            BrowserChatRequestFactory.MaxUrlCharacters),
                        Value = TextBoundary.SingleLine(
                            action.value,
                            BrowserActionPolicy.MaxTypedCharacters + 1),
                        SourceText = TextBoundary.PlainText(
                            action.sourceText,
                            TextBoundary.MaxUserPromptCharacters),
                        Key = TextBoundary.SingleLine(
                            action.key,
                            40),
                        FormHasPassword = action.formHasPassword,
                        FormHasPayment = action.formHasPayment,
                        FormHasPersonalData = action.formHasPersonalData
                    });
            var response = Success(
                service,
                requestId,
                decision.Message,
                service?.Model,
                false);
            response.actionAllowed = decision.Allowed;
            response.actionCode = TextBoundary.SingleLine(
                decision.Code,
                100);
            return response;
        }

        private static List<BrowserNativeTopic> BuildTopics(
            BrowserChatService service)
        {
            var result = new List<BrowserNativeTopic>();
            if (service == null)
            {
                return result;
            }

            foreach (var topic in service.Topics)
            {
                result.Add(new BrowserNativeTopic
                {
                    id = TextBoundary.SingleLine(topic.Id, 40),
                    name = TextBoundary.SingleLine(topic.Name, 80),
                    binding = TextBoundary.SingleLine(
                        topic.Binding,
                        100),
                    available = topic.Available
                });
            }

            return result;
        }

        private static string VersionText()
        {
            try
            {
                return FileVersionInfo.GetVersionInfo(
                        typeof(BrowserChatService)
                            .Assembly.Location)
                    .FileVersion ?? "unknown";
            }
            catch
            {
                return "unknown";
            }
        }

        private static string BundledExtensionVersion()
        {
            try
            {
                var assemblyDirectory = Path.GetDirectoryName(
                    typeof(BrowserChatService).Assembly.Location);
                var manifestPath = Path.Combine(
                    assemblyDirectory ?? string.Empty,
                    "BrowserExtension",
                    "manifest.json");
                if (!File.Exists(manifestPath))
                {
                    return string.Empty;
                }

                var document = Serializer().DeserializeObject(
                    File.ReadAllText(manifestPath, Encoding.UTF8)) as
                    IDictionary<string, object>;
                object value;
                return document != null &&
                       document.TryGetValue("version", out value)
                    ? TextBoundary.SingleLine(
                        Convert.ToString(value),
                        40)
                    : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string CodeFor(Exception exception)
        {
            var endpoint = exception as AiEndpointException;
            return endpoint != null
                ? endpoint.Code
                : "NATIVE_MESSAGE_FAILED";
        }

        private static string SafeMessage(Exception exception)
        {
            if (exception == null)
            {
                return "The native host failed.";
            }

            if (exception is InvalidDataException ||
                exception is FormatException)
            {
                return TextBoundary.PlainText(
                    exception.Message,
                    1200);
            }

            return
                "The native host failed (" +
                exception.GetType().Name + ").";
        }

        private static void ReadExactly(
            Stream input,
            byte[] buffer,
            int offset,
            int count)
        {
            var read = 0;
            while (read < count)
            {
                var current = input.Read(
                    buffer,
                    offset + read,
                    count - read);
                if (current <= 0)
                {
                    throw new EndOfStreamException(
                        "The native message ended before its declared length.");
                }

                read += current;
            }
        }
    }
}
