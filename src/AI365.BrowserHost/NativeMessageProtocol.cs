using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using OutlookLocalAIChat.Chat;
using OutlookLocalAIChat.Security;

namespace AI365.BrowserHost
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

        public string screenshotDataUrl { get; set; }
    }

    internal sealed class BrowserNativeRequest
    {
        public string type { get; set; }

        public string requestId { get; set; }

        public string prompt { get; set; }

        public List<BrowserHistoryTurn> history { get; set; }

        public BrowserPageContext context { get; set; }
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
    }

    internal static class NativeMessageProtocol
    {
        internal const int MaxRequestBytes = 16 * 1024 * 1024;
        internal const int MaxResponseBytes = 900 * 1024;
        internal const int MaxHistoryTurns = 12;

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
                        "AI365 native messaging failed: " +
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

                if (!string.Equals(
                    request.type,
                    "chat",
                    StringComparison.Ordinal))
                {
                    return Error(
                        requestId,
                        "REQUEST_TYPE_NOT_ALLOWED",
                        "Only ping and chat requests are allowed.",
                        service);
                }

                if (!service.IsConfigured)
                {
                    return Error(
                        requestId,
                        "CONFIGURATION_INCOMPLETE",
                        "Open AI365 in Outlook, Excel, PowerPoint, or Word and configure its model connection first.",
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
                            context.screenshotDataUrl,
                            cancellation.Token)
                        .GetAwaiter()
                        .GetResult();
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
                version = VersionText()
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
                version = VersionText()
            };
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
