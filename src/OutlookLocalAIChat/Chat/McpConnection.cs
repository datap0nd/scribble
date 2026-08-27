using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using OutlookLocalAIChat.Configuration;
using OutlookLocalAIChat.Security;
using OutlookLocalAIChat.Utilities;

namespace OutlookLocalAIChat.Chat
{
    public sealed class McpToolDescriptor
    {
        public McpToolDescriptor(
            string name,
            string description,
            object schema)
        {
            Name = name ?? string.Empty;
            Description = description ?? string.Empty;
            Schema = schema;
        }

        public string Name { get; }

        public string Description { get; }

        public object Schema { get; }
    }

    // One live MCP session over stdio (local command) or streamable
    // HTTP. JSON-RPC 2.0, newline-delimited on stdio. Every response
    // is size-bounded before parsing and every call is time-capped;
    // a timed-out stdio server is killed rather than left blocking.
    public sealed class McpConnection : IDisposable
    {
        private const string ProtocolVersion = "2025-03-26";
        // Real stdio servers (npm- or pip-launched) can take a long
        // time to cold-start before they answer initialize.
        private const int InitializeTimeoutMs = 60000;
        private const int ListTimeoutMs = 30000;
        private const int CallTimeoutMs = 60000;

        private readonly McpServerConfig _config;
        private readonly HttpClient _httpClient;
        private readonly int _maxOperationTimeoutMs;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();
        private readonly object _gate = new object();
        // Bounded tail of the server's stderr, surfaced in error
        // messages so a crashing or misconfigured server is
        // diagnosable instead of just silent.
        private readonly StringBuilder _stderrTail =
            new StringBuilder();
        private Process _process;
        private StreamWriter _stdin;
        private StreamReader _stdout;
        private string _httpSessionId = string.Empty;
        private bool _initialized;
        private bool _failed;
        private int _nextId = 1;

        public McpConnection(
            McpServerConfig config,
            HttpClient httpClient)
            : this(config, httpClient, int.MaxValue)
        {
        }

        internal McpConnection(
            McpServerConfig config,
            HttpClient httpClient,
            int maxOperationTimeoutMs)
        {
            _config = config ??
                throw new ArgumentNullException(nameof(config));
            _httpClient = httpClient;
            _maxOperationTimeoutMs = Math.Max(
                1000,
                maxOperationTimeoutMs);
        }

        public string ServerName
        {
            get { return _config.Name; }
        }

        public IReadOnlyList<McpToolDescriptor> ListTools()
        {
            lock (_gate)
            {
                EnsureInitialized();
                var result = Request(
                    "tools/list",
                    new Dictionary<string, object>(),
                    BoundedTimeout(ListTimeoutMs));
                var tools = new List<McpToolDescriptor>();
                object toolsValue;
                if (result == null ||
                    !result.TryGetValue("tools", out toolsValue))
                {
                    return tools;
                }

                var array = toolsValue as object[];
                if (array == null)
                {
                    return tools;
                }

                foreach (var item in array)
                {
                    var tool = item as
                        IDictionary<string, object>;
                    if (tool == null)
                    {
                        continue;
                    }

                    object nameValue;
                    object descriptionValue;
                    object schemaValue;
                    tool.TryGetValue("name", out nameValue);
                    tool.TryGetValue(
                        "description",
                        out descriptionValue);
                    tool.TryGetValue(
                        "inputSchema",
                        out schemaValue);
                    var name = TextBoundary.SingleLine(
                        Convert.ToString(nameValue),
                        64);
                    if (name.Length == 0)
                    {
                        continue;
                    }

                    tools.Add(new McpToolDescriptor(
                        name,
                        TextBoundary.PlainText(
                            Convert.ToString(descriptionValue),
                            600),
                        schemaValue));
                }

                return tools;
            }
        }

        // Returns the concatenated text content of the tool result;
        // isError reflects the MCP-level error flag.
        public string CallTool(
            string toolName,
            IDictionary<string, object> arguments,
            out bool isError)
        {
            lock (_gate)
            {
                EnsureInitialized();
                var result = Request(
                    "tools/call",
                    new Dictionary<string, object>
                    {
                        { "name", toolName },
                        {
                            "arguments",
                            arguments ??
                            new Dictionary<string, object>()
                        }
                    },
                    BoundedTimeout(CallTimeoutMs));
                isError = false;
                if (result == null)
                {
                    return string.Empty;
                }

                object isErrorValue;
                if (result.TryGetValue(
                        "isError",
                        out isErrorValue))
                {
                    isError = isErrorValue is bool &&
                        (bool)isErrorValue;
                }

                var builder = new StringBuilder();
                object contentValue;
                if (result.TryGetValue(
                        "content",
                        out contentValue))
                {
                    var array = contentValue as object[];
                    if (array != null)
                    {
                        foreach (var item in array)
                        {
                            var part = item as
                                IDictionary<string, object>;
                            if (part == null)
                            {
                                continue;
                            }

                            object typeValue;
                            part.TryGetValue(
                                "type",
                                out typeValue);
                            if (Convert.ToString(typeValue) !=
                                "text")
                            {
                                builder.Append(
                                    "[non-text content of type " +
                                    TextBoundary.SingleLine(
                                        Convert.ToString(
                                            typeValue),
                                        40) +
                                    " omitted]\n");
                                continue;
                            }

                            object textValue;
                            part.TryGetValue(
                                "text",
                                out textValue);
                            builder.Append(
                                Convert.ToString(textValue) ??
                                string.Empty);
                            builder.Append('\n');
                        }
                    }
                }

                return TextBoundary.PlainText(
                    builder.ToString(),
                    ContextScale.Scaled(
                        TextBoundary.MaxToolResultCharacters) / 2);
            }
        }

        public void Dispose()
        {
            lock (_gate)
            {
                KillProcess();
                _initialized = false;
            }
        }

        private void EnsureInitialized()
        {
            if (_initialized)
            {
                return;
            }

            if (_failed)
            {
                throw new InvalidOperationException(
                    "The MCP server " + _config.Name +
                    " failed earlier in this session.");
            }

            try
            {
                if (!_config.IsHttp)
                {
                    StartProcess();
                }

                var result = Request(
                    "initialize",
                    new Dictionary<string, object>
                    {
                        { "protocolVersion", ProtocolVersion },
                        {
                            "capabilities",
                            new Dictionary<string, object>()
                        },
                        {
                            "clientInfo",
                            new Dictionary<string, object>
                            {
                                { "name", "AI365" },
                                { "version", "2.0" }
                            }
                        }
                    },
                    BoundedTimeout(InitializeTimeoutMs));
                if (result == null)
                {
                    throw new InvalidOperationException(
                        "The MCP server returned no initialize result.");
                }

                Notify("notifications/initialized");
                _initialized = true;
            }
            catch (Exception)
            {
                _failed = true;
                KillProcess();
                throw;
            }
        }

        private void StartProcess()
        {
            if (_process != null && !_process.HasExited)
            {
                return;
            }

            // Canonical redirection only: the JSON-RPC wire stays
            // pure ASCII because the serializer escapes non-ASCII
            // characters, so no custom stream encoding is needed.
            var info = new ProcessStartInfo
            {
                FileName = _config.Target,
                Arguments = _config.Arguments,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            _process = Process.Start(info);
            if (_process == null)
            {
                throw new InvalidOperationException(
                    "The MCP server process could not be started.");
            }

            // Only the built-in standard-input writer reliably
            // delivers to the child (verified empirically: raw
            // writes to its BaseStream never arrive). Its first
            // write emits an encoding preamble (BOM), which strict
            // JSON-RPC servers reject when glued to a message - so
            // the preamble is flushed inside a discardable blank
            // line first, and every real message starts clean.
            _stdin = _process.StandardInput;
            _stdin.AutoFlush = true;
            _stdin.WriteLine();
            _stdout = _process.StandardOutput;
            // Drain stderr so a chatty server can never block,
            // keeping a bounded tail for diagnostics.
            _process.ErrorDataReceived +=
                (sender, eventArgs) =>
                {
                    var data = eventArgs.Data;
                    if (string.IsNullOrEmpty(data))
                    {
                        return;
                    }

                    lock (_stderrTail)
                    {
                        if (_stderrTail.Length > 2000)
                        {
                            _stderrTail.Length = 0;
                        }

                        _stderrTail.AppendLine(data);
                    }
                };
            _process.BeginErrorReadLine();
        }

        private void KillProcess()
        {
            try
            {
                if (_process != null && !_process.HasExited)
                {
                    _process.Kill();
                }
            }
            catch
            {
            }

            try
            {
                _process?.Dispose();
            }
            catch
            {
            }

            _process = null;
            _stdin = null;
            _stdout = null;
        }

        private void Notify(string method)
        {
            var message = new Dictionary<string, object>
            {
                { "jsonrpc", "2.0" },
                { "method", method }
            };
            if (_config.IsHttp)
            {
                try
                {
                    PostHttp(
                        _serializer.Serialize(message),
                        BoundedTimeout(InitializeTimeoutMs));
                }
                catch (Exception exception)
                {
                    Log.Error("McpNotify", exception);
                }

                return;
            }

            _stdin.WriteLine(_serializer.Serialize(message));
        }

        private IDictionary<string, object> Request(
            string method,
            Dictionary<string, object> parameters,
            int timeoutMs)
        {
            var id = _nextId++;
            var json = _serializer.Serialize(
                new Dictionary<string, object>
                {
                    { "jsonrpc", "2.0" },
                    { "id", id },
                    { "method", method },
                    { "params", parameters }
                });
            var responseJson = _config.IsHttp
                ? PostHttp(json, timeoutMs)
                : ExchangeStdio(json, id, timeoutMs);
            var response = ParseMessage(responseJson, id);
            if (response == null)
            {
                throw new InvalidOperationException(
                    "The MCP server returned no response to " +
                    method + ".");
            }

            object errorValue;
            if (response.TryGetValue("error", out errorValue))
            {
                var error = errorValue as
                    IDictionary<string, object>;
                object messageValue = null;
                error?.TryGetValue(
                    "message",
                    out messageValue);
                throw new InvalidOperationException(
                    "MCP error from " + _config.Name + ": " +
                    TextBoundary.SingleLine(
                        Convert.ToString(messageValue),
                        300));
            }

            object resultValue;
            response.TryGetValue("result", out resultValue);
            return resultValue as IDictionary<string, object>;
        }

        private string ExchangeStdio(
            string requestJson,
            int id,
            int timeoutMs)
        {
            _stdin.WriteLine(requestJson);
            var reader = _stdout;
            // Lines the server wrote that are not our response are
            // skipped, but a bounded sample is kept so a protocol
            // mismatch is diagnosable from the error message.
            var skippedLines = new List<string>();
            var task = Task.Run(() =>
            {
                while (true)
                {
                    var line = reader.ReadLine();
                    if (line == null)
                    {
                        return (string)null;
                    }

                    if (line.Length >
                        TextBoundary.MaxHttpResponseCharacters)
                    {
                        throw new InvalidOperationException(
                            "The MCP server response exceeded the size cap.");
                    }

                    // Some servers emit an encoding preamble before
                    // their first line; tolerate it.
                    line = line.TrimStart('\uFEFF');
                    if (MessageHasId(line, id))
                    {
                        return line;
                    }

                    lock (skippedLines)
                    {
                        if (skippedLines.Count < 5)
                        {
                            skippedLines.Add(
                                TextBoundary.SingleLine(
                                    line,
                                    200));
                        }
                    }
                }
            });
            if (!task.Wait(timeoutMs))
            {
                KillProcess();
                _initialized = false;
                _failed = true;
                throw new TimeoutException(
                    "The MCP server " + _config.Name +
                    " did not answer within " +
                    (timeoutMs / 1000) + " seconds." +
                    SkippedSuffix(skippedLines) +
                    StderrTailSuffix());
            }

            if (task.Result == null)
            {
                _failed = true;
                throw new InvalidOperationException(
                    "The MCP server " + _config.Name +
                    " closed its output stream." +
                    SkippedSuffix(skippedLines) +
                    StderrTailSuffix());
            }

            return task.Result;
        }

        private static string SkippedSuffix(
            List<string> skippedLines)
        {
            lock (skippedLines)
            {
                return skippedLines.Count > 0
                    ? " Unmatched server output: " +
                      string.Join(" | ", skippedLines)
                    : " The server wrote no output.";
            }
        }

        private string StderrTailSuffix()
        {
            lock (_stderrTail)
            {
                var tail = TextBoundary.SingleLine(
                    _stderrTail.ToString(),
                    600);
                return tail.Length > 0
                    ? " Server stderr: " + tail
                    : string.Empty;
            }
        }

        // Unrelated server notifications and requests are skipped;
        // only the response carrying our request id is returned.
        private bool MessageHasId(string line, int id)
        {
            try
            {
                var parsed = _serializer.DeserializeObject(line) as
                    IDictionary<string, object>;
                object idValue;
                if (parsed == null ||
                    !parsed.TryGetValue("id", out idValue))
                {
                    return false;
                }

                int parsedId;
                return int.TryParse(
                    Convert.ToString(idValue),
                    out parsedId) && parsedId == id;
            }
            catch
            {
                return false;
            }
        }

        private string PostHttp(string json, int timeoutMs)
        {
            var request = new HttpRequestMessage(
                HttpMethod.Post,
                _config.Target)
            {
                Content = new StringContent(
                    json,
                    new UTF8Encoding(false),
                    "application/json")
            };
            request.Headers.TryAddWithoutValidation(
                "Accept",
                "application/json, text/event-stream");
            foreach (var header in _config.ParsedHeaders())
            {
                request.Headers.TryAddWithoutValidation(
                    header.Key,
                    header.Value);
            }

            if (_httpSessionId.Length > 0)
            {
                request.Headers.TryAddWithoutValidation(
                    "Mcp-Session-Id",
                    _httpSessionId);
            }

            var sendTask = _httpClient.SendAsync(request);
            if (!sendTask.Wait(timeoutMs))
            {
                throw new TimeoutException(
                    "The MCP server " + _config.Name +
                    " did not answer within " +
                    (timeoutMs / 1000) + " seconds.");
            }

            using (var response = sendTask.Result)
            {
                IEnumerable<string> sessionValues;
                if (response.Headers.TryGetValues(
                        "Mcp-Session-Id",
                        out sessionValues))
                {
                    foreach (var value in sessionValues)
                    {
                        _httpSessionId = TextBoundary.SingleLine(
                            value,
                            200);
                        break;
                    }
                }

                var readTask =
                    response.Content.ReadAsStringAsync();
                if (!readTask.Wait(timeoutMs))
                {
                    throw new TimeoutException(
                        "The MCP server response body timed out.");
                }

                var body = readTask.Result ?? string.Empty;
                if (body.Length >
                    TextBoundary.MaxHttpResponseCharacters)
                {
                    throw new InvalidOperationException(
                        "The MCP server response exceeded the size cap.");
                }

                if (!response.IsSuccessStatusCode &&
                    (int)response.StatusCode != 202)
                {
                    throw new InvalidOperationException(
                        "MCP HTTP " +
                        (int)response.StatusCode +
                        " from " + _config.Name + ": " +
                        TextBoundary.SingleLine(body, 300));
                }

                return body;
            }
        }

        private int BoundedTimeout(int requested)
        {
            return Math.Min(requested, _maxOperationTimeoutMs);
        }

        // Accepts either a bare JSON-RPC message or an SSE stream
        // and returns the raw JSON of the message with our id.
        private IDictionary<string, object> ParseMessage(
            string body,
            int id)
        {
            if (string.IsNullOrWhiteSpace(body))
            {
                return null;
            }

            var trimmed = body.TrimStart();
            if (trimmed.StartsWith(
                "{",
                StringComparison.Ordinal))
            {
                return _serializer.DeserializeObject(trimmed) as
                    IDictionary<string, object>;
            }

            foreach (var rawLine in body.Split('\n'))
            {
                var line = rawLine.TrimEnd('\r');
                if (!line.StartsWith(
                    "data:",
                    StringComparison.Ordinal))
                {
                    continue;
                }

                var payload = line.Substring(5).Trim();
                if (payload.Length == 0)
                {
                    continue;
                }

                if (MessageHasId(payload, id))
                {
                    return _serializer
                        .DeserializeObject(payload) as
                        IDictionary<string, object>;
                }
            }

            return null;
        }
    }
}
