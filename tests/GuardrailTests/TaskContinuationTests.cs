using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Configuration;
using Scribble.Outlook;
using Scribble.Office;

namespace GuardrailTests
{
    internal static class TaskContinuationTests
    {
        public static void ContextRecoveryAndPairing()
        {
            var root = Path.Combine(Path.GetTempPath(), "scribble-context-" + Guid.NewGuid().ToString("N"));
            try
            {
                using (var endpoint = new Endpoint())
                using (var client = new OpenAiCompatibleClient())
                {
                    var settings = new AppSettings { BaseUrl = endpoint.Url, Model = "local-model", ApiKey = "test" };
                    var request = Request(); var store = new TaskCheckpointStore(root);
                    var context = new TaskContextManager(request, "test", "KEEP ORIGINAL CONSTRAINT 2026", store);
                    for (var index = 0; index < 150; index++)
                    {
                        try { context.CompleteAsync(client, settings, request, null, CancellationToken.None).GetAwaiter().GetResult(); }
                        catch (Exception ex) { throw new Exception("Continuation iteration " + index + ", fake requests " + endpoint.RequestCount + ": " + (endpoint.Failure ?? ex).ToString()); }
                        var call = new ChatToolCall { id = "step-" + index, type = "function", function = new ChatToolCallFunction { name = "read_page", arguments = "{}" } };
                        var response = new ChatCompletionResponseMessage { tool_calls = new List<ChatToolCall> { call } };
                        var results = new List<MailboxToolResult> { new MailboxToolResult(call.id, "Evidence " + index + new string('x', 2200), "Read") };
                        context.PrepareExchange(response, request);
                        context.AfterTool(call, results[0]);
                        context.RecordExchange(request, response, results);
                        ChatRequestFactory.AppendToolExchange(request, response, results, "local-model");
                        context.FinishExchange(request);
                        if (index == 74)
                        {
                            context.Pause("Test restart"); request = Request();
                            context = new TaskContextManager(request, "test", context.State.Objective, store, store.Load(context.State.Id));
                        }
                    }
                    if (endpoint.Failure != null) throw endpoint.Failure;
                    if (endpoint.RequestCount <= 150 || context.State.ContextBudget >= 96000 || context.State.EvidenceIds.Count < 150)
                        throw new Exception("Context rejection, compaction or durable evidence did not execute.");
                }
            }
            finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        }
        public static void ReviewRepairsAndAlignment()
        {
            var sources = new[] { new ExcelTaskCell { Id = "A1", Value = "원본" }, new ExcelTaskCell { Id = "A2", Value = "" } };
            string terminology;
            var values = ExcelReviewValidator.Validate("{\"rows\":[{\"id\":\"A1\",\"value\":\"Repaired translation\"},{\"id\":\"A2\",\"value\":\"\"}],\"terminology\":\"Shared term\"}", sources, out terminology);
            if (values[0] != "Repaired translation" || terminology != "Shared term") throw new Exception("Review repair was discarded.");
            var rejected = false;
            try { ExcelReviewValidator.Validate("{\"rows\":[{\"id\":\"A2\",\"value\":\"Wrong row\"},{\"id\":\"A1\",\"value\":\"\"}]}", sources, out terminology); }
            catch (InvalidOperationException) { rejected = true; }
            if (!rejected) throw new Exception("Review accepted swapped row identities.");
        }
        private static ChatCompletionRequest Request() { return new ChatCompletionRequest { model = "local-model", messages = new List<object> { new ChatCompletionInputMessage { role = "user", content = "KEEP ORIGINAL CONSTRAINT 2026" } } }; }

        private sealed class Endpoint : IDisposable
        {
            private readonly TcpListener _listener = new TcpListener(IPAddress.Loopback, 0);
            private readonly Task _worker;
            public string Url;
            public int RequestCount;
            public Exception Failure;
            public Endpoint()
            {
                _listener.Start(); Url = "http://127.0.0.1:" + ((IPEndPoint)_listener.LocalEndpoint).Port + "/v1";
                _worker = Task.Run((Action)Serve);
            }
            private void Serve()
            {
                try
                {
                    while (true) using (var client = _listener.AcceptTcpClient()) using (var stream = client.GetStream())
                    {
                        var header = new List<byte>(); int next;
                        while ((next = stream.ReadByte()) >= 0) { header.Add((byte)next); if (header.Count >= 4 && Encoding.ASCII.GetString(header.Skip(header.Count - 4).ToArray()) == "\r\n\r\n") break; }
                        var headers = Encoding.ASCII.GetString(header.ToArray());
                        var length = int.Parse(headers.Split('\n').First(line => line.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase)).Split(':')[1].Trim());
                        if (headers.IndexOf("100-continue", StringComparison.OrdinalIgnoreCase) >= 0) { var interim = Encoding.ASCII.GetBytes("HTTP/1.1 100 Continue\r\n\r\n"); stream.Write(interim, 0, interim.Length); }
                        var bytes = new byte[length]; var read = 0;
                        while (read < length) { var count = stream.Read(bytes, read, length - read); if (count == 0) throw new IOException("Request ended early"); read += count; }
                        var body = Encoding.UTF8.GetString(bytes);
                        var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.Deserialize<Dictionary<string, object>>(body);
                        var pending = new HashSet<string>();
                        foreach (Dictionary<string, object> message in (IEnumerable)json["messages"])
                        {
                            object calls;
                            if (message.TryGetValue("tool_calls", out calls) && calls != null) foreach (Dictionary<string, object> call in (IEnumerable)calls) pending.Add((string)call["id"]);
                            if ((string)message["role"] == "tool" && !pending.Remove((string)message["tool_call_id"])) throw new Exception("Orphan tool result in model request");
                        }
                        if (pending.Count != 0) throw new Exception("Model request omitted tool results");
                        var reject = RequestCount++ == 0;
                        var response = reject ? "{\"error\":{\"message\":\"context length exceeded\",\"code\":\"context_length_exceeded\"}}" : "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"content\":\"KEEP ORIGINAL CONSTRAINT 2026. Continue with archived evidence.\"}}]}";
                        bytes = Encoding.UTF8.GetBytes(response);
                        var prefix = Encoding.ASCII.GetBytes("HTTP/1.1 " + (reject ? "400 Bad Request" : "200 OK") + "\r\nContent-Type: application/json\r\nContent-Length: " + bytes.Length + "\r\nConnection: close\r\n\r\n");
                        stream.Write(prefix, 0, prefix.Length); stream.Write(bytes, 0, bytes.Length);
                    }
                }
                catch (SocketException) { }
                catch (ObjectDisposedException) { }
                catch (Exception ex) { Failure = ex; }
            }
            public void Dispose() { _listener.Stop(); _worker.Wait(TimeSpan.FromSeconds(2)); }
        }
    }
}
