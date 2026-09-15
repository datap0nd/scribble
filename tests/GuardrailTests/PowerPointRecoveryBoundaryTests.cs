using System;
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
using Scribble.Office;
using Scribble.Security;
using Scribble.Testing;

namespace GuardrailTests
{
    internal static class PowerPointRecoveryBoundaryTests
    {
        private const string Objective = "Create a launch presentation";
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        private static void Check(bool value, string message) { if (!value) throw new InvalidOperationException(message); }
        private static ChatCompletionRequest Request()
        {
            return new ChatCompletionRequest { model = "qwen3-vl", messages = new List<object> {
                new ChatCompletionInputMessage { role = "user", content = Objective } },
                tools = new List<ChatToolDefinition> { PresentationToolCatalog.DraftDefinition() } };
        }
        private static ChatToolCall SlideCall(string id)
        {
            return new ChatToolCall { id = id, function = new ChatToolCallFunction {
                name = PresentationToolCatalog.AddDraftSlides,
                arguments = Json.Serialize(new { plan = new[] { id }, slides = new[] { new { id, title = "Launch", layout = "cover" } } }) } };
        }
        private static MailboxToolResult Read(TaskContextManager task, string id)
        {
            return task.ReadEvidence(new ChatToolCall { id = "read", function = new ChatToolCallFunction {
                name = TaskContextManager.ReadEvidenceTool, arguments = Json.Serialize(new { id, offset = 0 }) } });
        }
        private static void CheckPayload(TaskContextManager task, string id, string expected)
        {
            var result = Read(task, id);
            Check(!result.Outcome.Failed, "The advertised recovery payload was rejected: " + result.Content);
            var data = Json.Deserialize<Dictionary<string, object>>(result.Content);
            Check(Convert.ToString(data["text"]) == expected, "Recovery did not return the exact original payload.");
        }

        public static void CanvasRejectionKeepsRetryAvailable()
        {
            var root = Path.Combine(Path.GetTempPath(), "scribble-canvas-boundary-" + Guid.NewGuid().ToString("N"));
            try
            {
                using (var endpoint = new ApprovingEndpoint())
                using (var client = new OpenAiCompatibleClient())
                {
                    var app = new RecoveryPowerPointApplication();
                    var store = new TaskCheckpointStore(root);
                    var request = Request();
                    var task = new TaskContextManager(request, "powerpoint", Objective, store);
                    var authorization = new OneShotDraftAuthorization(true);
                    var settings = new AppSettings { Model = "qwen3-vl", BaseUrl = endpoint.Url, ApiKey = "test" };
                    using (var host = new DocumentDraftHost("powerpoint", app))
                    {
                        host.BindTaskAsync(task, CancellationToken.None).GetAwaiter().GetResult();
                        var rejected = SlideCall("rejected-plan");
                        Check(task.ValidateArguments(rejected) == null, "The fixture must reach the native canvas preflight.");
                        task.BeforeTool(rejected, true);
                        var result = host.ExecuteAsync(rejected, authorization, true, Objective, client, settings, CancellationToken.None, null).GetAwaiter().GetResult();
                        task.AfterTool(rejected, result);
                        Check(result.Content.Contains("SAMSUNG_CANVAS_MISMATCH") && result.Outcome.PermissionConsumed == false,
                            "Canvas rejection was classified as an uncertain native write: " + result.Content);
                        Check(authorization.RemainingCalls == 1 && !authorization.IsConsumed, "Canvas rejection spent the write permission.");
                        Check(app.ActivePresentation.Tags.Writes == 0 && app.ActivePresentation.Slides.AddAttempts == 0,
                            "Canvas rejection mutated the presentation.");
                        Check(!task.State.HostData.ContainsKey("samsung_plan") && !task.State.HostData.ContainsKey("samsung_pending") &&
                            !task.State.ExpectedSourceIds.Any(id => id.StartsWith("ppt:")) && task.State.Writes.All(w => w.Status == "verified"),
                            "A rejected canvas left a locked plan or uncertain write.");

                        // The operator corrects the canvas; the next request may
                        // change the still-unwritten plan. A fault *after* Add
                        // starts must retain uncertain native state and receipts.
                        app.ActivePresentation.PageSetup.SlideWidth = 960;
                        app.ActivePresentation.PageSetup.SlideHeight = 540;
                        var retry = SlideCall("corrected-plan");
                        Check(task.ValidateArguments(retry) == null, "The corrected retry was blocked by the rejected attempt.");
                        task.BeforeTool(retry, true);
                        result = host.ExecuteAsync(retry, authorization, true, Objective, client, settings, CancellationToken.None, null).GetAwaiter().GetResult();
                        task.AfterTool(retry, result);
                        Check(result.Content.Contains("injected native interruption") && result.Outcome.PermissionConsumed == true,
                            "A partial native mutation was incorrectly marked safe to repeat: " + result.Content);
                        Check(app.ActivePresentation.Slides.AddAttempts == 1 && app.ActivePresentation.Slides.Count == 2 &&
                            task.State.Writes.Last().Status == "uncertain" && task.State.HostData.ContainsKey("samsung_pending"),
                            "The corrected call did not reach the journaled native write boundary.");
                        var payloadId = task.State.HostData["samsung_recovery_payload"];
                        CheckPayload(task, payloadId, retry.function.arguments);
                        task.SaveRequest(request);
                        var resumed = new TaskContextManager(Request(), "powerpoint", Objective, store, store.Load(task.State.Id));
                        CheckPayload(resumed, payloadId, retry.function.arguments);
                        var foreign = new TaskContextManager(Request(), "powerpoint", Objective, store);
                        Check(Read(foreign, payloadId).Outcome.Failed, "Recovery evidence became readable from another task.");
                    }
                }
            }
            finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        }

        public static void LegacyRecoveryPayloadsRemainTaskScoped()
        {
            var root = Path.Combine(Path.GetTempPath(), "scribble-payload-boundary-" + Guid.NewGuid().ToString("N"));
            try
            {
                var store = new TaskCheckpointStore(root);
                var request = Request();
                var task = new TaskContextManager(request, "powerpoint", Objective, store);
                var expected = new Dictionary<string, string> {
                    { "samsung_recovery_payload", "{\"slides\":[{\"title\":\"Original € slide\"}]}" },
                    { "powerpoint_revision_payload", "{\"operations\":[{\"kind\":\"notes_append\",\"text\":\"Original revision\"}]}" } };
                foreach (var pair in expected)
                {
                    var id = store.PutEvidence(task.State.Id, pair.Value);
                    task.State.HostData[pair.Key] = id;
                    Check(Read(task, id).Outcome.Failed, "Unregistered evidence unexpectedly bypassed the read boundary.");
                }
                var unrelated = store.PutEvidence(task.State.Id, "Unlisted task evidence");
                using (var host = new DocumentDraftHost("powerpoint", new object()))
                    host.BindTaskAsync(task, CancellationToken.None).GetAwaiter().GetResult();
                foreach (var pair in expected) CheckPayload(task, task.State.HostData[pair.Key], pair.Value);
                Check(Read(task, unrelated).Outcome.Failed, "Migration exposed unrelated archived evidence.");
                task.SaveRequest(request);
                var resumed = new TaskContextManager(Request(), "powerpoint", Objective, store, store.Load(task.State.Id));
                foreach (var pair in expected) CheckPayload(resumed, resumed.State.HostData[pair.Key], pair.Value);

                var foreign = new TaskContextManager(Request(), "powerpoint", Objective, store);
                var foreignId = foreign.RegisterEvidence("Private foreign recovery payload");
                Check(Read(resumed, foreignId).Outcome.Failed, "A foreign task's evidence ID was accepted.");
                resumed.State.HostData["samsung_recovery_payload"] = foreignId;
                var blocked = false;
                try
                {
                    using (var host = new DocumentDraftHost("powerpoint", new object()))
                        host.BindTaskAsync(resumed, CancellationToken.None).GetAwaiter().GetResult();
                }
                catch (IOException) { blocked = true; }
                Check(blocked && Read(resumed, foreignId).Outcome.Failed,
                    "A forged migration reference gained access to another task's payload.");
            }
            finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        }

        public static void OperatorArgumentsAndOutcomes()
        {
            var path = Path.Combine(Path.GetTempPath(), "scribble-operator-result-" + Guid.NewGuid().ToString("N") + ".json");
            var options = TestLabOperatorOptions.Parse(new[] { "--test-lab-run", "--result-json", path, "--case", "ex01" });
            Check(options.CaseId == "EX01" && options.ResultPath == path && !File.Exists(path), "Operator parsing changed scope or started work.");
            foreach (var invalid in new[] {
                new[] { "--test-lab-run" },
                new[] { "--test-lab-run", "--result-json", "relative.json" },
                new[] { "--test-lab-run", "--result-json", @"C:relative.json" },
                new[] { "--test-lab-run", "--result-json", path, "--case", "PP99" },
                new[] { "--test-lab-run", "--result-json", path, "--case", "CH01" },
                new[] { "--test-lab-run", "--result-json", path, "--case", "EX01", "--case", "PP01" },
                new[] { "--test-lab-run", "--result-json", path, "--unknown", "value" },
                new[] { "--test-lab-run", "--result-json", path, "--case" } })
            {
                var rejected = false;
                try { TestLabOperatorOptions.Parse(invalid); } catch (ArgumentException) { rejected = true; }
                Check(rejected, "Unsafe or ambiguous operator arguments were accepted: " + string.Join(" ", invalid));
            }
            var cases = TestLabOperatorOptions.OfficeCaseIds.Select(id => new SuiteCaseResult { id = id, status = "needs_review" }).ToArray();
            var result = TestLabOperatorResult.Classify(null, new SuiteState(), cases, "report.pdf", true, false, null);
            Check(result.exit_code == 0 && result.harness_complete && result.correctness_status == "needs_review" && result.completed_count == 16,
                "A fully completed suite lost its explicit model-review boundary.");
            cases[0].status = "failed";
            result = TestLabOperatorResult.Classify(null, new SuiteState(), cases, "report.pdf", true, false, null);
            Check(result.exit_code == 3 && result.harness_complete && result.correctness_status == "failed", "Correctness failure was confused with infrastructure completion.");
            foreach (var status in new[] { "blocked", "incomplete", "not_run", "stopped", "running" })
            {
                cases[0].status = status;
                result = TestLabOperatorResult.Classify(null, new SuiteState(), cases, "report.pdf", true, false, null);
                Check(result.exit_code == 2 && !result.harness_complete, "Operator reported success with case status " + status);
            }
            cases[0].status = "needs_review";
            Check(TestLabOperatorResult.Classify(null, null, cases.Take(15).ToArray(), "report.pdf", true, false, null).exit_code == 2,
                "Fewer than all requested cases counted as success.");
            Check(TestLabOperatorResult.Classify(null, null, new SuiteCaseResult[0], "report.pdf", true, false, null).exit_code == 2,
                "An empty PDF run counted as success.");
            Check(TestLabOperatorResult.Classify(null, null, cases, null, false, false, null).exit_code == 2,
                "Missing final PDF counted as success.");
            Check(TestLabOperatorResult.Classify(null, null, cases, "report.pdf", true, true, null).exit_code == 4,
                "Cancellation was not reported distinctly.");
            var one = new[] { new SuiteCaseResult { id = "EX01", status = "needs_review" } };
            Check(TestLabOperatorResult.Classify("EX01", null, one, "report.pdf", true, false, null).exit_code == 0 &&
                TestLabOperatorResult.Classify("PP01", null, one, "report.pdf", true, false, null).exit_code == 2,
                "Single-case operator output did not verify exact case identity.");
        }

        public static void OperatorLoggingFailureRetainsTerminalBoundary()
        {
            var runner = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            var logged = 0;
            var terminal = TestLabOperatorExecution.RunToTerminalAsync(() => runner.Task, error => {
                logged++; throw new UnauthorizedAccessException("diagnostic folder denied");
            });
            Check(!terminal.IsCompleted && logged == 0, "The operator released a still-running suite.");
            runner.SetException(new IOException("runner output folder denied"));
            var failure = terminal.GetAwaiter().GetResult();
            Check(logged == 1 && failure.Contains("runner output folder denied") && failure.Contains("diagnostic folder denied"),
                "A logging failure escaped or hid the terminal runner failure.");
            Check(TestLabOperatorExecution.RunToTerminalAsync(() => Task.FromResult(true), error => {
                throw new InvalidOperationException("Successful runners must not log a failure.");
            }).GetAwaiter().GetResult() == null, "A successful runner was reported as failed.");
        }

        private sealed class ApprovingEndpoint : IDisposable
        {
            private readonly TcpListener listener = new TcpListener(IPAddress.Loopback, 0);
            private readonly Task worker;
            internal string Url { get; }
            internal ApprovingEndpoint()
            {
                listener.Start();
                Url = "http://127.0.0.1:" + ((IPEndPoint)listener.LocalEndpoint).Port + "/v1";
                worker = Task.Run(() =>
                {
                    try
                    {
                        while (true)
                        using (var connection = listener.AcceptTcpClient())
                        using (var stream = connection.GetStream())
                        {
                            stream.ReadTimeout = 5000;
                            var header = new StringBuilder();
                            while (!header.ToString().EndsWith("\r\n\r\n", StringComparison.Ordinal))
                            {
                                var next = stream.ReadByte();
                                if (next < 0 || header.Length > 20000) throw new IOException("Incomplete fixture request.");
                                header.Append((char)next);
                            }
                            var length = int.Parse(header.ToString().Split(new[] { "\r\n" }, StringSplitOptions.None)
                                .Single(line => line.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase)).Split(':')[1]);
                            var buffer = new byte[4096];
                            while (length > 0)
                            {
                                var read = stream.Read(buffer, 0, Math.Min(length, buffer.Length));
                                if (read == 0) throw new IOException("Incomplete fixture request body.");
                                length -= read;
                            }
                            var body = Encoding.UTF8.GetBytes("{\"choices\":[{\"message\":{\"role\":\"assistant\",\"content\":\"{\\\"approved\\\":true,\\\"findings\\\":[]}\"}}]}");
                            var response = Encoding.ASCII.GetBytes("HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nContent-Length: " + body.Length + "\r\nConnection: close\r\n\r\n");
                            stream.Write(response, 0, response.Length); stream.Write(body, 0, body.Length);
                        }
                    }
                    catch (SocketException) { }
                    catch (ObjectDisposedException) { }
                });
            }
            public void Dispose() { listener.Stop(); worker.GetAwaiter().GetResult(); }
        }
    }

    // These doubles model only the read-before-write boundary. They deliberately
    // interrupt the first native add and do not claim to validate Office rendering.
    public sealed class RecoveryPowerPointApplication
    {
        public RecoveryPresentation ActivePresentation { get; } = new RecoveryPresentation();
    }
    public sealed class RecoveryPresentation
    {
        public RecoveryPageSetup PageSetup { get; } = new RecoveryPageSetup();
        public RecoverySlideCollection Slides { get; } = new RecoverySlideCollection();
        public RecoveryTags Tags { get; } = new RecoveryTags();
    }
    public sealed class RecoveryPageSetup
    {
        public float SlideWidth { get; set; } = 720;
        public float SlideHeight { get; set; } = 540;
    }
    public sealed class RecoverySlideCollection
    {
        public int Count { get; private set; } = 1;
        public int AddAttempts { get; private set; }
        public RecoverySlide this[int index] => new RecoverySlide { SlideID = index };
        public object Add(int index, int layout)
        {
            AddAttempts++; Count++;
            throw new InvalidOperationException("injected native interruption after the slide was added");
        }
    }
    public sealed class RecoverySlide { public int SlideID { get; set; } }
    public sealed class RecoveryTags
    {
        private readonly Dictionary<string, string> values = new Dictionary<string, string>();
        public int Writes { get; private set; }
        public string this[string key] => values.ContainsKey(key) ? values[key] : "";
        public void Add(string key, string value) { values[key] = value; Writes++; }
    }
}
