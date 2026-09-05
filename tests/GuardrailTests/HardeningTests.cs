using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Outlook;
using Scribble.Utilities;

namespace GuardrailTests
{
    internal static class HardeningTests
    {
        private static void Check(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
        private static ChatToolCall Call(string name, string args) { return new ChatToolCall { id = Guid.NewGuid().ToString("N"), type = "function", function = new ChatToolCallFunction { name = name, arguments = args } }; }
        private static ChatCompletionRequest Request() { return new ChatCompletionRequest { model = "synthetic", messages = new List<object>() }; }

        public static void SparseMailboxThroughCoordinator()
        {
            foreach (var rejectFilter in new[] { false, true })
            {
                var root = Path.Combine(Path.GetTempPath(), "scribble-sparse-" + Guid.NewGuid().ToString("N"));
                try
                {
                    var from = DateTimeOffset.Parse("2026-09-04T17:00:13+04:00");
                    var to = DateTimeOffset.Parse("2026-09-05T14:12:17+04:00");
                    var app = new FakeOutlookApplication();
                    var items = Enumerable.Range(0, 950).Select(i => new FakeSelectedMailItem("read-" + i, "Already read", to.LocalDateTime.AddSeconds(-i), false)).ToList();
                    items.Add(new FakeSelectedMailItem("lower", "Inclusive lower", from.LocalDateTime, true));
                    items.Add(new FakeSelectedMailItem("upper", "Inclusive upper", to.LocalDateTime, true));
                    items.Add(new FakeSelectedMailItem("outside-before", "Outside", from.LocalDateTime.AddSeconds(-1), true));
                    items.Add(new FakeSelectedMailItem("outside-after", "Outside", to.LocalDateTime.AddSeconds(1), true));
                    var folder = new FakeMailFolder { Items = new FakeSearchItems(items), RejectDateFilter = rejectFilter };
                    app.Session.RegisterFolder(6, folder);
                    var request = Request();
                    var task = new TaskContextManager(request, "outlook", "Summarize unread messages", new TaskCheckpointStore(root));
                    var json = new JavaScriptSerializer();
                    var subjects = new List<string>();
                    string cursor = "";
                    long scanned = 0;
                    using (var host = new MailboxToolHost(app, null))
                    {
                        host.BindTaskAsync(task, CancellationToken.None).GetAwaiter().GetResult();
                        for (var pageNumber = 0; ; pageNumber++)
                        {
                            Check(pageNumber < 30, "Sparse cursor never completed.");
                            var call = Call(MailboxToolCatalog.SearchMailbox, cursor.Length == 0 ? json.Serialize(new { query = "", folder = "inbox", unread_only = true,
                                received_after = from.ToString("O"), received_before = to.ToString("O"), max_results = 1 }) : json.Serialize(new { cursor, max_results = 1 }));
                            var response = new ChatCompletionResponseMessage { tool_calls = new List<ChatToolCall> { call } };
                            task.PrepareExchange(response, request);
                            task.BeforeTool(call, false);
                            var result = host.ExecuteAsync(call, CancellationToken.None).GetAwaiter().GetResult();
                            Check(!result.Outcome.Failed, result.Content);
                            task.AfterTool(call, result);
                            task.RecordExchange(request, response, new List<MailboxToolResult> { result });
                            task.FinishExchange(request);
                            var data = json.Deserialize<Dictionary<string, object>>(result.Content);
                            var progress = (Dictionary<string, object>)data["progress"];
                            Check(Convert.ToInt64(progress["scanned_rows"]) >= scanned, "Scan position went backwards.");
                            scanned = Convert.ToInt64(progress["scanned_rows"]);
                            Check(Convert.ToBoolean(data["unread_only"]), "Continuation lost the unread filter.");
                            Check(DateTimeOffset.Parse((string)data["received_after"]) == from, "Continuation lost the original time boundary.");
                            foreach (var raw in (System.Collections.IEnumerable)data["results"]) subjects.Add((string)((Dictionary<string, object>)raw)["subject"]);
                            cursor = (string)data["next_cursor"];
                            if (cursor.Length == 0) break;
                        }
                    }
                    Check(subjects.OrderBy(s => s).SequenceEqual(new[] { "Inclusive lower", "Inclusive upper" }), "Sparse coverage or inclusive timestamps were wrong.");
                    Check(scanned == 954, "Did not enumerate the complete sparse fixture.");
                    Check(folder.Filters[0].Contains("2026-09-04 13:00") && folder.Filters[0].Contains("2026-09-05 10:13"), "Provider filter did not widen UTC minute boundaries.");
                    if (rejectFilter) Check(folder.Filters.Count == 2 && folder.Filters[1] == "", "Unsupported restrictions did not fall back to a complete scan.");
                }
                finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
            }
        }

        public static void TypedOutcomesAndDiagnostics()
        {
            Check(!ToolOutcome.Parse("{\"messages\":[{\"body\":\"example: error_code\",\"error_code\":\"quoted source\"}]}").Failed, "Nested source content became a host error.");
            Check(!ToolOutcome.Parse("Article about \"error_code\" and [TASK_FAILED]").Failed, "Page text became an error.");
            Check(ToolOutcome.Parse("{\"error_code\":\"SLIDE_FAILED\",\"permission_consumed\": false}").PermissionConsumed == false, "Whitespace changed permission classification.");
            Check(ToolOutcome.Parse("[WEB_FETCH_HTTP_404] Missing").Failed, "Legacy fetch failure was ignored.");
            var recorder = new DiagnosticsRecorder(); recorder.BeginRequest("test", "model", false);
            recorder.BindTask("permanent-task-id", "effective-vision-model");
            for (var i = 1; i <= 160; i++) recorder.RecordEvent("Event " + i);
            var report = recorder.BuildReport("test");
            Check(report.Contains("Local diagnostic ID: permanent-task-id") && report.Contains("Model: effective-vision-model"), "Diagnostic identity or effective model fell out of the event ring.");
            Check(report.Contains("Event 160") && report.Contains("Events: 160; retained tail: 128") && !report.Contains("  Event 1\r\n"), "Diagnostics lost the recent failure tail.");
            var root = Path.Combine(Path.GetTempPath(), "scribble-flight-" + Guid.NewGuid().ToString("N"));
            try
            {
                var request = Request();
                var task = new TaskContextManager(request, "test", "Private request", new TaskCheckpointStore(root));
                for (var i = 0; i < 160; i++) task.Diagnostics.Record("step", new { i, source = "Private source" });
                task.Pause("Failure after step 160");
                var replay = task.Diagnostics.ReadLocalReplay();
                Check(replay.Count == 128 && replay.Last().Stage == "task_paused" && replay.Last().Detail.Contains("step 160"), "Replay lost its failure context.");
                Check(!task.Diagnostics.RedactedReport().Contains("Private") && !task.Diagnostics.RedactedReport().Contains("step 160"), "Redacted export leaked content.");
                Check(Directory.GetFiles(Path.Combine(root, task.State.Id), "diagnostic-*.dat").All(p => !System.Text.Encoding.UTF8.GetString(File.ReadAllBytes(p)).Contains("Private source")), "Trace was stored in plaintext.");
                Check(!new JavaScriptSerializer().Serialize(request).Contains("Diagnostics"), "Local diagnostics leaked into model request schema.");
            }
            finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        }

        public static void WebCacheAndRedirects()
        {
            var root = Path.Combine(Path.GetTempPath(), "scribble-web-" + Guid.NewGuid().ToString("N"));
            try
            {
                var task = new TaskContextManager(Request(), "word", "Research", new TaskCheckpointStore(root));
                var handler = new WebFixture();
                using (var client = new HttpClient(handler))
                {
                    var first = WebReadTool.Execute(Call(WebReadTool.FetchWebPage, "{\"url\":\"https://example.com/old#section\"}"), task, client);
                    Check(first.Content.Contains("https://example.com/new/child") && first.Content.Contains("Final page"), "Relative links were not resolved against the final response URL.");
                    for (var i = 0; i < 6; i++) Check(WebReadTool.Execute(Call(WebReadTool.FetchWebPage, "{\"url\":\"https://example.com/old\"}"), task, client).Content.Contains("\"cached\":true"), "Repeated fetch was not a source receipt.");
                    Check(handler.Calls == 1, "Repeated reads made extra network requests.");
                    var missing = WebReadTool.Execute(Call(WebReadTool.FetchWebPage, "{\"url\":\"https://example.com/missing\"}"), task, client);
                    Check(missing.Outcome.Failed && missing.Content.Contains("Source not found") && !missing.Content.Contains("blocks automated"), "404 was misclassified as bot blocking.");
                }
            }
            finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        }

        public static void SourceSpansAndContracts()
        {
            var root = Path.Combine(Path.GetTempPath(), "scribble-sources-" + Guid.NewGuid().ToString("N"));
            try
            {
                var request = Request(); request.tools = new List<ChatToolDefinition> { PresentationToolCatalog.DraftDefinition() };
                var task = new TaskContextManager(request, "powerpoint", "Create five slides on launch", new TaskCheckpointStore(root));
                task.State.EnumerationComplete = true;
                Check(!task.State.CanComplete(false), "A requested five-slide deck completed before any slide tool ran.");
                task.Sources.Add("Source A", "Sales were 10 units.\nLaunch review.");
                task.Sources.Add("Source B", "Delivery begins on Monday.");
                var spans = task.Sources.Spans();
                var evidence = task.Sources.Resolve(spans.Select(s => s.Id));
                Check(evidence.Contains("10 units") && evidence.Contains("Monday"), "Multiple host-issued sources could not be resolved.");
                var rejected = false;
                try { task.Sources.Resolve(new[] { "fabricated-span" }); } catch (InvalidOperationException) { rejected = true; }
                Check(rejected, "An invented source span was trusted.");
                var valid = Call(PresentationToolCatalog.AddDraftSlides, "{\"plan\":[\"a\"],\"slides\":\"[{\\\"id\\\":\\\"a\\\",\\\"title\\\":\\\"Unicode € 한글\\\"}]\"}");
                Check(task.ValidateArguments(valid) == null && !valid.function.arguments.Contains("\\\"id\\\""), "Known encoded arrays did not normalize before validation.");
                var invalid = Call(PresentationToolCatalog.AddDraftSlides, "{\"slides\":[{\"title\":false}],\"plan\":[\"a\"]}");
                var error = task.ValidateArguments(invalid);
                Check(error != null && error.Outcome.PermissionConsumed == false && error.Content.Contains("$.slides[0].title"), "Malformed slide fields reached the write boundary.");
                Check(task.State.Writes.Count == 0, "Schema rejection consumed a write.");
            }
            finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        }
        private sealed class WebFixture : HttpMessageHandler
        {
            public int Calls;
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
            {
                Calls++;
                return Task.FromResult(new HttpResponseMessage(request.RequestUri.AbsolutePath == "/missing" ? HttpStatusCode.NotFound : HttpStatusCode.OK)
                { RequestMessage = new HttpRequestMessage(HttpMethod.Get, "https://example.com/new/index"),
                    Content = new StringContent("<title>Final page</title><p>Verified source</p><a href='child'>Child</a>") });
            }
        }
    }
}
