using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Configuration;
using Scribble.Office;
using Scribble.Security;

namespace GuardrailTests
{
    // Exercises the request factory, HTTP transport, task journal, public tool
    // host and completion receipt. The fake model only proposes fixture edits;
    // all execution and checks are production code. No real model is contacted.
    internal static class PilotRouteNativeAcceptance
    {
        internal static void FakeTransportSeparatesChatAndReview()
        {
            var repair = new ChatToolCall { id = "repair", type = "function", function = new ChatToolCallFunction {
                name = PresentationToolCatalog.ReviseSlides, arguments = "{}" } };
            using (var endpoint = new Endpoint(repair))
            using (var client = new OpenAiCompatibleClient())
            {
                var settings = new AppSettings { BaseUrl = endpoint.BaseUrl, ApiKey = "offline-test", Model = "qwen/qwen3.8-27b" };
                var request = new ChatCompletionRequest { model = settings.Model, max_tokens = 1024,
                    tools = PresentationToolCatalog.CreateDefinitions().ToList(),
                    messages = new List<object> { new ChatCompletionInputMessage { role = "user", content = "Offline transport test" } } };
                var inspected = client.CompleteAsync(settings, request, CancellationToken.None).GetAwaiter().GetResult();
                Check(inspected.tool_calls.Count == 6, "FAKE_INSPECTION_RESPONSE_INVALID");
                var revised = client.CompleteAsync(settings, request, CancellationToken.None).GetAwaiter().GetResult();
                Check(revised.tool_calls.Single().function.name == PresentationToolCatalog.ReviseSlides, "FAKE_REVISION_RESPONSE_INVALID");
                request.tools = null;
                var review = client.CompleteAsync(settings, request, CancellationToken.None).GetAwaiter().GetResult();
                Check((review.RawContent ?? review.content).Contains("\"approved\":true"), "FAKE_REVIEW_RESPONSE_INVALID");
                request.tools = PresentationToolCatalog.CreateDefinitions().ToList();
                var final = client.CompleteAsync(settings, request, CancellationToken.None).GetAwaiter().GetResult();
                Check(final.content.Contains("nothing was saved") && endpoint.Count == 4, "FAKE_TERMINAL_RESPONSE_INVALID");
            }
        }

        internal static int Run(string sourcePath, string workbookPath, string reportPath)
        {
            var output = Path.GetDirectoryName(Path.GetFullPath(reportPath));
            Directory.CreateDirectory(output);
            var json = new JavaScriptSerializer { MaxJsonLength = 16000000 };
            var sourceHash = ExternalContextDocument.FingerprintFile(sourcePath);
            var workbookHash = ExternalContextDocument.FingerprintFile(workbookPath);
            var priorFlag = Environment.GetEnvironmentVariable(AnalysisDocumentPilot.FeatureFlag);
            dynamic app = null, source = null, draft = null;
            var passed = false;
            var failure = "";
            var stage = "setup";
            var requestCount = 0;
            var promptCharacters = 0;
            var terminal = false;
            var inspected = 0;
            try
            {
                Environment.SetEnvironmentVariable(AnalysisDocumentPilot.FeatureFlag, "1");
                Check(PresentationRevisionAcceptance.Enabled, "NATIVE_RECEIPT_REQUIRED");
                app = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application", true));
                app.Visible = -1;
                source = app.Presentations.Open(sourcePath, -1, 0, -1);
                source.Activate();
                var prompt = "Repair and improve the layout of all six slides as a new editable draft using the attached workbook; preserve the original slides and keep the source deck unchanged. Retain the source content, notes and owner/due-date pairs.";
                var documents = new[] { new ExternalContextDocument("WB01.xlsx",
                    json.Serialize(WorkbookMonthlyChartFacts.ReadSalesLedger(workbookPath, CancellationToken.None)), workbookPath) };
                var request = DocumentChatRequestFactory.Create("qwen/qwen3.8-27b", "powerpoint", "", new List<ChatTurn>(), prompt, true, documents);
                Check(request.tools.Any(tool => tool.function.name == PresentationToolCatalog.ReviseSlides) &&
                    !request.tools.Any(tool => tool.function.name == PresentationToolCatalog.AddDraftSlides), "PILOT_TOOL_ROUTE_INCORRECT");
                var task = new TaskContextManager(request, "powerpoint", prompt,
                    new TaskCheckpointStore(Path.Combine(output, "checkpoint")));
                new TaskRecoveryInput { Prompt = prompt, Documents = documents.Select(document => new SavedReference {
                    Name = document.Name, Content = document.Content, SourcePath = document.SourcePath,
                    SourceFingerprint = document.SourceFingerprint }).ToList() }.PersistTo(task.State);
                task.Checkpoint();
                var proposed = Replacement((object)source, json);
                using (var endpoint = new Endpoint(proposed))
                using (var client = new OpenAiCompatibleClient())
                using (var host = new DocumentDraftHost("powerpoint", (object)app))
                {
                    var settings = new AppSettings { BaseUrl = endpoint.BaseUrl, ApiKey = "offline-test", Model = request.model };
                    host.BindTaskAsync(task, CancellationToken.None).GetAwaiter().GetResult();
                    var reads = new PresentationToolHost((object)app);
                    var authorization = new OneShotDraftAuthorization(true);
                    for (var round = 0; round < 4; round++)
                    {
                        stage = "chat_round_" + round;
                        var response = task.CompleteAsync(client, settings, request, null, CancellationToken.None).GetAwaiter().GetResult();
                        if (response.tool_calls == null || response.tool_calls.Count == 0)
                        {
                            task.State.EnumerationComplete = true;
                            Check(task.State.CanComplete(false), "TASK_TERMINAL_RECEIPT_MISSING");
                            task.CompleteTask(request);
                            terminal = true;
                            break;
                        }
                        task.PrepareExchange(response, request);
                        var results = new List<MailboxToolResult>();
                        foreach (var call in response.tool_calls)
                        {
                            var definition = request.tools.Single(tool => tool.function.name == call.function.name);
                            Check(ToolContractValidator.Validate(call, definition).Count == 0, "MODEL_TOOL_SCHEMA_INVALID");
                            var write = call.function.name == PresentationToolCatalog.ReviseSlides;
                            task.BeforeTool(call, write);
                            MailboxToolResult result;
                            if (write)
                            {
                                stage = "public_revise_slides";
                                Check(response.tool_calls.Count == 1 && inspected == 6, "WRITE_WITHOUT_COMPLETE_INSPECTION");
                                result = host.ExecuteAsync(call, authorization, true, prompt, client, settings,
                                    CancellationToken.None, null).GetAwaiter().GetResult();
                            }
                            else
                            {
                                result = reads.Execute(call);
                                inspected++;
                            }
                            Check(!result.Outcome.Failed, "PRODUCTION_TOOL_FAILED: " + result.Content);
                            task.AfterTool(call, result);
                            results.Add(result);
                        }
                        request.messages.Add(response);
                        foreach (var result in results) request.messages.Add(new ChatCompletionToolResultMessage {
                            role = "tool", tool_call_id = result.ToolCallId, content = result.Content });
                        task.RecordExchange(request, response, results);
                    }
                    requestCount = endpoint.Count;
                    promptCharacters = endpoint.PromptCharacters;
                    Check(terminal && authorization.IsCreated && task.State.Writes.All(write => write.Status == "verified"), "ROUTE_NOT_COMPLETED");
                    Check(requestCount <= 18, "ROUTE_CALL_BUDGET_EXCEEDED");
                    Check(task.State.HostData["delivery_request_count"] == requestCount.ToString(), "TRANSPORT_REQUEST_ACCOUNTING_MISMATCH");
                    for (var index = 1; index <= (int)app.Presentations.Count; index++)
                    {
                        dynamic candidate = app.Presentations[index];
                        if (Convert.ToString(candidate.Tags["ScribbleRevisionDraft"]) == task.State.Id) draft = candidate;
                    }
                    Check((object)draft != null && (int)draft.Slides.Count == 6 && string.IsNullOrEmpty(Convert.ToString(draft.Path)), "OWNED_UNSAVED_DRAFT_MISSING");
                    Check((int)source.Slides.Count == 6, "SOURCE_SLIDES_CHANGED");
                    stage = "capture_native_output";
                    draft.SaveCopyAs(Path.Combine(output, "candidate.pptx"));
                    draft.SaveAs(Path.Combine(output, "candidate.pdf"), 32);
                    passed = true;
                }
            }
            catch (Exception error) { failure = stage + ": " + error; }
            finally
            {
                if ((object)draft != null) try { draft.Close(); } catch { }
                if ((object)source != null) try { source.Close(); } catch { }
                if ((object)app != null) try { if ((int)app.Presentations.Count == 0) app.Quit(); } catch { }
                Environment.SetEnvironmentVariable(AnalysisDocumentPilot.FeatureFlag, priorFlag);
            }
            var sourcePreserved = ExternalContextDocument.FingerprintFile(sourcePath) == sourceHash;
            var workbookPreserved = ExternalContextDocument.FingerprintFile(workbookPath) == workbookHash;
            var report = new { execution_kind = "native_fake_endpoint_production_route", assembly_sha256 = PresentationRevisionAcceptance.AssemblyHash(),
                production_route_passed = passed, terminal_receipt_passed = terminal, inspected_slides = inspected,
                source_preserved = sourcePreserved, workbook_preserved = workbookPreserved,
                model_requests = requestCount, request_characters = promptCharacters, paid_model_calls = 0,
                visual_approved = false, full_acceptance_passed = false, failure };
            File.WriteAllText(reportPath, json.Serialize(report));
            Console.WriteLine(json.Serialize(report));
            return passed && sourcePreserved && workbookPreserved ? 0 : 1;
        }

        private static ChatToolCall Replacement(object sourceDeck, JavaScriptSerializer json)
        {
            dynamic source = sourceDeck;
            dynamic page = source.Slides[4];
            string text = Enumerable.Range(1, (int)page.Shapes.Count).Select(index => (object)page.Shapes[index])
                .Where(shape => (int)((dynamic)shape).HasTextFrame != 0).Select(shape => (string)((dynamic)shape).TextFrame.TextRange.Text)
                .Single(value => value.Contains("The monthly comparison covers"));
            string[] paragraphs = Regex.Split(text, @"(?:\r\n|\r|\n){2,}").Where(value => !string.IsNullOrWhiteSpace(value)).ToArray();
            Check(paragraphs.Length == 4, "FIXTURE_PARAGRAPHS_INVALID");
            var labels = new[] { "June measure", "Cost and margin", "Comparison scope", "Interpretation" };
            return new ChatToolCall { id = "repair", type = "function", function = new ChatToolCallFunction {
                name = PresentationToolCatalog.ReviseSlides, arguments = json.Serialize(new {
                    presentation_id = PresentationInspection.IdentityFor(sourceDeck), operations = new[] { new {
                        kind = "replace_slide", slide_id = (int)page.SlideID, fingerprint = PresentationInspection.Fingerprint((object)page),
                        slide = new { title = "Operating review and evidence boundaries", subtitle = "Source-backed measures and interpretation limits",
                            layout = "cards", cards = paragraphs.Select((paragraph, index) => new { heading = labels[index], points = new[] { paragraph } }).ToArray(),
                            sources = "WB01 Ledger and History", footnote = "Fictional operational source", evidence = text }
                    } } }) } };
        }

        private static void Check(bool condition, string message)
        { if (!condition) throw new InvalidOperationException(message); }

        private sealed class Endpoint : IDisposable
        {
            private readonly TcpListener _listener = new TcpListener(IPAddress.Loopback, 0);
            private readonly Task _worker;
            private readonly ChatToolCall _repair;
            private int _chatRound;
            internal int Count, PromptCharacters;
            internal string BaseUrl { get; }
            internal Endpoint(ChatToolCall repair)
            {
                _repair = repair;
                _listener.Start();
                BaseUrl = "http://127.0.0.1:" + ((IPEndPoint)_listener.LocalEndpoint).Port + "/v1";
                _worker = Task.Run((Action)Serve);
            }
            private void Serve()
            {
                try { while (true) using (var connection = _listener.AcceptTcpClient()) Respond(connection); }
                catch (SocketException) { }
                catch (ObjectDisposedException) { }
            }
            private void Respond(TcpClient connection)
            {
                using (var stream = connection.GetStream())
                {
                    stream.ReadTimeout = 30000;
                    var header = new StringBuilder();
                    while (!header.ToString().EndsWith("\r\n\r\n", StringComparison.Ordinal))
                    {
                        var value = stream.ReadByte();
                        Check(value >= 0 && header.Length < 16384, "HTTP_HEADER_INVALID");
                        header.Append((char)value);
                    }
                    var length = int.Parse(Regex.Match(header.ToString(), @"Content-Length:\s*(\d+)", RegexOptions.IgnoreCase).Groups[1].Value);
                    Check(length > 0 && length < 16000000, "HTTP_BODY_INVALID");
                    var bytes = new byte[length];
                    for (var offset = 0; offset < bytes.Length;)
                    { var read = stream.Read(bytes, offset, bytes.Length - offset); Check(read > 0, "HTTP_BODY_TRUNCATED"); offset += read; }
                    var json = new JavaScriptSerializer { MaxJsonLength = 16000000 };
                    var body = json.Deserialize<Dictionary<string, object>>(Encoding.UTF8.GetString(bytes));
                    Count++; PromptCharacters += length;
                    Check(Count <= 18, "MODEL_CALL_LIMIT");
                    object message;
                    object toolValue;
                    if (body.TryGetValue("tools", out toolValue) &&
                        toolValue is System.Collections.IList &&
                        ((System.Collections.IList)toolValue).Count > 0)
                    {
                        if (_chatRound++ == 0) message = new { role = "assistant", tool_calls = Enumerable.Range(1, 6).Select(index => new ChatToolCall {
                            id = "inspect-" + index, type = "function", function = new ChatToolCallFunction { name = PresentationToolCatalog.InspectSlide,
                                arguments = json.Serialize(new { index, preview = false }) } }).ToArray() };
                        else if (_chatRound == 2) message = new { role = "assistant", tool_calls = new[] { _repair } };
                        else message = new { role = "assistant", content = "The six-slide draft is ready for visual review. Sources remain unchanged; nothing was saved." };
                    }
                    else message = new { role = "assistant", content = "{\"approved\":true,\"issues\":\"\",\"findings\":[]}" };
                    var response = Encoding.UTF8.GetBytes(json.Serialize(new { choices = new[] { new { index = 0, message, finish_reason = "stop" } } }));
                    var responseHeader = Encoding.ASCII.GetBytes("HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nContent-Length: " + response.Length + "\r\nConnection: close\r\n\r\n");
                    stream.Write(responseHeader, 0, responseHeader.Length);
                    stream.Write(response, 0, response.Length);
                }
            }
            public void Dispose()
            { _listener.Stop(); try { _worker.Wait(1000); } catch { } }
        }
    }
}
