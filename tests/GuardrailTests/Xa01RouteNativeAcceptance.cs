using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
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
    // Complete fake-response XA01 route through the public workbook/deck hosts.
    // Both Office documents are synthetic, unsaved, and owned by this process.
    internal static class Xa01RouteNativeAcceptance
    {
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window,
            out uint processId);

        private static void Check(bool condition, string message)
        { if (!condition) throw new InvalidOperationException(message); }

        internal static int Run(string reportPath)
        { return RunCore(reportPath, false, false, false, false, false, false, false); }

        internal static int RunFailedWorkbook(string reportPath)
        { return RunCore(reportPath, true, false, false, false, false, false, false); }

        internal static int RunMalformedWorkbook(string reportPath)
        { return RunCore(reportPath, false, true, false, false, false, false, false); }

        internal static int RunCancelled(string reportPath)
        { return RunCore(reportPath, false, false, true, false, false, false, false); }

        internal static int RunTransientRetry(string reportPath)
        { return RunCore(reportPath, false, false, false, true, false, false, false); }

        internal static int RunTransportExhausted(string reportPath)
        { return RunCore(reportPath, false, false, false, false, false, true, false); }

        internal static int RunRejectedReview(string reportPath)
        { return RunCore(reportPath, false, false, false, false, true, false, false); }

        internal static int RunRestartReconcile(string reportPath)
        { return RunCore(reportPath, false, false, false, false, true, false, true); }

        private static int RunCore(string reportPath, bool failWorkbook,
            bool malformedWorkbook, bool cancelAfterRead,
            bool transientRetry, bool rejectedReview,
            bool exhaustedTransport, bool restartReconcile)
        {
            var output = Path.GetDirectoryName(Path.GetFullPath(reportPath));
            Directory.CreateDirectory(output);
            var json = new JavaScriptSerializer { MaxJsonLength = 16000000 };
            var previousFlag = Environment.GetEnvironmentVariable(
                AnalysisDocumentPilot.FeatureFlag);
            var existingExcel = new HashSet<int>(Process.GetProcessesByName(
                "EXCEL").Select(process => process.Id));
            dynamic excel = null, source = null, ledger = null, powerpoint = null,
                draft = null;
            var ownedExcelProcessId = 0;
            var forcedExcelCleanup = false;
            TaskContextManager task = null;
            var stage = "setup";
            var failure = "";
            var terminal = false;
            var failedWorkbookBlocked = false;
            var workbookFailureObserved = false;
            var malformedWorkbookBlocked = false;
            var malformedProposalObserved = false;
            var cancellationObserved = false;
            var transportRetriedIdentically = false;
            var rejectedReviewBlocked = false;
            var reviewRejectionObserved = false;
            var deckWriteUncertain = false;
            var transportFailurePaused = false;
            var restartUncertainDeckBlocked = false;
            var restartExactReconciled = false;
            var recoveryReviewRequests = 0;
            var journalMismatchedSlides = "";
            var journalVolatileSlides = "";
            var sourcePreserved = false;
            var workbookDraft = false;
            var deckDraft = false;
            var workbookFacts = false;
            var slideFacts = false;
            var writesVerified = false;
            var requests = 0;
            var reviewRequests = 0;
            string sourceBefore = null;
            Endpoint endpoint = null;
            AnalysisNativeAcceptance.AnalysisReviewEndpoint reviewer = null;
            ChatToolCall originalDeckCall = null;
            try
            {
                Check(Process.GetProcessesByName("POWERPNT").Length == 0,
                    "NATIVE_POWERPOINT_SESSION_ALREADY_OPEN");
                Environment.SetEnvironmentVariable(
                    AnalysisDocumentPilot.FeatureFlag, "1");
                excel = Activator.CreateInstance(Type.GetTypeFromProgID(
                    "Excel.Application", true));
                uint excelProcessId;
                GetWindowThreadProcessId(new IntPtr((int)excel.Hwnd),
                    out excelProcessId);
                if (existingExcel.Contains((int)excelProcessId))
                {
                    excel = null;
                    throw new InvalidOperationException(
                        "NATIVE_EXCEL_SESSION_NOT_OWNED");
                }
                ownedExcelProcessId = (int)excelProcessId;
                excel.Visible = true;
                source = excel.Workbooks.Add();
                ledger = source.Worksheets[1];
                ledger.Name = "Ledger";
                ledger.Cells[1, 2].Value2 = "Period";
                ledger.Cells[1, 9].Value2 = "RevenueEUR";
                ledger.Cells[1, 10].Value2 = "CostEUR";
                ledger.Cells[2, 2].Value2 = "2026-05";
                ledger.Cells[2, 9].Value2 = 85519d;
                ledger.Cells[2, 10].Value2 = 36702d;
                ledger.Cells[3, 2].Value2 = "2026-06";
                ledger.Cells[3, 9].Value2 = 82992d;
                ledger.Cells[3, 10].Value2 = 36714d;
                sourceBefore = SourceValues(ledger);
                powerpoint = Activator.CreateInstance(Type.GetTypeFromProgID(
                    "PowerPoint.Application", true));
                powerpoint.Visible = -1;

                var prompt = "Analyze the current Ledger and create a new draft worksheet and a verified four-slide PowerPoint deck from the same revenue and cost facts. Preserve the source worksheet and leave both outputs unsaved.";
                var request = DocumentChatRequestFactory.Create(
                    "qwen/qwen3.8-27b", "excel", "", new List<ChatTurn>(),
                    prompt, true);
                Check(request.tools.Any(tool => tool.function.name ==
                    WorkbookToolCatalog.ReadCells) && request.tools.Any(tool =>
                    tool.function.name == WorkbookToolCatalog.WriteDraftSheet) &&
                    request.tools.Any(tool => tool.function.name ==
                    CrossAppToolCatalog.SendToPowerPoint),
                    "XA01_REQUEST_TOOLS_MISSING");
                var checkpoint = Path.Combine(Path.GetTempPath(),
                    "scribble-xa01-" + Guid.NewGuid().ToString("N"));
                task = new TaskContextManager(request, "excel", prompt,
                    new TaskCheckpointStore(checkpoint));
                Func<int, ChatToolCall> proposal = round =>
                {
                    if (round == 0) return new ChatToolCall {
                        id = "xa01-read", type = "function",
                        function = new ChatToolCallFunction {
                            name = WorkbookToolCatalog.ReadCells,
                            arguments = json.Serialize(new {
                                sheet = "Ledger", range = "B1:J3",
                                analysis_binding = new {
                                    period_header = "Period",
                                    metrics = new[] {
                                        new { header = "RevenueEUR", currency = "EUR" },
                                        new { header = "CostEUR", currency = "EUR" }
                                    }
                                }
                            })
                        }
                    };
                    var artifact = task.LoadAnalysis();
                    Check(artifact != null && artifact.Facts.Count == 4,
                        "XA01_TYPED_ANALYSIS_MISSING");
                    var fixture = AnalysisNativeAcceptance.Fixture(artifact);
                    if (round == 1) return new ChatToolCall {
                        id = "xa01-workbook", type = "function",
                        function = new ChatToolCallFunction {
                            name = WorkbookToolCatalog.WriteDraftSheet,
                            arguments = malformedWorkbook ?
                                json.Serialize(new {
                                    analysis_id = artifact.AnalysisId,
                                    title = fixture.Item2.WorkbookTitle,
                                    unexpected = "not in exposed schema"
                                }) : json.Serialize(new {
                                    analysis_id = failWorkbook ?
                                        "stale-analysis-id" : artifact.AnalysisId,
                                    title = fixture.Item2.WorkbookTitle
                                })
                        }
                    };
                    return new ChatToolCall {
                        id = "xa01-deck", type = "function",
                        function = new ChatToolCallFunction {
                            name = CrossAppToolCatalog.SendToPowerPoint,
                            arguments = json.Serialize(
                                AnalysisNativeAcceptance.ModelPlanValue(
                                    json.DeserializeObject(json.Serialize(new {
                                        AnalysisId = artifact.AnalysisId,
                                        WorkbookTitle = fixture.Item2.WorkbookTitle,
                                        Slides = fixture.Item2.Slides
                                    }))))
                        }
                    };
                };
                using (endpoint = new Endpoint(proposal, transientRetry,
                    exhaustedTransport))
                using (reviewer = new AnalysisNativeAcceptance
                    .AnalysisReviewEndpoint(rejectedReview))
                using (var client = new OpenAiCompatibleClient())
                using (var reviewClient = new OpenAiCompatibleClient())
                using (var host = new DocumentDraftHost("excel", (object)excel))
                using (var cancellation = new CancellationTokenSource())
                {
                    var settings = new AppSettings { BaseUrl = endpoint.BaseUrl,
                        ApiKey = "offline-test", Model = request.model };
                    var reviewSettings = new AppSettings {
                        BaseUrl = reviewer.BaseUrl, ApiKey = "offline-test",
                        Model = request.model };
                    host.BindTaskAsync(task, CancellationToken.None)
                        .GetAwaiter().GetResult();
                    var reads = new WorkbookToolHost((object)excel);
                    for (var round = 0; round < 4; round++)
                    {
                        stage = "chat_round_" + round;
                        ChatCompletionResponseMessage response;
                        try
                        {
                            response = task.CompleteAsync(client, settings,
                                request, null, cancelAfterRead ?
                                    cancellation.Token : CancellationToken.None)
                                .GetAwaiter().GetResult();
                        }
                        catch (OperationCanceledException) when (cancelAfterRead)
                        {
                            cancellationObserved = task.State.Lifecycle ==
                                TaskLifecycle.Paused &&
                                task.State.UserPaused == false &&
                                !string.IsNullOrEmpty(task.State.Blocker);
                            Check(cancellationObserved,
                                "XA01_CANCELLATION_NOT_CHECKPOINTED");
                            break;
                        }
                        catch (AiEndpointException error) when (
                            exhaustedTransport && round == 2)
                        {
                            transportFailurePaused = error.HttpStatus == 503 &&
                                task.State.Lifecycle == TaskLifecycle.Paused &&
                                !task.State.UserPaused &&
                                !string.IsNullOrEmpty(task.State.Blocker);
                            Check(transportFailurePaused,
                                "XA01_TRANSPORT_FAILURE_NOT_CHECKPOINTED");
                            break;
                        }
                        if (response.tool_calls == null ||
                            response.tool_calls.Count == 0)
                        {
                            task.State.EnumerationComplete = true;
                            task.Checkpoint();
                            if (failWorkbook || malformedWorkbook ||
                                rejectedReview)
                            {
                                var blocked = !task.State.CanComplete(false) &&
                                    (rejectedReview ?
                                        string.IsNullOrEmpty(task.State.PresentationReviewReceipt) :
                                        string.IsNullOrEmpty(task.State.WorkbookDraftReceipt));
                                if (failWorkbook) failedWorkbookBlocked = blocked;
                                else if (malformedWorkbook)
                                    malformedWorkbookBlocked = blocked;
                                else rejectedReviewBlocked = blocked;
                                Check(blocked,
                                    "MISSING_WORKBOOK_ALLOWED_TERMINAL_COMPLETION");
                                break;
                            }
                            Check(task.State.CanComplete(false),
                                "XA01_TERMINAL_RECEIPT_MISSING");
                            task.CompleteTask(request);
                            terminal = true;
                            break;
                        }
                        task.PrepareExchange(response, request);
                        var results = new List<MailboxToolResult>();
                        foreach (var call in response.tool_calls)
                        {
                            if (round == 2 && rejectedReview)
                                originalDeckCall = call;
                            var definition = request.tools.FirstOrDefault(tool =>
                                tool.function.name == call.function.name);
                            Check(definition != null,
                                "XA01_TOOL_NOT_EXPOSED: " + call.function.name +
                                "; available=" + string.Join(",",
                                    request.tools.Select(tool =>
                                        tool.function.name)));
                            var invalid = task.ValidateArguments(call);
                            if (malformedWorkbook && round == 1)
                            {
                                malformedProposalObserved = invalid != null &&
                                    invalid.Outcome.ErrorCode ==
                                        "TOOL_ARGUMENTS_INVALID" &&
                                    invalid.Outcome.PermissionConsumed == false;
                                Check(malformedProposalObserved,
                                    "MALFORMED_WORKBOOK_PROPOSAL_NOT_REJECTED");
                                results.Add(invalid);
                                continue;
                            }
                            Check(invalid == null,
                                "XA01_UNEXPECTED_ARGUMENT_REJECTION: " +
                                (invalid == null ? "" : invalid.Content));
                            Check(ToolContractValidator.Validate(call,
                                definition).Count == 0,
                                "XA01_TOOL_SCHEMA_INVALID");
                            var write = call.function.name !=
                                WorkbookToolCatalog.ReadCells;
                            task.BeforeTool(call, write);
                            var result = write ? host.ExecuteAsync(call,
                                new OneShotDraftAuthorization(true), true,
                                prompt, reviewClient, reviewSettings,
                                CancellationToken.None, null).GetAwaiter()
                                .GetResult() : reads.Execute(call);
                            if (failWorkbook && round == 1)
                            {
                                workbookFailureObserved =
                                    result.Outcome.Failed &&
                                    result.Outcome.ErrorCode ==
                                        "ANALYSIS_DRAFT_PREFLIGHT_FAILED" &&
                                    result.Outcome.PermissionConsumed == false;
                                Check(workbookFailureObserved,
                                    "EXPECTED_WORKBOOK_PREFLIGHT_FAILURE_MISSING: " +
                                    result.Content);
                            }
                            else if (rejectedReview && round == 2)
                            {
                                reviewRejectionObserved =
                                    result.Outcome.Failed &&
                                    result.Content.Contains(
                                        "ANALYSIS_DECK_REVIEW_REJECTED");
                                Check(reviewRejectionObserved,
                                    "EXPECTED_DECK_REVIEW_REJECTION_MISSING: " +
                                    result.Content);
                            }
                            else Check(!result.Outcome.Failed,
                                "XA01_PUBLIC_TOOL_FAILED: " + result.Content);
                            task.AfterTool(call, result);
                            if (!write)
                                DocumentChatRequestFactory.ApplyAnalysisPilot(
                                    request, task.LoadAnalysis(), "excel");
                            results.Add(result);
                        }
                        request.messages.Add(response);
                        foreach (var result in results)
                            request.messages.Add(
                                new ChatCompletionToolResultMessage {
                                    role = "tool",
                                    tool_call_id = result.ToolCallId,
                                    content = result.Content });
                        task.RecordExchange(request, response, results);
                        if (cancelAfterRead && round == 0)
                            cancellation.Cancel();
                    }
                    requests = endpoint.Count;
                    transportRetriedIdentically = endpoint.RetryIdentical;
                    reviewRequests = reviewer.RequestCount;
                    sourcePreserved = SourceValues(ledger) == sourceBefore;
                    workbookDraft = (int)source.Worksheets.Count == 2 &&
                        Convert.ToString(source.Worksheets[2].Name)
                            .StartsWith("Scribble Draft",
                                StringComparison.Ordinal);
                    if (workbookDraft)
                    {
                        dynamic outputSheet = source.Worksheets[2];
                        workbookFacts = Convert.ToDouble(
                            outputSheet.Range("B4").Value2) == 85519d &&
                            Convert.ToDouble(outputSheet.Range("C4").Value2) == 82992d &&
                            Convert.ToDouble(outputSheet.Range("B5").Value2) == 36702d &&
                            Convert.ToDouble(outputSheet.Range("C5").Value2) == 36714d &&
                            Convert.ToString(outputSheet.Range("B4").Formula)
                                .StartsWith("=SUMIF(",
                                    StringComparison.OrdinalIgnoreCase);
                    }
                    for (var index = 1; index <=
                        (int)powerpoint.Presentations.Count; index++)
                    {
                        dynamic candidate = powerpoint.Presentations[index];
                        if (Convert.ToString(candidate.Tags["ScribbleTask"]) ==
                            task.State.Id) draft = candidate;
                    }
                    deckDraft = (object)draft != null &&
                        (int)draft.Slides.Count == 4 &&
                        string.IsNullOrEmpty(Convert.ToString(draft.Path));
                    if (deckDraft)
                    {
                        var pages = Enumerable.Range(1, 4).Select(index => {
                            var page = PresentationInspection.ReadPage(
                                (object)draft, (object)draft.Slides[index],
                                0, false);
                            return (string)json.Deserialize<Dictionary<string,
                                object>>(json.Serialize(page))["content"];
                        }).ToArray();
                        slideFacts = pages[0].Contains("82,992") &&
                            pages[0].Contains("36,714") &&
                            pages[1].Contains("85,519") &&
                            pages[1].Contains("82,992") &&
                            PresentationInspection.ContainsNativeChart(
                                (object)draft.Slides[2]) &&
                            pages[2].Contains("36,702") &&
                            pages[2].Contains("36,714") &&
                            pages[3].Contains("Verified workbook range");
                    }
                    writesVerified = task.State.Writes.Count ==
                        (cancelAfterRead ? 0 :
                            (malformedWorkbook || exhaustedTransport) ? 1 : 2) &&
                        task.State.Writes.All(write => write.Status == "verified");
                    deckWriteUncertain = task.State.Writes.Count == 2 &&
                        task.State.Writes[0].Status == "verified" &&
                        task.State.Writes[1].Status == "uncertain";
                    if (cancelAfterRead)
                        Check(cancellationObserved && !terminal &&
                            sourcePreserved && !workbookDraft && !deckDraft &&
                            writesVerified && requests == 1 &&
                            reviewRequests == 0 &&
                            string.IsNullOrEmpty(task.State.WorkbookDraftReceipt),
                            "XA01_CANCELLED_ROUTE_INCOMPLETE");
                    else if (malformedWorkbook)
                        Check(malformedWorkbookBlocked &&
                            malformedProposalObserved && !terminal &&
                            sourcePreserved && !workbookDraft &&
                            deckDraft && slideFacts && writesVerified &&
                            requests == 4 && reviewRequests == 5,
                            "XA01_MALFORMED_WORKBOOK_ROUTE_INCOMPLETE");
                    else if (failWorkbook)
                        Check(failedWorkbookBlocked &&
                            workbookFailureObserved && !terminal &&
                            sourcePreserved && !workbookDraft &&
                            deckDraft && slideFacts && writesVerified &&
                            requests == 4 && reviewRequests == 5,
                            "XA01_FAILED_WORKBOOK_ROUTE_INCOMPLETE");
                    else if (rejectedReview)
                    {
                        Check(rejectedReviewBlocked &&
                            reviewRejectionObserved && !terminal &&
                            sourcePreserved && workbookDraft && workbookFacts &&
                            deckWriteUncertain &&
                            !task.State.HostData.ContainsKey(
                                "analysis_deck_complete") &&
                            requests == 4 && reviewRequests == 1,
                            "XA01_REJECTED_REVIEW_ROUTE_INCOMPLETE");
                        stage = "restart_uncertain_deck";
                        var idsBefore = string.Join(",", Enumerable.Range(1,
                            (int)draft.Slides.Count).Select(index =>
                            Convert.ToInt32(draft.Slides[index].SlideID)));
                        var restoredState = task.Store.Load(task.State.Id);
                        Check(restoredState.Writes.Count == 2 &&
                            restoredState.Writes[1].Status == "uncertain",
                            "XA01_RESTART_LOST_UNCERTAIN_WRITE");
                        journalMismatchedSlides = JournalDrift(
                            restoredState, draft, out journalVolatileSlides);
                        var restoredRequest = DocumentChatRequestFactory.Create(
                            request.model, "excel", "", new List<ChatTurn>(),
                            prompt, true);
                        var resumed = new TaskContextManager(restoredRequest,
                            "excel", prompt, task.Store, restoredState);
                        using (var resumedHost = new DocumentDraftHost("excel",
                            (object)excel))
                        {
                            resumedHost.BindTaskAsync(resumed,
                                CancellationToken.None).GetAwaiter().GetResult();
                            var changedCall = new ChatToolCall {
                                id = "xa01-changed-deck-after-restart",
                                type = "function",
                                function = new ChatToolCallFunction {
                                    name = CrossAppToolCatalog.SendToPowerPoint,
                                    arguments = "{}"
                                }
                            };
                            try { resumed.BeforeTool(changedCall, true); }
                            catch (InvalidOperationException error)
                            {
                                restartUncertainDeckBlocked =
                                    error.Message.Contains(
                                        "interrupted document write is uncertain");
                            }
                            var idsAfter = string.Join(",", Enumerable.Range(1,
                                (int)draft.Slides.Count).Select(index =>
                                Convert.ToInt32(draft.Slides[index].SlideID)));
                            Check(restartUncertainDeckBlocked &&
                                !resumed.State.CanComplete(false) &&
                                idsAfter == idsBefore &&
                                SourceValues(ledger) == sourceBefore,
                                "XA01_RESTART_UNCERTAIN_DECK_REPLAYED");
                            if (restartReconcile)
                            {
                                stage = "reconcile_original_deck";
                                Check(originalDeckCall != null,
                                    "XA01_RESTART_ORIGINAL_CALL_MISSING");
                                var exactCall = new ChatToolCall {
                                    id = "xa01-deck-exact-restart",
                                    type = "function",
                                    function = originalDeckCall.function
                                };
                                using (var approval = new
                                    AnalysisNativeAcceptance
                                        .AnalysisReviewEndpoint(
                                            approveImmediately: true))
                                {
                                    var approvalSettings = new AppSettings {
                                        BaseUrl = approval.BaseUrl,
                                        ApiKey = "offline-test",
                                        Model = request.model
                                    };
                                    resumed.BeforeTool(exactCall, true);
                                    var recovered = resumedHost.ExecuteAsync(
                                        exactCall,
                                        new OneShotDraftAuthorization(true),
                                        true, prompt, reviewClient,
                                        approvalSettings,
                                        CancellationToken.None, null)
                                        .GetAwaiter().GetResult();
                                    Check(!recovered.Outcome.Failed,
                                        "XA01_EXACT_RESTART_FAILED: " +
                                        recovered.Content);
                                    resumed.AfterTool(exactCall, recovered);
                                    recoveryReviewRequests =
                                        approval.RequestCount;
                                }
                                var reconciledIds = string.Join(",",
                                    Enumerable.Range(1,
                                        (int)draft.Slides.Count).Select(index =>
                                        Convert.ToInt32(
                                            draft.Slides[index].SlideID)));
                                Check(reconciledIds == idsBefore &&
                                    (int)powerpoint.Presentations.Count == 1 &&
                                    (int)draft.Slides.Count == 4 &&
                                    SourceValues(ledger) == sourceBefore &&
                                    resumed.State.Writes.Count == 3 &&
                                    resumed.State.Writes.All(write =>
                                        write.Status == "verified") &&
                                    !string.IsNullOrEmpty(resumed.State
                                        .PresentationReviewReceipt) &&
                                    recoveryReviewRequests == 1 &&
                                    resumed.State.CanComplete(false),
                                    "XA01_EXACT_RESTART_DID_NOT_RECONCILE");
                                resumed.CompleteTask(restoredRequest);
                                restartExactReconciled =
                                    resumed.State.Lifecycle ==
                                        TaskLifecycle.Completed;
                                terminal = restartExactReconciled;
                                writesVerified = restartExactReconciled;
                            }
                            else resumed.Pause(
                                "Uncertain deck requires native reconciliation.");
                        }
                    }
                    else if (exhaustedTransport)
                        Check(transportFailurePaused &&
                            transportRetriedIdentically && !terminal &&
                            sourcePreserved && workbookDraft && workbookFacts &&
                            !deckDraft && writesVerified &&
                            task.State.Writes.Count == 1 &&
                            string.IsNullOrEmpty(
                                task.State.PresentationReviewReceipt) &&
                            requests == 4 && reviewRequests == 0,
                            "XA01_TRANSPORT_EXHAUSTED_ROUTE_INCOMPLETE");
                    else
                    {
                        Check(terminal && sourcePreserved && workbookDraft &&
                            workbookFacts && deckDraft && slideFacts &&
                            writesVerified &&
                            requests == (transientRetry ? 5 : 4) &&
                            reviewRequests == 5 &&
                            (!transientRetry || transportRetriedIdentically),
                            "XA01_ROUTE_INCOMPLETE");
                        stage = "capture_native_output";
                        draft.SaveCopyAs(Path.Combine(output, "xa01-candidate.pptx"));
                        draft.SaveAs(Path.Combine(output, "xa01-candidate.pdf"),
                            32);
                    }
                }
            }
            catch (Exception error) { failure = stage + ": " + error; }
            finally
            {
                if (endpoint != null) {
                    requests = endpoint.Count;
                    transportRetriedIdentically = endpoint.RetryIdentical;
                }
                if (reviewer != null)
                    reviewRequests = reviewer.RequestCount;
                if ((object)source != null && sourceBefore != null)
                    try { sourcePreserved = SourceValues(source.Worksheets[1]) ==
                        sourceBefore; } catch { }
                if ((object)draft != null) try { draft.Close(); } catch { }
                if ((object)source != null) try { source.Close(false); }
                    catch { }
                if ((object)excel != null) try {
                    excel.Quit(); }
                    catch { }
                if ((object)powerpoint != null) try {
                    if ((int)powerpoint.Presentations.Count == 0)
                        powerpoint.Quit(); }
                    catch { }
                ReleaseCom((object)ledger);
                ReleaseCom((object)source);
                ReleaseCom((object)excel);
                ReleaseCom((object)draft);
                ReleaseCom((object)powerpoint);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                forcedExcelCleanup = StopOwnedExcelIfStillRunning(
                    ownedExcelProcessId, existingExcel);
                Environment.SetEnvironmentVariable(
                    AnalysisDocumentPilot.FeatureFlag, previousFlag);
            }
            var report = new {
                execution_kind = restartReconcile ?
                    "native_fake_endpoint_xa01_restart_reconcile" :
                    exhaustedTransport ?
                    "native_fake_endpoint_xa01_transport_exhausted" :
                    rejectedReview ?
                    "native_fake_endpoint_xa01_rejected_review" :
                    transientRetry ?
                    "native_fake_endpoint_xa01_transport_retry" :
                    cancelAfterRead ?
                    "native_fake_endpoint_xa01_cancelled" :
                    malformedWorkbook ?
                    "native_fake_endpoint_xa01_malformed_workbook" :
                    failWorkbook ? "native_fake_endpoint_xa01_failed_workbook" :
                    "native_fake_endpoint_xa01_route",
                assembly_sha256 = PresentationRevisionAcceptance.AssemblyHash(),
                terminal_receipt_passed = terminal,
                failed_workbook_blocked = failedWorkbookBlocked,
                workbook_failure_observed = workbookFailureObserved,
                malformed_workbook_blocked = malformedWorkbookBlocked,
                malformed_proposal_observed = malformedProposalObserved,
                cancellation_observed = cancellationObserved,
                transport_retried_identically = transportRetriedIdentically,
                rejected_review_blocked = rejectedReviewBlocked,
                review_rejection_observed = reviewRejectionObserved,
                deck_write_uncertain = deckWriteUncertain,
                transport_failure_paused = transportFailurePaused,
                restart_uncertain_deck_blocked =
                    restartUncertainDeckBlocked,
                restart_exact_reconciled = restartExactReconciled,
                recovery_review_requests = recoveryReviewRequests,
                journal_mismatched_slides = journalMismatchedSlides,
                journal_volatile_slides = journalVolatileSlides,
                source_preserved = sourcePreserved,
                workbook_draft_passed = workbookDraft,
                workbook_facts_passed = workbookFacts,
                deck_draft_passed = deckDraft,
                slide_facts_passed = slideFacts,
                writes_verified = writesVerified,
                model_requests = requests,
                review_requests = reviewRequests,
                paid_model_calls = 0,
                test_owned_excel_forced_cleanup = forcedExcelCleanup,
                full_acceptance_passed = false,
                failure
            };
            File.WriteAllText(reportPath, json.Serialize(report));
            Console.WriteLine(json.Serialize(report));
            return failure.Length == 0 ? 0 : 1;
        }

        private static string SourceValues(dynamic sheet)
        {
            return string.Join("|", new[] { "B1", "I1", "J1", "B2",
                "I2", "J2", "B3", "I3", "J3" }.Select(address =>
                Convert.ToString(sheet.Range(address).Value2,
                    System.Globalization.CultureInfo.InvariantCulture)));
        }

        private static string JournalDrift(DurableTaskState state,
            dynamic draft, out string volatileSlides)
        {
            string saved;
            Check(state.HostData.TryGetValue("samsung_pending", out saved),
                "XA01_RESTART_JOURNAL_MISSING");
            var json = new JavaScriptSerializer { MaxJsonLength = 16000000 };
            var journal = json.Deserialize<Dictionary<string, object>>(saved);
            var receipts = (IList)journal["Receipts"];
            var method = typeof(PresentationInspection).GetMethod(
                "FingerprintForJournal", BindingFlags.Static |
                    BindingFlags.NonPublic);
            Check(method != null, "XA01_RESTART_FINGERPRINT_UNAVAILABLE");
            var mismatch = new List<string>();
            var unstable = new List<string>();
            foreach (Dictionary<string, object> receipt in receipts)
            {
                var slideId = Convert.ToInt32(receipt["SlideId"]);
                object slide = null;
                for (var index = 1; index <= (int)draft.Slides.Count; index++)
                    if ((int)draft.Slides[index].SlideID == slideId)
                        slide = (object)draft.Slides[index];
                Check(slide != null, "XA01_RESTART_SLIDE_MISSING");
                var first = (string)method.Invoke(null, new[] { slide });
                var second = (string)method.Invoke(null, new[] { slide });
                if (first != (string)receipt["Fingerprint"])
                    mismatch.Add(slideId.ToString());
                if (first != second) unstable.Add(slideId.ToString());
            }
            volatileSlides = string.Join(",", unstable);
            return string.Join(",", mismatch);
        }

        private static void ReleaseCom(object value)
        {
            try { if (value != null && Marshal.IsComObject(value))
                Marshal.FinalReleaseComObject(value); }
            catch (InvalidComObjectException) { }
            catch (COMException) { }
        }

        private static bool StopOwnedExcelIfStillRunning(int processId,
            ISet<int> preexisting)
        {
            if (processId <= 0 || preexisting.Contains(processId)) return false;
            try
            {
                using (var process = Process.GetProcessById(processId))
                {
                    if (!string.Equals(process.ProcessName, "EXCEL",
                        StringComparison.OrdinalIgnoreCase)) return false;
                    if (process.WaitForExit(2500)) return false;
                    process.Kill();
                    Check(process.WaitForExit(5000),
                        "TEST_OWNED_EXCEL_CLEANUP_FAILED");
                    return true;
                }
            }
            catch (ArgumentException) { return false; }
        }

        private sealed class Endpoint : IDisposable
        {
            private readonly TcpListener _listener = new TcpListener(
                IPAddress.Loopback, 0);
            private readonly Task _worker;
            private readonly Func<int, ChatToolCall> _proposal;
            private readonly bool _transientRetry;
            private readonly bool _exhaustedTransport;
            private int _round;
            private string _retryBody;
            internal int Count;
            internal bool RetryIdentical;
            internal string BaseUrl { get; }
            internal Endpoint(Func<int, ChatToolCall> proposal,
                bool transientRetry, bool exhaustedTransport)
            {
                _proposal = proposal;
                _transientRetry = transientRetry;
                _exhaustedTransport = exhaustedTransport;
                _listener.Start();
                BaseUrl = "http://127.0.0.1:" +
                    ((IPEndPoint)_listener.LocalEndpoint).Port + "/v1";
                _worker = Task.Run((Action)Serve);
            }
            private void Serve()
            {
                try { while (true) using (var connection =
                    _listener.AcceptTcpClient()) Respond(connection); }
                catch (SocketException) { }
                catch (ObjectDisposedException) { }
            }
            private void Respond(TcpClient connection)
            {
                using (var stream = connection.GetStream())
                {
                    stream.ReadTimeout = 30000;
                    var header = new StringBuilder();
                    while (!header.ToString().EndsWith("\r\n\r\n",
                        StringComparison.Ordinal))
                    {
                        var value = stream.ReadByte();
                        Check(value >= 0 && header.Length < 16384,
                            "XA01_HTTP_HEADER_INVALID");
                        header.Append((char)value);
                    }
                    var length = int.Parse(Regex.Match(header.ToString(),
                        @"Content-Length:\s*(\d+)", RegexOptions.IgnoreCase)
                        .Groups[1].Value);
                    Check(length > 0 && length < 16000000,
                        "XA01_HTTP_BODY_INVALID");
                    var bytes = new byte[length];
                    for (var offset = 0; offset < bytes.Length;)
                    {
                        var read = stream.Read(bytes, offset,
                            bytes.Length - offset);
                        Check(read > 0, "XA01_HTTP_BODY_TRUNCATED");
                        offset += read;
                    }
                    var json = new JavaScriptSerializer {
                        MaxJsonLength = 16000000 };
                    var requestBody = Encoding.UTF8.GetString(bytes);
                    json.Deserialize<Dictionary<string, object>>(requestBody);
                    var attempt = Count++;
                    Check(attempt < (_transientRetry ? 5 : 4),
                        "XA01_MODEL_CALL_LIMIT");
                    if ((_transientRetry || _exhaustedTransport) &&
                        _round == 2 && _retryBody == null &&
                        !RetryIdentical)
                    {
                        _retryBody = requestBody;
                        WriteTransientFailure(stream);
                        return;
                    }
                    if (_retryBody != null)
                    {
                        RetryIdentical = requestBody == _retryBody;
                        Check(RetryIdentical,
                            "XA01_TRANSIENT_RETRY_BODY_CHANGED");
                        _retryBody = null;
                        if (_exhaustedTransport)
                        {
                            WriteTransientFailure(stream);
                            return;
                        }
                    }
                    var round = _round++;
                    Check(round < 4, "XA01_MODEL_RESPONSE_LIMIT");
                    object message = round < 3
                        ? (object)new { role = "assistant",
                            tool_calls = new[] { _proposal(round) } }
                        : new { role = "assistant", content =
                            "The verified workbook and four-slide deck are ready for review; nothing was saved." };
                    var response = Encoding.UTF8.GetBytes(json.Serialize(
                        new { choices = new[] { new { index = 0, message,
                            finish_reason = "stop" } } }));
                    var responseHeader = Encoding.ASCII.GetBytes(
                        "HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nContent-Length: " +
                        response.Length + "\r\nConnection: close\r\n\r\n");
                    stream.Write(responseHeader, 0,
                        responseHeader.Length);
                    stream.Write(response, 0, response.Length);
                }
            }
            private static void WriteTransientFailure(NetworkStream stream)
            {
                var failure = Encoding.UTF8.GetBytes(
                    "{\"error\":{\"message\":\"synthetic transient\"}}");
                var header = Encoding.ASCII.GetBytes(
                    "HTTP/1.1 503 Service Unavailable\r\n" +
                    "Content-Type: application/json\r\n" +
                    "Retry-After: 0\r\nContent-Length: " +
                    failure.Length + "\r\nConnection: close\r\n\r\n");
                stream.Write(header, 0, header.Length);
                stream.Write(failure, 0, failure.Length);
            }
            public void Dispose()
            { _listener.Stop(); try { _worker.Wait(1000); } catch { } }
        }
    }
}
