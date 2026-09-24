using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Xml.Linq;
using Scribble.Chat;
using Scribble.Office;
using Scribble.Security;

namespace GuardrailTests
{
    // Explicit opt-in. All Office destinations are created by this process,
    // remain unsaved, and are closed without touching an existing document.
    internal static class AnalysisNativeAcceptance
    {
        private static void Check(bool condition, string message)
        { if (!condition) throw new InvalidOperationException(message); }

        internal static int Run(string reportPath)
        {
            dynamic excel = null, workbook = null, powerPoint = null,
                deck = null, typedDeck = null, faultDeck = null,
                recoveryDeck = null;
            var failure = string.Empty;
            var workbookPassed = false;
            var slidesPassed = false;
            var sourcePreserved = false;
            var recoveryPassed = false;
            var typedReviewPassed = false;
            var rendererRepairPassed = false;
            var contentRecoveryPassed = false;
            var partialContentRollbackPassed = false;
            var cardContentRecoveryPassed = false;
            var layoutRecoveryPassed = false;
            var partialLayoutRecoveryPassed = false;
            var typedDeckHandoffPassed = false;
            var nativeDateColumnPassed = false;
            var powerpointExited = false;
            string taskOwner = null;
            var images = new List<string>();
            var stage = "setup";
            var output = Path.GetDirectoryName(Path.GetFullPath(reportPath));
            Directory.CreateDirectory(output);
            var priorFlag = Environment.GetEnvironmentVariable(
                AnalysisDocumentPilot.FeatureFlag);
            const string pdfDiagnosticFlag =
                "SCRIBBLE_ANALYSIS_PDF_DIAGNOSTIC_DIR";
            var priorPdfDiagnostic = Environment.GetEnvironmentVariable(
                pdfDiagnosticFlag);
            try
            {
                Environment.SetEnvironmentVariable(
                    AnalysisDocumentPilot.FeatureFlag, "1");
                Environment.SetEnvironmentVariable(pdfDiagnosticFlag, output);
                stage = "excel_start";
                excel = Activator.CreateInstance(Type.GetTypeFromProgID(
                    "Excel.Application", true));
                stage = "excel_real_date_column";
                dynamic dateWorkbook = excel.Workbooks.Add();
                try
                {
                    dynamic dateSheet = dateWorkbook.Worksheets[1];
                    dateSheet.Range("A1").Value2 = "Date";
                    dateSheet.Range("B1").Value2 = "Amount";
                    dateSheet.Range("A2").Value2 = 45808d;
                    dateSheet.Range("A2").NumberFormat = "m/d/yy";
                    dateSheet.Range("B2").Value2 = 120d;
                    dateSheet.Range("B2").NumberFormat = "#,##0";
                    dynamic dateRange = dateSheet.Range("A1:B2");
                    object mixed = dateRange.NumberFormat;
                    Check(mixed == DBNull.Value,
                        "Native Excel did not return DBNull for mixed formats.");
                    var formats = WorkbookTypedCapture.ResolveMixedNumberFormats(
                        mixed, 2, 2,
                        column => (object)dateRange.Columns[column + 1].NumberFormat,
                        (row, column) => (object)dateRange.Cells[row + 1,
                            column + 1].NumberFormat);
                    var dateTable = WorkbookTypedCapture.Capture(
                        "native_date", "Dates", (object)dateRange.Value2,
                        (object)dateRange.Formula, formats, null,
                        2, 2, 1, 1);
                    Check(dateTable.Cells.Single(cell =>
                            cell.Reference == "A2").ValueType ==
                            AnalysisContract.DateValue &&
                        dateTable.Cells.Single(cell =>
                            cell.Reference == "B2").ValueType ==
                            AnalysisContract.DecimalValue,
                        "The real Excel date serial was typed as a number.");
                    nativeDateColumnPassed = true;
                }
                finally { dateWorkbook.Close(false); }
                workbook = excel.Workbooks.Add();
                dynamic ledger = workbook.Worksheets[1];
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
                var sourceBefore = SourceFingerprint(ledger);
                stage = "excel_task_analysis_binding";
                var readCall = new ChatToolCall
                {
                    id = "native-analysis-read",
                    type = "function",
                    function = new ChatToolCallFunction
                    {
                        name = WorkbookToolCatalog.ReadCells,
                        arguments = new JavaScriptSerializer().Serialize(new
                        {
                            sheet = "Ledger", range = "B1:J3",
                            analysis_binding = new
                            {
                                period_header = "Period",
                                metrics = new[]
                                {
                                    new { header = "RevenueEUR", currency = "EUR" },
                                    new { header = "CostEUR", currency = "EUR" }
                                }
                            }
                        })
                    }
                };
                var readResult = new WorkbookToolHost((object)excel)
                    .Execute(readCall);
                Check(!readResult.Outcome.Failed &&
                    readResult.Content.Contains("analysis_id"),
                    "The model-facing native read did not bind typed facts.");
                var readInput = new ChatCompletionRequest
                {
                    model = "offline-test",
                    tools = new List<ChatToolDefinition>
                    {
                        WorkbookToolCatalog.DraftDefinition(),
                        CrossAppToolCatalog.CreateDefinitions("excel")
                            .Single(tool => tool.function.name ==
                                CrossAppToolCatalog.SendToPowerPoint)
                    },
                    messages = new List<object>
                    {
                        new ChatCompletionInputMessage
                        {
                            role = "user", content = "Analyze the disposable ledger"
                        }
                    }
                };
                var readStore = new TaskCheckpointStore(Path.Combine(output,
                    "read-checkpoint"));
                var readTask = new TaskContextManager(readInput, "excel",
                    "Analyze the disposable ledger", readStore);
                taskOwner = readTask.State.Id;
                readTask.AfterTool(readCall, readResult);
                var bound = readTask.LoadAnalysis();
                Check(bound != null && bound.Facts.Count == 4 &&
                    bound.Facts.Any(fact => fact.Metric == "RevenueEUR" &&
                        fact.Period == "2026-05" && fact.Value == "85519" &&
                        fact.Locators[0].Cell == "I2") &&
                    readStore.Load(readTask.State.Id).AnalysisArtifactEvidenceId ==
                        readTask.State.AnalysisArtifactEvidenceId,
                    "The task did not durably retain facts from native Excel cells.");
                DocumentChatRequestFactory.ApplyAnalysisPilot(readInput,
                    bound, "excel");
                Check(readInput.tools.Any(tool =>
                        tool.function.name == WorkbookToolCatalog.WriteDraftSheet &&
                        new JavaScriptSerializer().Serialize(
                            tool.function.parameters).Contains("analysis_id")) &&
                    readInput.tools.Any(tool =>
                        tool.function.name ==
                            CrossAppToolCatalog.SendToPowerPoint &&
                        new JavaScriptSerializer().Serialize(
                            tool.function.parameters).Contains("AnalysisId")),
                    "The active request did not switch to typed workbook and deck contracts.");
                var fixture = Fixture(bound);
                stage = "excel_model_plan_boundary";
                var planJson = new JavaScriptSerializer().Serialize(new
                {
                    AnalysisId = fixture.Item1.AnalysisId,
                    WorkbookTitle = fixture.Item2.WorkbookTitle,
                    Slides = fixture.Item2.Slides
                });
                var parsedPlan = AnalysisSlidePlanContract.Parse(
                    fixture.Item1, planJson);
                var injectedFormula = planJson.Replace("\"Formula\":null",
                    "\"Formula\":\"=1\"");
                Check(injectedFormula != planJson,
                    "The native fixture did not contain a table cell for the formula-injection check.");
                var formulaRejected = false;
                try { AnalysisSlidePlanContract.Parse(fixture.Item1,
                    injectedFormula); }
                catch (InvalidOperationException error)
                { formulaRejected = error.Message.Contains(
                    "ANALYSIS_PLAN_FORMULA_FORBIDDEN"); }
                Check(formulaRejected,
                    "A model-authored formula crossed the plan boundary.");
                Check(parsedPlan.WorkbookRows.Count == 3 &&
                    parsedPlan.WorkbookRows[1].Cells[1].Formula
                        .StartsWith("=SUMIF(", StringComparison.Ordinal),
                    "Model plan parsing did not supply host-owned workbook formulas.");
                fixture = Tuple.Create(fixture.Item1, parsedPlan);
                var compiled = AnalysisDocumentCompiler.Compile(
                    fixture.Item1, fixture.Item2);
                var draftCall = new ChatToolCall
                {
                    id = "native-analysis-draft",
                    type = "function",
                    function = new ChatToolCallFunction
                    {
                        name = WorkbookToolCatalog.WriteDraftSheet,
                        arguments = new JavaScriptSerializer().Serialize(new
                        {
                            analysis_id = fixture.Item1.AnalysisId,
                            title = fixture.Item2.WorkbookTitle
                        })
                    }
                };
                Check(ToolContractValidator.Validate(draftCall,
                    readInput.tools.Single(tool => tool.function.name ==
                        WorkbookToolCatalog.WriteDraftSheet)).Count == 0,
                    "The typed draft call does not match the model-facing schema.");
                var draftAuthorization = new OneShotDraftAuthorization(true);
                using (var draftHost = new DocumentDraftHost("excel",
                    (object)excel))
                {
                    draftHost.BindTaskAsync(readTask,
                        CancellationToken.None).GetAwaiter().GetResult();
                    stage = "excel_source_freshness";
                    ledger.Range("I2").Value2 = 85520d;
                    MailboxToolResult staleResult = null;
                    try { staleResult = draftHost.Execute(draftCall,
                        draftAuthorization, true,
                        "Create a verified Excel report"); }
                    finally { ledger.Range("I2").Value2 = 85519d; }
                    Check(staleResult != null && staleResult.Outcome.Failed &&
                        staleResult.Content.Contains("ANALYSIS_SOURCE_CHANGED") &&
                        !draftAuthorization.IsConsumed &&
                        (int)workbook.Worksheets.Count == 1,
                        "A changed source cell did not fail preflight cleanly: " +
                        (staleResult == null ? "no result" : staleResult.Content) +
                        "; consumed=" + draftAuthorization.IsConsumed +
                        "; sheets=" + (int)workbook.Worksheets.Count);
                    stage = "excel_write_and_readback";
                    var draftResult = draftHost.Execute(draftCall,
                        draftAuthorization, true,
                        "Create a verified Excel report");
                    Check(!draftResult.Outcome.Failed &&
                        draftAuthorization.IsCreated &&
                        draftResult.Content.Contains(
                            fixture.Item1.AnalysisId),
                        "The model-facing typed draft did not write a verified report.");
                }
                dynamic draft = workbook.Worksheets["Scribble Draft"];
                Check(compiled.ExpectedFormulaFacts.ContainsKey("B4") &&
                    compiled.ExpectedFormulaFacts.ContainsKey("C4"),
                    "The compiled plan lost its formula readback bindings.");
                Check(Convert.ToDouble(draft.Range("B4").Value2) == 85519d &&
                    Convert.ToDouble(draft.Range("C4").Value2) == 82992d &&
                    Convert.ToDouble(draft.Range("B5").Value2) == 36702d &&
                    Convert.ToDouble(draft.Range("C5").Value2) == 36714d,
                    "Excel native formula results differed from the independent ledger oracle.");
                Check(Convert.ToString(draft.Range("B4").Formula)
                    .StartsWith("=SUMIF(", StringComparison.OrdinalIgnoreCase),
                    "The draft lost its live formula.");
                sourcePreserved = SourceFingerprint(ledger) == sourceBefore;
                Check(sourcePreserved, "The source ledger changed while adding the draft.");
                workbookPassed = true;
                stage = "excel_export";
                draft.ExportAsFixedFormat(0,
                    Path.Combine(output, "analysis-workbook.pdf"));

                stage = "excel_failure_and_retry";
                var formulaCell = fixture.Item2.WorkbookRows[1].Cells[1];
                var correctFactId = formulaCell.ExpectedFactId;
                formulaCell.ExpectedFactId = fixture.Item1.Facts.Single(fact =>
                    fact.Metric == "CostEUR" && fact.Period == "2026-05")
                    .FactId;
                var mismatchRejected = false;
                try
                {
                    AnalysisDocumentPilot.WriteWorkbook((object)excel,
                        fixture.Item1, fixture.Item2);
                }
                catch (InvalidOperationException error)
                {
                    mismatchRejected = error.Message.Contains(
                        "ANALYSIS_PILOT_FORMULA_MISMATCH");
                }
                finally { formulaCell.ExpectedFactId = correctFactId; }
                string failedName = Convert.ToString((object)workbook
                    .Worksheets["Scribble Draft 2"].Name);
                Check(mismatchRejected && failedName == "Scribble Draft 2",
                    "A wrong expected formula was not retained as an isolated failed draft.");
                AnalysisDocumentPilot.WriteWorkbook((object)excel,
                    fixture.Item1, fixture.Item2);
                double corrected = Convert.ToDouble((object)workbook
                    .Worksheets["Scribble Draft 3"].Range("B4").Value2);
                double original = Convert.ToDouble((object)draft
                    .Range("B4").Value2);
                Check(corrected == 85519d && original == 85519d &&
                    SourceFingerprint(ledger) == sourceBefore,
                    "A corrected retry changed the source or the earlier draft.");
                recoveryPassed = true;

                stage = "powerpoint_start";
                powerPoint = Activator.CreateInstance(Type.GetTypeFromProgID(
                    "PowerPoint.Application", true));
                powerPoint.Visible = -1;
                // Exercise the production route before the separate
                // structural writer test, so the chart host starts clean.
                stage = "active_typed_deck_handoff";
                var deckCall = new ChatToolCall
                {
                    id = "native-analysis-deck",
                    type = "function",
                    function = new ChatToolCallFunction
                    {
                        name = CrossAppToolCatalog.SendToPowerPoint,
                        arguments = new JavaScriptSerializer().Serialize(
                            ModelPlanValue(new JavaScriptSerializer()
                                .DeserializeObject(planJson)))
                    }
                };
                Check(ToolContractValidator.Validate(deckCall,
                    readInput.tools.Single(tool => tool.function.name ==
                        CrossAppToolCatalog.SendToPowerPoint)).Count == 0,
                    "The typed deck payload failed its model-facing schema.");
                using (var endpoint = new AnalysisReviewEndpoint())
                using (var client = new OpenAiCompatibleClient())
                using (var deckHost = new DocumentDraftHost("excel",
                    (object)excel))
                {
                    deckHost.BindTaskAsync(readTask,
                        CancellationToken.None).GetAwaiter().GetResult();
                    var settings = new Scribble.Configuration.AppSettings
                    {
                        BaseUrl = endpoint.BaseUrl,
                        ApiKey = "offline-test",
                        Model = "qwen/qwen3.8-27b"
                    };
                    var deckAuthorization =
                        new OneShotDraftAuthorization(true);
                    var beforeDecks =
                        (int)powerPoint.Presentations.Count;
                    ledger.Range("I2").Value2 = 85520d;
                    MailboxToolResult staleDeckResult;
                    try
                    {
                        staleDeckResult = deckHost.ExecuteAsync(deckCall,
                            deckAuthorization, true,
                            "Create a verified four-slide deck",
                            client, settings, CancellationToken.None, null)
                            .GetAwaiter().GetResult();
                    }
                    finally { ledger.Range("I2").Value2 = 85519d; }
                    Check(staleDeckResult.Outcome.Failed &&
                        staleDeckResult.Content.Contains(
                            "ANALYSIS_SOURCE_CHANGED") &&
                        !deckAuthorization.IsConsumed &&
                        (int)powerPoint.Presentations.Count == beforeDecks,
                        "A stale source created a native PowerPoint draft.");
                    var routeResult = deckHost.ExecuteAsync(deckCall,
                        deckAuthorization, true,
                        "Create a verified four-slide deck",
                        client, settings, CancellationToken.None, null)
                        .GetAwaiter().GetResult();
                    Check(!routeResult.Outcome.Failed,
                        "The active typed handoff failed offline review: " +
                        routeResult.Content + "; source_cells_unchanged=" +
                        (SourceFingerprint(ledger) == sourceBefore));
                    for (var p = 1; p <=
                        (int)powerPoint.Presentations.Count; p++)
                    {
                        dynamic candidate = powerPoint.Presentations[p];
                        if ((string)candidate.Tags["ScribbleTask"] ==
                            readTask.State.Id)
                            typedDeck = candidate;
                    }
                    endpoint.Wait();
                    Check(deckAuthorization.IsCreated &&
                        endpoint.ImageCount == 4 &&
                        endpoint.AnalysisId == fixture.Item1.AnalysisId &&
                        routeResult.Content.Contains(
                            fixture.Item1.AnalysisId),
                        "The active typed handoff failed offline review: " +
                        routeResult.Content);
                    Check((object)typedDeck != null,
                        "The handoff created no task-owned draft deck.");
                    Check((int)typedDeck.Slides.Count == 4 &&
                        (string)typedDeck.Tags["ScribbleTask"] ==
                            readTask.State.Id &&
                        readTask.State.HostData.ContainsKey(
                            "analysis_deck_complete"),
                        "The typed handoff lost its task-owned native deck.");
                    var savedPlan = new JavaScriptSerializer()
                        .Deserialize<AnalysisDocumentPlan>(readTask.State
                            .HostData["analysis_deck_plan"]);
                    var repairBudget = AnalysisRepairBudget.Read(readTask.State
                        .HostData["analysis_repair_budget"]);
                    Check(savedPlan.Slides[0].Title ==
                            "Verified June revenue from ledger" &&
                        savedPlan.Slides[0].Layout == "cards" &&
                        (string)typedDeck.Slides[1].Shapes[1]
                            .TextFrame.TextRange.Text ==
                            "Verified June revenue from ledger" &&
                        repairBudget.ModelCalls == 5 &&
                        repairBudget.PatchedTargets.Contains(
                            savedPlan.Slides[0].Id + "/title") &&
                        repairBudget.PatchedTargets.Contains(
                            savedPlan.Slides[0].Id + "/page") &&
                        !readTask.State.HostData.ContainsKey(
                            "analysis_pending_content_patch"),
                        "The typed text/layout repairs lost native readback or recovery receipts.");
                    var typedLayoutImage = Path.Combine(output,
                        "analysis-typed-layout-reflow.png");
                    typedDeck.Slides[1].Export(typedLayoutImage,
                        "PNG", 1920, 1080);
                    Check(File.Exists(typedLayoutImage) &&
                        new FileInfo(typedLayoutImage).Length > 1000,
                        "The reviewed layout reflow did not render.");
                    images.Add(typedLayoutImage);
                    typedDeck.Close();
                    typedDeck = null;
                    typedDeckHandoffPassed = true;
                }
                stage = "powerpoint_write";
                AnalysisDocumentPilot.WritePresentation((object)powerPoint,
                    fixture.Item1, fixture.Item2);
                deck = powerPoint.ActivePresentation;
                Check((int)deck.Slides.Count == 4,
                    "The native deck did not contain exactly four slides.");
                for (var index = 1; index <= 4; index++)
                {
                    dynamic slide = deck.Slides[index];
                    if (PresentationInspection.ContainsNativeChart(
                            (object)slide)) continue;
                    stage = "powerpoint_export_" + index;
                    var path = Path.Combine(output,
                        "analysis-slide-" + index.ToString("00") + ".png");
                    slide.Export(path, "PNG", 1920, 1080);
                    Check(File.Exists(path) && new FileInfo(path).Length > 1000,
                        "A native slide image was not rendered.");
                    images.Add(path);
                }
                stage = "powerpoint_review_metadata";
                var taskInput = new ChatCompletionRequest
                {
                    model = "offline-test",
                    messages = new List<object>
                    {
                        new ChatCompletionInputMessage
                        {
                            role = "user", content =
                                "Review the disposable analysis deck"
                        }
                    }
                };
                var taskStore = new TaskCheckpointStore(Path.Combine(output,
                    "review-checkpoint"));
                var reviewTask = new TaskContextManager(taskInput,
                    "excel", "Review the disposable analysis deck",
                    taskStore);
                reviewTask.PersistAnalysis(fixture.Item1);
                var reviewBefore = Path.Combine(output,
                    "analysis-review-before-export.pptx");
                var reviewAfter = Path.Combine(output,
                    "analysis-review-after-export.pptx");
                deck.SaveCopyAs(reviewBefore);
                AnalysisNativeReviewSession nativeReview;
                try
                {
                    nativeReview = AnalysisDocumentPilot.ReserveNativeReview(
                        reviewTask, (object)deck, fixture.Item1,
                        fixture.Item2, true);
                }
                finally { deck.SaveCopyAs(reviewAfter); }
                var pages = nativeReview.Context.Pages;
                Check(nativeReview.PageImages.Count == pages.Count &&
                    nativeReview.PageImages.Select((image, index) =>
                    {
                        const string prefix = "data:image/png;base64,";
                        if (!image.StartsWith(prefix, StringComparison.Ordinal))
                            return false;
                        var bytes = Convert.FromBase64String(
                            image.Substring(prefix.Length));
                        using (var sha = SHA256.Create())
                            return BitConverter.ToString(sha.ComputeHash(bytes))
                                .Replace("-", "").ToLowerInvariant() ==
                                pages[index].RenderFingerprint;
                    }).All(matches => matches),
                    "Review images did not match their page fingerprints.");
                var chartReviewImage = Path.Combine(output,
                    "analysis-slide-02.png");
                File.WriteAllBytes(chartReviewImage,
                    Convert.FromBase64String(nativeReview.PageImages[1]
                        .Substring("data:image/png;base64,".Length)));
                images.Add(chartReviewImage);
                Check(pages.Count == 4 && pages.Select(page =>
                    page.NativeSlideId).Distinct().Count() == 4 &&
                    pages.Select(page => page.ExpectedPageNumber)
                        .SequenceEqual(new[] { 1, 2, 3, 4 }) &&
                    pages.All(page => page.PageOrdinal == 0 &&
                        page.RenderFingerprint.Length == 64 &&
                        page.NativeStateFingerprint.Length == 64),
                    "The review contract lost native identity, page number or rendered fingerprint.");
                var measurements = nativeReview.Context.Measurements;
                Check(measurements.Count == 0,
                    "The native output has a measured page or text geometry defect: " +
                    string.Join(", ", measurements.Select(item => item.MeasurementId)));
                var reviewContext = nativeReview.Context;
                var resumedReviewTask = new TaskContextManager(taskInput,
                    "excel", "Review the disposable analysis deck",
                    taskStore, taskStore.Load(reviewTask.State.Id));
                Check(AnalysisRepairBudget.Read(nativeReview.Request.BudgetReceipt)
                    .ModelCalls == 1 &&
                    resumedReviewTask.State.HostData["analysis_repair_budget"]
                        == nativeReview.Request.BudgetReceipt &&
                    nativeReview.Request.Content.Contains(
                        reviewContext.ContextId),
                    "The native review request was not checkpointed before inference.");
                var cleanVerdict = new JavaScriptSerializer().Serialize(new
                {
                    contract_version = AnalysisReviewContract.Version,
                    context_id = reviewContext.ContextId,
                    approved = true,
                    findings = new object[0]
                });
                Check(AnalysisDocumentPilot.CompleteNativeReview((object)deck,
                    nativeReview, cleanVerdict).Approved,
                    "A native page-bound typed review could not be parsed.");
                dynamic reviewedTable = null;
                dynamic reviewedSlide = deck.Slides[2];
                for (var shapeIndex = 1;
                    shapeIndex <= (int)reviewedSlide.Shapes.Count; shapeIndex++)
                    if ((int)reviewedSlide.Shapes[shapeIndex].HasTable != 0)
                        reviewedTable = reviewedSlide.Shapes[shapeIndex].Table;
                Check(reviewedTable != null,
                    "The reviewed comparison slide lost its native table.");
                dynamic reviewedCell = reviewedTable.Cell(2, 2).Shape
                    .TextFrame.TextRange;
                var originalCellText = Convert.ToString(reviewedCell.Text);
                var originalCellFontSize = (float)reviewedCell.Font.Size;
                reviewedCell.Text = originalCellText + " edited";
                var staleTableReviewRejected = false;
                try
                {
                    AnalysisDocumentPilot.CompleteNativeReview((object)deck,
                        nativeReview, cleanVerdict);
                }
                catch (InvalidOperationException error)
                {
                    staleTableReviewRejected = error.Message.Contains(
                        "REVIEW_NATIVE_STATE_CHANGED");
                }
                reviewedCell.Text = originalCellText;
                reviewedCell.Font.Size = originalCellFontSize;
                var restoredTableStillNeedsReview = false;
                try
                {
                    AnalysisDocumentPilot.CompleteNativeReview((object)deck,
                        nativeReview, cleanVerdict);
                }
                catch (InvalidOperationException error)
                {
                    restoredTableStillNeedsReview = error.Message.Contains(
                        "REVIEW_NATIVE_STATE_CHANGED");
                }
                Check(staleTableReviewRejected &&
                    restoredTableStillNeedsReview,
                    "A native table-cell edit inherited an earlier review approval.");
                nativeReview = AnalysisDocumentPilot.ReserveNativeReview(
                    reviewTask, (object)deck, fixture.Item1,
                    fixture.Item2, true);
                cleanVerdict = new JavaScriptSerializer().Serialize(new
                {
                    contract_version = AnalysisReviewContract.Version,
                    context_id = nativeReview.Context.ContextId,
                    approved = true,
                    findings = new object[0]
                });
                Check(AnalysisDocumentPilot.CompleteNativeReview((object)deck,
                    nativeReview, cleanVerdict).Approved,
                    "A fresh native review could not bind the restored table.");
                dynamic reviewedChart = null;
                for (var shapeIndex = 1;
                    shapeIndex <= (int)reviewedSlide.Shapes.Count; shapeIndex++)
                    if ((int)reviewedSlide.Shapes[shapeIndex].HasChart != 0)
                        reviewedChart = reviewedSlide.Shapes[shapeIndex].Chart;
                Check(reviewedChart != null,
                    "The reviewed comparison slide lost its native chart.");
                dynamic reviewedSeries = reviewedChart.SeriesCollection(1);
                var originalSeriesName = Convert.ToString(reviewedSeries.Name);
                reviewedSeries.Name = originalSeriesName + " edited";
                var staleChartReviewRejected = false;
                try
                {
                    AnalysisDocumentPilot.CompleteNativeReview((object)deck,
                        nativeReview, cleanVerdict);
                }
                catch (InvalidOperationException error)
                {
                    staleChartReviewRejected = error.Message.Contains(
                        "REVIEW_NATIVE_STATE_CHANGED");
                }
                reviewedSeries.Name = originalSeriesName;
                var restoredChartStillNeedsReview = false;
                try
                {
                    AnalysisDocumentPilot.CompleteNativeReview((object)deck,
                        nativeReview, cleanVerdict);
                }
                catch (InvalidOperationException error)
                {
                    restoredChartStillNeedsReview = error.Message.Contains(
                        "REVIEW_NATIVE_STATE_CHANGED");
                }
                Check(staleChartReviewRejected &&
                    restoredChartStillNeedsReview,
                    "A native chart-series edit inherited an earlier review approval.");
                nativeReview = AnalysisDocumentPilot.ReserveNativeReview(
                    reviewTask, (object)deck, fixture.Item1,
                    fixture.Item2, true);
                cleanVerdict = new JavaScriptSerializer().Serialize(new
                {
                    contract_version = AnalysisReviewContract.Version,
                    context_id = nativeReview.Context.ContextId,
                    approved = true,
                    findings = new object[0]
                });
                Check(AnalysisDocumentPilot.CompleteNativeReview((object)deck,
                    nativeReview, cleanVerdict).Approved,
                    "A fresh native review could not bind the restored chart.");
                typedReviewPassed = true;
                stage = "powerpoint_renderer_repair";
                dynamic firstSlide = deck.Slides[1];
                dynamic folio = null;
                for (var shapeIndex = 1;
                    shapeIndex <= (int)firstSlide.Shapes.Count; shapeIndex++)
                {
                    dynamic shape = firstSlide.Shapes[shapeIndex];
                    if ((int)shape.HasTextFrame != 0 &&
                        (Convert.ToString(shape.TextFrame.TextRange.Text) ??
                            string.Empty).Trim() == "- 1 -")
                        folio = shape;
                }
                Check(folio != null, "The native slide has no editable folio.");
                folio.TextFrame.TextRange.Text = "- 9 -";
                var staleReviewRejected = false;
                try
                {
                    AnalysisDocumentPilot.CompleteNativeReview((object)deck,
                        nativeReview, cleanVerdict);
                }
                catch (InvalidOperationException error)
                {
                    staleReviewRejected = error.Message.Contains(
                        "REVIEW_NATIVE_STATE_CHANGED");
                }
                Check(staleReviewRejected,
                    "A native edit inherited an earlier review approval.");
                var damagedPages = AnalysisDocumentPilot.CapturePresentationPages(
                    (object)deck, fixture.Item1, fixture.Item2);
                var damagedMeasurements =
                    AnalysisDocumentPilot.CaptureNativeMeasurements(
                        (object)deck, damagedPages);
                var budgetBeforeUnresolved = reviewTask.State.HostData[
                    "analysis_repair_budget"];
                var unresolvedReviewRejected = false;
                try
                {
                    AnalysisDocumentPilot.ReserveNativeReview(reviewTask,
                        (object)deck, fixture.Item1, fixture.Item2, true);
                }
                catch (InvalidOperationException error)
                {
                    unresolvedReviewRejected = error.Message.StartsWith(
                        "ANALYSIS_DECK_GEOMETRY_UNRESOLVED:",
                        StringComparison.Ordinal);
                }
                Check(unresolvedReviewRejected &&
                    reviewTask.State.HostData["analysis_repair_budget"] ==
                        budgetBeforeUnresolved,
                    "Unresolved geometry spent a model-review call.");
                var folioDefect = damagedMeasurements.Single(item =>
                    item.Code == "PAGE_NUMBER" &&
                    item.NativeSlideId == damagedPages[0].NativeSlideId);
                var reservedFolio = AnalysisRepairBudget.CrossApp()
                    .ConsumePatch(damagedPages[0].LogicalSlideId,
                        folioDefect.TargetId);
                var folioReservation = new AnalysisPatchReservation
                {
                    LogicalSlideId = damagedPages[0].LogicalSlideId,
                    NativeSlideId = damagedPages[0].NativeSlideId,
                    TargetId = folioDefect.TargetId,
                    MeasurementId = folioDefect.MeasurementId,
                    NativeStateFingerprint =
                        damagedPages[0].NativeStateFingerprint,
                    BudgetReceipt = reservedFolio
                };
                var repairReceipt = AnalysisDocumentPilot.RepairNativeMeasurement(
                    (object)deck, damagedPages, folioDefect,
                    reservedFolio, folioReservation);
                var correctedPages =
                    AnalysisDocumentPilot.CapturePresentationPages(
                        (object)deck, fixture.Item1, fixture.Item2);
                Check(AnalysisDocumentPilot.CaptureNativeMeasurements(
                    (object)deck, correctedPages).Count == 0 &&
                    (Convert.ToString(folio.TextFrame.TextRange.Text) ??
                        string.Empty).Trim() == "- 1 -",
                    "The renderer did not correct and read back a native folio defect.");
                folio.Left = -10f;
                var displacedPages =
                    AnalysisDocumentPilot.CapturePresentationPages(
                        (object)deck, fixture.Item1, fixture.Item2);
                var displaced = AnalysisDocumentPilot.CaptureNativeMeasurements(
                    (object)deck, displacedPages).Single(item =>
                        item.Code == "OUT_OF_BOUNDS" &&
                        item.NativeSlideId == displacedPages[0].NativeSlideId &&
                        item.TargetId == "shape:" + (int)folio.Id);
                repairReceipt = AnalysisDocumentPilot.RepairNativeMeasurement(
                    (object)deck, displacedPages, displaced, repairReceipt);
                var placedPages = AnalysisDocumentPilot.CapturePresentationPages(
                    (object)deck, fixture.Item1, fixture.Item2);
                Check(AnalysisDocumentPilot.CaptureNativeMeasurements(
                    (object)deck, placedPages).Count == 0 &&
                    (float)folio.Left >= 0f &&
                    AnalysisRepairBudget.Read(repairReceipt).PatchedTargets.Count == 2,
                    "The renderer did not correct and read back an out-of-bounds shape.");
                dynamic comparisonSlide = deck.Slides[2];
                dynamic nativeChart = null, table = null;
                for (var shapeIndex = 1;
                    shapeIndex <= (int)comparisonSlide.Shapes.Count; shapeIndex++)
                {
                    dynamic shape = comparisonSlide.Shapes[shapeIndex];
                    if ((int)shape.HasChart != 0) nativeChart = shape;
                    if ((int)shape.HasTable != 0) table = shape;
                }
                Check(nativeChart != null && table != null,
                    "The comparison slide lost its native chart or table.");
                nativeChart.Left = (float)table.Left + (float)table.Width - 100f;
                var collidedPages = AnalysisDocumentPilot.CapturePresentationPages(
                    (object)deck, fixture.Item1, fixture.Item2);
                var collision = AnalysisDocumentPilot.CaptureNativeMeasurements(
                    (object)deck, collidedPages).Single(item =>
                        item.Code == "COLLISION" &&
                        item.NativeSlideId == collidedPages[1].NativeSlideId &&
                        ((item.TargetId == "shape:" + (int)nativeChart.Id &&
                          item.OtherTargetId == "shape:" + (int)table.Id) ||
                         (item.TargetId == "shape:" + (int)table.Id &&
                          item.OtherTargetId == "shape:" + (int)nativeChart.Id)));
                repairReceipt = AnalysisDocumentPilot.RepairNativeMeasurement(
                    (object)deck, collidedPages, collision, repairReceipt);
                var clearedPages = AnalysisDocumentPilot.CapturePresentationPages(
                    (object)deck, fixture.Item1, fixture.Item2);
                Check(AnalysisDocumentPilot.CaptureNativeMeasurements(
                    (object)deck, clearedPages).Count == 0 &&
                    AnalysisRepairBudget.Read(repairReceipt).PatchedTargets.Count == 3,
                    "The renderer did not clear and read back a chart/table collision.");
                var originalChartTop = (float)nativeChart.Top;
                var outsideCanvasRejected = false;
                try
                {
                    dynamic slideTitle = comparisonSlide.Shapes[1];
                    nativeChart.Top = (float)slideTitle.Top;
                    AnalysisDocumentPilot.CaptureNativeMeasurements(
                        (object)deck, clearedPages);
                }
                catch (InvalidOperationException error)
                {
                    outsideCanvasRejected = error.Message.StartsWith(
                        "ANALYSIS_PILOT_GEOMETRY_UNSUPPORTED:",
                        StringComparison.Ordinal);
                }
                finally { nativeChart.Top = originalChartTop; }
                Check(outsideCanvasRejected &&
                    AnalysisDocumentPilot.CaptureNativeMeasurements(
                        (object)deck, clearedPages).Count == 0,
                    "A chart overlapping the title escaped geometry review.");
                dynamic bodyLabel = null;
                for (var shapeIndex = 1;
                    shapeIndex <= (int)firstSlide.Shapes.Count; shapeIndex++)
                {
                    dynamic shape = firstSlide.Shapes[shapeIndex];
                    if ((int)shape.HasTextFrame != 0 &&
                        (Convert.ToString(shape.TextFrame.TextRange.Text) ??
                            string.Empty).Trim() == "JUNE REVENUE")
                        bodyLabel = shape;
                }
                Check(bodyLabel != null,
                    "The scorecard lost its editable revenue label.");
                var originalBodyLeft = (float)bodyLabel.Left;
                var originalBodyTop = (float)bodyLabel.Top;
                var footerOverlapRejected = false;
                try
                {
                    bodyLabel.Left = 36f;
                    bodyLabel.Top = (float)deck.PageSetup.SlideHeight * .91f;
                    AnalysisDocumentPilot.CaptureNativeMeasurements(
                        (object)deck, clearedPages);
                }
                catch (InvalidOperationException error)
                {
                    footerOverlapRejected = error.Message.StartsWith(
                        "ANALYSIS_PILOT_GEOMETRY_UNSUPPORTED:",
                        StringComparison.Ordinal);
                }
                finally
                {
                    bodyLabel.Left = originalBodyLeft;
                    bodyLabel.Top = originalBodyTop;
                }
                Check(footerOverlapRejected &&
                    AnalysisDocumentPilot.CaptureNativeMeasurements(
                        (object)deck, clearedPages).Count == 0,
                    "Body text overlapping the source footer escaped geometry review.");
                rendererRepairPassed = true;
                stage = "powerpoint_save_copy";
                var deckCopy = Path.Combine(output, "analysis-deck.pptx");
                deck.SaveCopyAs(deckCopy);
                stage = "powerpoint_package_oracle";
                Check(File.Exists(deckCopy) && new FileInfo(deckCopy).Length > 1000,
                    "PowerPoint did not create a native draft package.");
                using (var package = ZipFile.OpenRead(deckCopy))
                {
                    var headline = PackageText(package, "ppt/slides/slide1.xml");
                    var comparison = PackageText(package, "ppt/slides/slide2.xml");
                    var chartParts = package.Entries.Where(entry =>
                        entry.FullName.StartsWith("ppt/charts/chart",
                            StringComparison.OrdinalIgnoreCase) &&
                        entry.FullName.EndsWith(".xml",
                            StringComparison.OrdinalIgnoreCase))
                        .Select(PackageText).ToArray();
                    Check(chartParts.Any(chart => chart.Contains("85519") &&
                        chart.Contains("82992")),
                        "The native chart cache did not match the verified periods.");
                    Check(HasChartRelationship(package,
                        "ppt/slides/_rels/slide2.xml.rels"),
                        "The comparison slide lost its native chart relationship.");
                    Check(headline.Contains("82,992") &&
                        headline.Contains("36,714") &&
                        comparison.Contains("85,519") &&
                        comparison.Contains("82,992"),
                        "The native slide content differed from the independent fact oracle.");
                    var notes = package.Entries.Where(entry =>
                        entry.FullName.StartsWith("ppt/notesSlides/notesSlide",
                            StringComparison.OrdinalIgnoreCase) &&
                        entry.FullName.EndsWith(".xml",
                            StringComparison.OrdinalIgnoreCase))
                        .Select(PackageText).ToArray();
                    Check(headline.Contains("Source: Ledger!") &&
                        notes.Any(note => note.Contains(
                            fixture.Item1.Snapshots[0].SourceInstanceId)),
                        "The visible source label or exact speaker-note citation was lost.");
                }
                stage = "powerpoint_content_recovery_injection";
                var recoveryStore = new TaskCheckpointStore(Path.Combine(
                    output, "content-recovery-checkpoint"));
                var recoveryInput = new ChatCompletionRequest
                {
                    model = "offline-test",
                    messages = new List<object>
                    {
                        new ChatCompletionInputMessage
                        { role = "user", content = "Recover the native title" }
                    }
                };
                var recoveryTask = new TaskContextManager(recoveryInput,
                    "excel", "Recover the native title", recoveryStore);
                recoveryTask.PersistAnalysis(fixture.Item1);
                var recoveryPages = AnalysisDocumentPilot
                    .CapturePresentationPages((object)deck, fixture.Item1,
                        fixture.Item2);
                var recoveryPage = recoveryPages[0];
                var originalTitle = fixture.Item2.Slides[0].Title;
                var correctedTitle =
                    "Verified June revenue from ledger";
                var desiredPlan = new JavaScriptSerializer
                    { MaxJsonLength = 16000000 }
                    .Deserialize<AnalysisDocumentPlan>(new JavaScriptSerializer
                    { MaxJsonLength = 16000000 }.Serialize(fixture.Item2));
                desiredPlan.Slides[0].Title = correctedTitle;
                var nativePatch = new AnalysisDocumentPatch
                {
                    ContextId = "native-failure-injection",
                    LogicalSlideId = recoveryPage.LogicalSlideId,
                    TargetId = "title", SegmentIndex = 0,
                    ExpectedText = originalTitle,
                    ReplacementText = correctedTitle
                };
                var expectedBudget = AnalysisRepairBudget.CrossApp()
                    .ConsumePatch(nativePatch.LogicalSlideId,
                        nativePatch.TargetId);
                var contentReservation = recoveryTask
                    .ReserveAnalysisContentPatch(recoveryPage,
                        nativePatch, desiredPlan, originalTitle,
                        correctedTitle, "native-input", expectedBudget,
                        true);
                var staleReservation = new AnalysisContentPatchReservation
                {
                    LogicalSlideId = contentReservation.LogicalSlideId,
                    NativeSlideId = contentReservation.NativeSlideId,
                    NativeStateFingerprint = "stale-native-state",
                    NativeBeforeText = originalTitle,
                    NativeAfterText = correctedTitle
                };
                var stalePatchRejected = false;
                try
                {
                    AnalysisDocumentPilot.ApplyNativeContentPatch(
                        (object)deck, recoveryPage, staleReservation);
                }
                catch (InvalidOperationException error)
                {
                    stalePatchRejected = error.Message.Contains(
                        "REPAIR_RESERVATION_CHANGED");
                }
                Check(stalePatchRejected &&
                    AnalysisDocumentPilot.ReadNativePatchText(
                        (object)deck, recoveryPage, originalTitle) ==
                        originalTitle,
                    "A stale content reservation changed the native slide.");
                var injectedRollbackPassed = false;
                try
                {
                    AnalysisDocumentPilot.ApplyNativeContentPatch(
                        (object)deck, recoveryPage, contentReservation,
                        () => { throw new InvalidOperationException(
                            "INJECT_AFTER_NATIVE_WRITE"); });
                }
                catch (InvalidOperationException error)
                {
                    injectedRollbackPassed = error.Message ==
                        "INJECT_AFTER_NATIVE_WRITE";
                }
                Check(injectedRollbackPassed &&
                    recoveryTask.State.HostData.ContainsKey(
                        "analysis_pending_content_patch") &&
                    AnalysisDocumentPilot.ReadNativePatchText(
                        (object)deck, recoveryPage, originalTitle) ==
                        originalTitle,
                    "A partial native text write was not fully rolled back " +
                    "with its pending receipt intact.");
                partialContentRollbackPassed = true;
                AnalysisDocumentPilot.ApplyNativeContentPatch(
                    (object)deck, recoveryPage, contentReservation);
                var changedPages = AnalysisDocumentPilot
                    .CapturePresentationPages((object)deck, fixture.Item1,
                        desiredPlan);
                var changedPage = changedPages[0];
                var resumedRecovery = new TaskContextManager(recoveryInput,
                    "excel", recoveryTask.State.Objective, recoveryStore,
                    recoveryStore.Load(recoveryTask.State.Id));
                var blockedReview = false;
                try
                {
                    resumedRecovery.ReserveAnalysisReview(fixture.Item1,
                        fixture.Item2, AnalysisReviewContract.Context(
                            fixture.Item1, fixture.Item2, recoveryPages),
                        true);
                }
                catch (InvalidOperationException error)
                {
                    blockedReview = error.Message.Contains(
                        "REPAIR_PENDING_RECONCILIATION");
                }
                Check(blockedReview && changedPage.NativeStateFingerprint !=
                    recoveryPage.NativeStateFingerprint,
                    "An unreceipted native text write resumed as approved.");
                dynamic otherText = null;
                dynamic changedSlide = deck.Slides[
                    changedPage.ExpectedPageNumber];
                for (var shapeIndex = 1;
                    shapeIndex <= (int)changedSlide.Shapes.Count;
                    shapeIndex++)
                {
                    dynamic shape = changedSlide.Shapes[shapeIndex];
                    if ((int)shape.HasTextFrame == 0 ||
                        (int)shape.HasChart != 0 ||
                        (int)shape.HasTable != 0) continue;
                    var value = Convert.ToString(
                        shape.TextFrame.TextRange.Text) ?? string.Empty;
                    if (value.Length == 0 || value == correctedTitle)
                        continue;
                    otherText = shape.TextFrame.TextRange;
                    break;
                }
                Check(otherText != null,
                    "The native text receipt fixture has no independent shape.");
                var otherTextBefore = (string)otherText.Text;
                var concurrentTextEditRejected = false;
                try
                {
                    otherText.Text = otherTextBefore + " (user edit)";
                    AnalysisDocumentPilot.ReadNativePatchText((object)deck,
                        changedPage, correctedTitle);
                }
                catch (InvalidOperationException error)
                {
                    concurrentTextEditRejected = error.Message.Contains(
                        "REPAIR_NATIVE_PAGE_CHANGED");
                }
                finally { otherText.Text = otherTextBefore; }
                Check(concurrentTextEditRejected &&
                    AnalysisDocumentPilot.ReadNativePatchText((object)deck,
                        changedPage, correctedTitle) == correctedTitle,
                    "A concurrent edit outside the target text cleared its receipt.");
                resumedRecovery.ReconcileAnalysisContentPatch(
                    contentReservation, changedPage,
                    AnalysisDocumentPilot.ReadNativePatchText(
                        (object)deck, changedPage, correctedTitle));
                Check(!resumedRecovery.State.HostData.ContainsKey(
                        "analysis_pending_content_patch") &&
                    resumedRecovery.State.HostData.ContainsKey(
                        "analysis_deck_plan"),
                    "Native text readback did not reconcile the pending patch.");
                contentRecoveryPassed = true;
                stage = "powerpoint_card_content_recovery";
                var cardPage = changedPages.Single(page =>
                    page.LogicalSlideId == "provenance");
                var cardBefore = AnalysisDocumentPilot.CompiledNativeText(
                    fixture.Item1, desiredPlan, "provenance", "cards[0]");
                Check(AnalysisDocumentPilot.ReadNativePatchText(
                        (object)deck, cardPage, cardBefore) == cardBefore,
                    "The card literal did not identify one native text shape.");
                var cardPlan = new JavaScriptSerializer
                    { MaxJsonLength = 16000000 }
                    .Deserialize<AnalysisDocumentPlan>(new JavaScriptSerializer
                    { MaxJsonLength = 16000000 }.Serialize(desiredPlan));
                cardPlan.Slides.Single(slide => slide.Id == "provenance")
                    .Cards[0].Points[0].Text =
                    "Verified workbook range and checked formulas";
                var cardAfter = AnalysisDocumentPilot.CompiledNativeText(
                    fixture.Item1, cardPlan, "provenance", "cards[0]");
                var cardTask = new TaskContextManager(new ChatCompletionRequest
                {
                    model = "offline-test",
                    messages = new List<object> { new ChatCompletionInputMessage
                        { role = "user", content = "Correct the card text" } }
                }, "excel", "Correct the card text",
                    new TaskCheckpointStore(Path.Combine(output,
                        "card-recovery-checkpoint")));
                cardTask.PersistAnalysis(fixture.Item1);
                var cardPatch = new AnalysisDocumentPatch
                {
                    ContextId = "native-card-recovery",
                    LogicalSlideId = "provenance", TargetId = "cards[0]",
                    SegmentIndex = 0, ExpectedText = cardBefore,
                    ReplacementText = cardAfter
                };
                var cardReservation = cardTask.ReserveAnalysisContentPatch(
                    cardPage, cardPatch, cardPlan, cardBefore, cardAfter,
                    "native-card-input", AnalysisRepairBudget.CrossApp()
                        .ConsumePatch("provenance", "cards[0]"), true);
                AnalysisDocumentPilot.ApplyNativeContentPatch(
                    (object)deck, cardPage, cardReservation);
                var cardReadbackPage = AnalysisDocumentPilot
                    .CapturePresentationPages((object)deck, fixture.Item1,
                        cardPlan).Single(page => page.LogicalSlideId ==
                            "provenance");
                cardTask.ReconcileAnalysisContentPatch(cardReservation,
                    cardReadbackPage,
                    AnalysisDocumentPilot.ReadNativePatchText(
                        (object)deck, cardReadbackPage, cardAfter));
                Check(!cardTask.State.HostData.ContainsKey(
                        "analysis_pending_content_patch") &&
                    desiredPlan.Slides.Single(slide => slide.Id ==
                        "provenance").Cards[0].Points[0].Text == cardBefore,
                    "Native card repair did not preserve the original plan or reconcile the receipt.");
                cardContentRecoveryPassed = true;
                stage = "powerpoint_layout_recovery";
                var layoutPages = AnalysisDocumentPilot
                    .CapturePresentationPages((object)deck, fixture.Item1,
                        cardPlan);
                var layoutPage = layoutPages.Single(page =>
                    page.LogicalSlideId == "headline");
                var layoutPlan = new JavaScriptSerializer
                    { MaxJsonLength = 16000000 }
                    .Deserialize<AnalysisDocumentPlan>(new JavaScriptSerializer
                    { MaxJsonLength = 16000000 }.Serialize(cardPlan));
                layoutPlan.Slides.Single(slide => slide.Id == "headline")
                    .Layout = "cards";
                var layoutInput = new ChatCompletionRequest
                    {
                        model = "offline-test",
                        messages = new List<object>
                        {
                            new ChatCompletionInputMessage
                            { role = "user", content = "Reflow the KPI slide" }
                        }
                    };
                var layoutStore = new TaskCheckpointStore(Path.Combine(
                    output, "layout-recovery-checkpoint"));
                var layoutTask = new TaskContextManager(layoutInput,
                    "excel", "Reflow the KPI slide", layoutStore);
                layoutTask.PersistAnalysis(fixture.Item1);
                var layoutPatch = new AnalysisDocumentPatch
                {
                    ContextId = "native-layout-recovery",
                    LogicalSlideId = "headline", TargetId = "page",
                    SegmentIndex = 0, ExpectedText = "scorecard",
                    ReplacementText = "cards"
                };
                var layoutReservation = layoutTask
                    .ReserveAnalysisContentPatch(layoutPage, layoutPatch,
                        layoutPlan, "scorecard", "cards",
                        "native-layout-input", AnalysisRepairBudget.CrossApp()
                            .ConsumePatch("headline", "page"), true);
                var staleLayoutReservation = new AnalysisContentPatchReservation
                {
                    TargetId = "page", LogicalSlideId = "headline",
                    NativeSlideId = layoutPage.NativeSlideId,
                    NativeStateFingerprint = "stale-native-state",
                    NativeBeforeText = "scorecard",
                    NativeAfterText = "cards"
                };
                var staleLayoutRejected = false;
                try
                {
                    AnalysisDocumentPilot.ApplyNativeLayoutPatch(
                        (object)deck, layoutPage, staleLayoutReservation,
                        fixture.Item1, layoutPlan);
                }
                catch (InvalidOperationException error)
                {
                    staleLayoutRejected = error.Message.Contains(
                        "REPAIR_RESERVATION_CHANGED");
                }
                Check(staleLayoutRejected,
                    "A stale layout reservation was allowed to mutate the slide.");
                AnalysisDocumentPilot.ApplyNativeLayoutPatch((object)deck,
                    layoutPage, layoutReservation, fixture.Item1,
                    layoutPlan);
                var reflowedPage = AnalysisDocumentPilot
                    .CapturePresentationPages((object)deck, fixture.Item1,
                        layoutPlan).Single(page => page.LogicalSlideId ==
                            "headline");
                Check(reflowedPage.NativeSlideId == layoutPage.NativeSlideId &&
                    reflowedPage.NativeStateFingerprint !=
                        layoutPage.NativeStateFingerprint &&
                    AnalysisDocumentPilot.ReadNativeLayoutPatch(
                        (object)deck, reflowedPage, "cards") == "cards",
                    "The native layout reflow lost its slide identity or readback.");
                dynamic reflowedTitle = deck.Slides[
                    reflowedPage.ExpectedPageNumber].Shapes[1]
                    .TextFrame.TextRange;
                var titleBeforeConcurrentEdit = (string)reflowedTitle.Text;
                var concurrentEditRejected = false;
                try
                {
                    reflowedTitle.Text = titleBeforeConcurrentEdit +
                        " (user edit)";
                    AnalysisDocumentPilot.ReadNativeLayoutPatch(
                        (object)deck, reflowedPage, "cards");
                }
                catch (InvalidOperationException error)
                {
                    concurrentEditRejected = error.Message.Contains(
                        "REPAIR_NATIVE_PAGE_CHANGED");
                }
                finally
                {
                    reflowedTitle.Text = titleBeforeConcurrentEdit;
                }
                Check(concurrentEditRejected &&
                    AnalysisDocumentPilot.ReadNativeLayoutPatch(
                        (object)deck, reflowedPage, "cards") == "cards",
                    "A concurrent native edit was accepted by layout readback.");
                var resumedLayoutTask = new TaskContextManager(layoutInput,
                    "excel", layoutTask.State.Objective, layoutStore,
                    layoutStore.Load(layoutTask.State.Id));
                var pendingLayoutRejected = false;
                try
                {
                    resumedLayoutTask.ReserveAnalysisReview(fixture.Item1,
                        layoutPlan, AnalysisReviewContract.Context(
                            fixture.Item1, layoutPlan,
                            AnalysisDocumentPilot.CapturePresentationPages(
                                (object)deck, fixture.Item1, layoutPlan)),
                        true);
                }
                catch (InvalidOperationException error)
                {
                    pendingLayoutRejected = error.Message.Contains(
                        "REPAIR_PENDING_RECONCILIATION");
                }
                Check(pendingLayoutRejected,
                    "An unreceipted layout reflow resumed as approved.");
                var wrongLayoutRejected = false;
                try
                {
                    resumedLayoutTask.ReconcileAnalysisContentPatch(
                        layoutReservation, reflowedPage, "scorecard");
                }
                catch (InvalidOperationException error)
                {
                    wrongLayoutRejected = error.Message.Contains(
                        "REPAIR_PENDING_RECONCILIATION");
                }
                Check(wrongLayoutRejected,
                    "A wrong layout readback cleared the pending receipt.");
                resumedLayoutTask.ReconcileAnalysisContentPatch(
                    layoutReservation, reflowedPage,
                    AnalysisDocumentPilot.ReadNativeLayoutPatch(
                        (object)deck, reflowedPage, "cards"));
                Check(!resumedLayoutTask.State.HostData.ContainsKey(
                        "analysis_pending_content_patch") &&
                    resumedLayoutTask.State.HostData[
                        "analysis_deck_plan"].Contains("\"Layout\":\"cards\"") &&
                    cardPlan.Slides[0].Layout == "scorecard",
                    "The native layout receipt did not preserve and commit the plan.");
                layoutRecoveryPassed = true;
                slidesPassed = true;
                stage = "powerpoint_partial_layout_recovery";
                AnalysisDocumentPilot.WritePresentation((object)powerPoint,
                    fixture.Item1, layoutPlan);
                faultDeck = powerPoint.ActivePresentation;
                var faultPage = AnalysisDocumentPilot
                    .CapturePresentationPages((object)faultDeck,
                        fixture.Item1, layoutPlan)
                    .Single(page => page.LogicalSlideId == "headline");
                var faultStore = new TaskCheckpointStore(Path.Combine(
                    output, "partial-layout-checkpoint"));
                var faultTask = new TaskContextManager(layoutInput,
                    "excel", "Inject a partial layout write", faultStore);
                faultTask.PersistAnalysis(fixture.Item1);
                var faultPatch = new AnalysisDocumentPatch
                {
                    ContextId = "native-layout-partial-write",
                    LogicalSlideId = "headline", TargetId = "page",
                    SegmentIndex = 0, ExpectedText = "cards",
                    ReplacementText = "scorecard"
                };
                var faultReservation = faultTask
                    .ReserveAnalysisContentPatch(faultPage, faultPatch,
                        cardPlan, "cards", "scorecard",
                        "native-layout-partial-input",
                        AnalysisRepairBudget.CrossApp().ConsumePatch(
                            "headline", "page"), true);
                var presentationNames = new HashSet<string>(
                    StringComparer.Ordinal);
                for (var presentationIndex = 1;
                    presentationIndex <= (int)powerPoint.Presentations.Count;
                    presentationIndex++)
                    presentationNames.Add(Convert.ToString(
                        powerPoint.Presentations[presentationIndex].Name));
                var partialLayoutRejected = false;
                try
                {
                    AnalysisDocumentPilot.ApplyNativeLayoutPatch(
                        (object)faultDeck, faultPage, faultReservation,
                        fixture.Item1, cardPlan,
                        () => { throw new InvalidOperationException(
                            "INJECT_AFTER_LAYOUT_DELETE"); });
                }
                catch (InvalidOperationException error)
                {
                    partialLayoutRejected = error.Message.Contains(
                        "SLIDE_REPAIR_RECOVERY_REQUIRED");
                }
                var recoveryCandidates = new List<object>();
                for (var presentationIndex = 1;
                    presentationIndex <= (int)powerPoint.Presentations.Count;
                    presentationIndex++)
                {
                    dynamic candidate =
                        powerPoint.Presentations[presentationIndex];
                    if (!presentationNames.Contains(
                        Convert.ToString(candidate.Name)))
                        recoveryCandidates.Add((object)candidate);
                }
                Check(recoveryCandidates.Count == 1,
                    "The partial layout write did not retain exactly one " +
                    "new native recovery presentation.");
                recoveryDeck = recoveryCandidates[0];
                dynamic backupSlide = recoveryDeck.Slides[2];
                var backupHasTitle = false;
                for (var shapeIndex = 1;
                    shapeIndex <= (int)backupSlide.Shapes.Count;
                    shapeIndex++)
                {
                    dynamic shape = backupSlide.Shapes[shapeIndex];
                    if ((int)shape.HasTextFrame != 0 &&
                        Convert.ToString(shape.TextFrame.TextRange.Text) ==
                            correctedTitle)
                        backupHasTitle = true;
                }
                var persistedFault = faultStore.Load(faultTask.State.Id);
                var faultSlideBlank = (int)faultDeck.Slides[
                    faultPage.ExpectedPageNumber].Shapes.Count == 0;
                var twoRecoverySlides = (int)recoveryDeck.Slides.Count == 2;
                var backupOwnerMatches =
                    (string)backupSlide.Tags["ScribbleTask"] ==
                    (string)faultDeck.Tags["ScribbleTask"];
                var pendingFaultReceipt = persistedFault.HostData.ContainsKey(
                    "analysis_pending_content_patch");
                Check(partialLayoutRejected && faultSlideBlank &&
                    twoRecoverySlides && backupOwnerMatches &&
                    backupHasTitle && pendingFaultReceipt,
                    "A partial native layout write lost its original " +
                    "slide or the pending repair receipt: rejected=" +
                    partialLayoutRejected + ", blank=" + faultSlideBlank +
                    ", two_recovery_slides=" + twoRecoverySlides +
                    ", owner=" + backupOwnerMatches + ", title=" +
                    backupHasTitle + ", receipt=" + pendingFaultReceipt);
                partialLayoutRecoveryPassed = true;
            }
            catch (Exception error)
            {
                powerpointExited = PowerPointExited(error);
                failure = (powerpointExited ? "POWERPOINT_EXITED at " +
                    stage + ": " : stage + ": ") + error;
            }
            finally
            {
                Environment.SetEnvironmentVariable(
                    AnalysisDocumentPilot.FeatureFlag, priorFlag);
                Environment.SetEnvironmentVariable(pdfDiagnosticFlag,
                    priorPdfDiagnostic);
                if ((object)deck != null) try { deck.Close(); } catch { }
                if ((object)faultDeck != null)
                    try { faultDeck.Close(); } catch { }
                if ((object)recoveryDeck != null)
                    try { recoveryDeck.Close(); } catch { }
                if ((object)typedDeck != null)
                    try { typedDeck.Close(); } catch { }
                if ((object)powerPoint != null && taskOwner != null)
                    try
                    {
                        for (var index = (int)powerPoint.Presentations.Count;
                            index >= 1; index--)
                        {
                            dynamic candidate =
                                powerPoint.Presentations[index];
                            if ((string)candidate.Tags["ScribbleTask"] ==
                                taskOwner)
                                candidate.Close();
                        }
                    }
                    catch { }
                if ((object)workbook != null) try { workbook.Close(false); } catch { }
                if ((object)excel != null) try { excel.Quit(); } catch { }
                // PowerPoint can be a shared singleton, so close only our deck.
            }
            var report = new
            {
                execution_kind = "native_disposable_analysis_pilot",
                workbook_passed = workbookPassed,
                four_slides_passed = slidesPassed,
                source_preserved = sourcePreserved,
                isolated_retry_passed = recoveryPassed,
                typed_review_contract_passed = typedReviewPassed,
                renderer_repair_passed = rendererRepairPassed,
                content_recovery_passed = contentRecoveryPassed,
                partial_content_rollback_passed =
                    partialContentRollbackPassed,
                card_content_recovery_passed = cardContentRecoveryPassed,
                layout_recovery_passed = layoutRecoveryPassed,
                partial_layout_recovery_passed =
                    partialLayoutRecoveryPassed,
                typed_deck_handoff_passed = typedDeckHandoffPassed,
                native_date_column_passed = nativeDateColumnPassed,
                powerpoint_exited = powerpointExited,
                visual_review_unavailable = failure.Contains(
                    "ANALYSIS_VISUAL_REVIEW_UNAVAILABLE"),
                rendered_images = images,
                full_acceptance_passed = false,
                note = "Hand-authored structural pilot and offline fake reviewer only; no model, visual attestation, or recovery qualification.",
                failure
            };
            var json = new JavaScriptSerializer().Serialize(report);
            File.WriteAllText(reportPath, json);
            Console.WriteLine(json);
            return workbookPassed && slidesPassed && sourcePreserved &&
                recoveryPassed && typedReviewPassed && rendererRepairPassed &&
                typedDeckHandoffPassed && contentRecoveryPassed &&
                partialContentRollbackPassed &&
                cardContentRecoveryPassed && layoutRecoveryPassed
                && partialLayoutRecoveryPassed
                ? 0 : 1;
        }

        private static string SourceFingerprint(dynamic sheet)
        {
            return string.Join("|", new[] { "B2", "I2", "J2", "B3", "I3", "J3" }
                .Select(cell => Convert.ToString(sheet.Range(cell).Value2,
                    CultureInfo.InvariantCulture)));
        }

        private static bool PowerPointExited(Exception error)
        {
            for (var current = error; current != null;
                current = current.InnerException)
            {
                if (current is COMException &&
                    (unchecked((uint)current.HResult) == 0x800706BA ||
                     unchecked((uint)current.HResult) == 0x800706BE ||
                     unchecked((uint)current.HResult) == 0x80010108))
                    return true;
                if (current.Message.Contains(
                        "\"error_code\":\"POWERPOINT_EXITED\"") ||
                    current.Message.Contains("0x800706BA"))
                    return true;
            }
            return false;
        }

        private static object ModelPlanValue(object value)
        {
            var map = value as IDictionary<string, object>;
            if (map != null)
            {
                var result = new Dictionary<string, object>(
                    StringComparer.Ordinal);
                foreach (var item in map)
                {
                    if (item.Key == "Formula" ||
                        item.Key == "ExpectedFactId" || item.Value == null)
                        continue;
                    var child = ModelPlanValue(item.Value);
                    var array = child as object[];
                    if (array != null && array.Length == 0) continue;
                    result.Add(item.Key, child);
                }
                return result;
            }
            var list = value as IList;
            return list == null ? value : list.Cast<object>()
                .Select(ModelPlanValue).ToArray();
        }

        // The native route sends its rendered pages to a loopback endpoint.
        // It accepts only an exact image/hash pairing and replies with one
        // typed approval for the supplied context. No paid model is involved.
        private sealed class AnalysisReviewEndpoint : IDisposable
        {
            private readonly TcpListener _listener = new TcpListener(
                IPAddress.Loopback, 0);
            private readonly Task _worker;

            public AnalysisReviewEndpoint()
            {
                _listener.Start();
                BaseUrl = "http://127.0.0.1:" +
                    ((IPEndPoint)_listener.LocalEndpoint).Port + "/v1";
                _worker = Task.Run((Action)Handle);
            }

            public string BaseUrl { get; }
            public string AnalysisId { get; private set; }
            public int ImageCount { get; private set; }

            public void Wait()
            {
                if (!_worker.Wait(TimeSpan.FromSeconds(30)))
                    throw new InvalidOperationException(
                        "The offline reviewer received no request.");
                if (_worker.IsFaulted)
                    throw _worker.Exception.GetBaseException();
            }

            public void Dispose()
            {
                _listener.Stop();
                try { _worker.Wait(TimeSpan.FromSeconds(1)); }
                catch { }
            }

            private void Handle()
            {
                for (var round = 0; round < 5; round++)
                    HandleOne(round);
            }

            private void HandleOne(int round)
            {
                using (var client = _listener.AcceptTcpClient())
                using (var stream = client.GetStream())
                using (var reader = new StreamReader(stream,
                    Encoding.ASCII, false, 4096, true))
                {
                    var requestLine = reader.ReadLine() ?? string.Empty;
                    Check(requestLine.Contains("/chat/completions"),
                        "The native reviewer used an unexpected endpoint.");
                    var contentLength = 0;
                    string line;
                    while (!string.IsNullOrEmpty(line = reader.ReadLine()))
                    {
                        if (line.StartsWith("Content-Length:",
                            StringComparison.OrdinalIgnoreCase))
                            int.TryParse(line.Substring(15).Trim(),
                                out contentLength);
                    }
                    Check(contentLength > 0 && contentLength < 16000000,
                        "The native reviewer request size is invalid.");
                    var buffer = new char[contentLength];
                    var offset = 0;
                    while (offset < buffer.Length)
                    {
                        var read = reader.Read(buffer, offset,
                            buffer.Length - offset);
                        if (read <= 0) break;
                        offset += read;
                    }
                    Check(offset == buffer.Length,
                        "The native reviewer request was truncated.");
                    var json = new JavaScriptSerializer
                        { MaxJsonLength = 16000000 };
                    var request = (IDictionary<string, object>)
                        json.DeserializeObject(new string(buffer));
                    var messages = (IList)request["messages"];
                    var user = (IDictionary<string, object>)
                        messages[messages.Count - 1];
                    string decision;
                    if (round == 1 || round == 3)
                    {
                        var proposal = (IDictionary<string, object>)
                            json.DeserializeObject((string)user["content"]);
                        if (round == 1)
                        {
                            var segments = (IList)proposal[
                                "editable_literals"];
                            var segment = (IDictionary<string, object>)
                                segments[0];
                            Check((string)proposal["target_id"] == "title" &&
                                (string)segment["text"] ==
                                    "June revenue at a glance",
                                "The repair prompt lost its exact title literal.");
                            decision = json.Serialize(new
                            {
                                context_id = (string)proposal["context_id"],
                                logical_slide_id =
                                    (string)proposal["logical_slide_id"],
                                target_id = "title", segment_index = 0,
                                expected_text = (string)segment["text"],
                                replacement_text =
                                    "Verified June revenue from ledger"
                            });
                        }
                        else
                        {
                            Check((string)proposal["target_id"] == "page" &&
                                (string)proposal["current_layout"] ==
                                    "scorecard" &&
                                ((IList)proposal["allowed_layouts"])
                                    .Cast<string>().Contains("cards"),
                                "The layout prompt lost its bounded recipe choices.");
                            decision = json.Serialize(new
                            {
                                context_id = (string)proposal["context_id"],
                                logical_slide_id =
                                    (string)proposal["logical_slide_id"],
                                target_id = "page", segment_index = 0,
                                expected_text = "scorecard",
                                replacement_text = "cards"
                            });
                        }
                    }
                    else
                    {
                        var parts = (IList)user["content"];
                        var textPart = (IDictionary<string, object>)parts[0];
                        var content = (IDictionary<string, object>)
                            json.DeserializeObject((string)textPart["text"]);
                        AnalysisId = (string)content["analysis_id"];
                        var pages = (IList)content["pages"];
                        ImageCount = parts.Count - 1;
                        Check(ImageCount == pages.Count && ImageCount == 4,
                            "The typed reviewer did not receive four pages.");
                        for (var index = 0; index < ImageCount; index++)
                        {
                            var part = (IDictionary<string, object>)
                                parts[index + 1];
                            var image = (IDictionary<string, object>)
                                part["image_url"];
                            var url = (string)image["url"];
                            const string prefix = "data:image/png;base64,";
                            Check(url.StartsWith(prefix,
                                StringComparison.Ordinal),
                                "The reviewer image is not an inline PNG.");
                            var bytes = Convert.FromBase64String(
                                url.Substring(prefix.Length));
                            var page = (IDictionary<string, object>)pages[index];
                            using (var sha = SHA256.Create())
                                Check(BitConverter.ToString(
                                    sha.ComputeHash(bytes)).Replace("-", "")
                                    .ToLowerInvariant() ==
                                    (string)page["RenderFingerprint"],
                                    "The reviewer image changed after capture.");
                        }
                        var firstPage = (IDictionary<string, object>)pages[0];
                        if (round == 0)
                            decision = json.Serialize(new
                            {
                                contract_version =
                                    AnalysisReviewContract.Version,
                                context_id = (string)content["context_id"],
                                approved = false,
                                findings = new[] { new
                                {
                                    code = "UNSUPPORTED_CLAIM",
                                    owner = "content",
                                    logical_slide_id = (string)
                                        firstPage["LogicalSlideId"],
                                    native_slide_id = Convert.ToInt32(
                                        firstPage["NativeSlideId"]),
                                    target_id = "title", fact_id = "",
                                    measurement_id = "",
                                    severity = "blocker",
                                    action = "revise_text",
                                    evidence = "Make ledger scope explicit."
                                } }
                            });
                        else if (round == 2)
                        {
                            var slides = (IList)content["logical_slides"];
                            var firstSlide = (IDictionary<string, object>)
                                slides[0];
                            Check((string)firstSlide["title"] ==
                                "Verified June revenue from ledger" &&
                                (string)firstSlide["layout"] == "scorecard",
                                "The corrected title did not reach layout review.");
                            decision = json.Serialize(new
                            {
                                contract_version =
                                    AnalysisReviewContract.Version,
                                context_id = (string)content["context_id"],
                                approved = false,
                                findings = new[] { new
                                {
                                    code = "VISUAL_HIERARCHY",
                                    owner = "content",
                                    logical_slide_id = (string)
                                        firstPage["LogicalSlideId"],
                                    native_slide_id = Convert.ToInt32(
                                        firstPage["NativeSlideId"]),
                                    target_id = "page", fact_id = "",
                                    measurement_id = "",
                                    severity = "blocker",
                                    action = "revise_layout",
                                    evidence = "The two KPIs need stronger hierarchy."
                                } }
                            });
                        }
                        else
                        {
                            var slides = (IList)content["logical_slides"];
                            var firstSlide = (IDictionary<string, object>)
                                slides[0];
                            Check((string)firstSlide["layout"] == "cards",
                                "The reflowed layout did not reach review.");
                            decision = json.Serialize(new
                            {
                                contract_version =
                                    AnalysisReviewContract.Version,
                                context_id = (string)content["context_id"],
                                approved = true,
                                findings = new object[0]
                            });
                        }
                    }
                    var response = json.Serialize(new
                    {
                        choices = new[] { new
                        {
                            message = new
                            {
                                role = "assistant", content = decision
                            }
                        } }
                    });
                    var bytesOut = Encoding.UTF8.GetBytes(response);
                    var headers = Encoding.ASCII.GetBytes(
                        "HTTP/1.1 200 OK\r\n" +
                        "Content-Type: application/json\r\n" +
                        "Content-Length: " + bytesOut.Length +
                        "\r\nConnection: close\r\n\r\n");
                    stream.Write(headers, 0, headers.Length);
                    stream.Write(bytesOut, 0, bytesOut.Length);
                    stream.Flush();
                }
            }
        }

        private static string PackageText(ZipArchive package,
            string path)
        {
            var entry = package.GetEntry(path);
            Check(entry != null, "Native PowerPoint package is missing " + path);
            return PackageText(entry);
        }

        private static bool HasChartRelationship(ZipArchive package,
            string path)
        {
            var entry = package.GetEntry(path);
            Check(entry != null, "Native PowerPoint package is missing " + path);
            using (var stream = entry.Open())
                return XDocument.Load(stream).Descendants().Any(node =>
                    node.Name.LocalName == "Relationship" &&
                    ((string)node.Attribute("Target") ?? string.Empty)
                        .Contains("charts/"));
        }

        private static string PackageText(ZipArchiveEntry entry)
        {
            using (var stream = entry.Open())
            {
                var document = XDocument.Load(stream);
                return string.Join("|", document.Descendants()
                    .Where(node => node.Name.LocalName == "t" ||
                        node.Name.LocalName == "v")
                    .Select(node => node.Value));
            }
        }

        private static Tuple<AnalysisArtifact, AnalysisDocumentPlan> Fixture(
            AnalysisArtifact artifact)
        {
            var mayRevenue = artifact.Facts.Single(fact =>
                fact.Metric == "RevenueEUR" && fact.Period == "2026-05");
            var juneRevenue = artifact.Facts.Single(fact =>
                fact.Metric == "RevenueEUR" && fact.Period == "2026-06");
            var mayCost = artifact.Facts.Single(fact =>
                fact.Metric == "CostEUR" && fact.Period == "2026-05");
            var juneCost = artifact.Facts.Single(fact =>
                fact.Metric == "CostEUR" && fact.Period == "2026-06");
            Func<string, AnalysisPlanCell> label = value =>
                new AnalysisPlanCell { Text = value };
            Func<VerifiedFact, AnalysisPlanCell> reference = value =>
                new AnalysisPlanCell { FactId = value.FactId };
            Func<AnalysisPlanCell[], AnalysisPlanRow> row = values =>
                new AnalysisPlanRow { Cells = values.ToList() };
            var plan = new AnalysisDocumentPlan
            {
                AnalysisId = artifact.AnalysisId,
                WorkbookTitle = "Revenue and cost audit",
                WorkbookRows = AnalysisWorkbookPlanBuilder.Build(artifact),
                Slides = new List<AnalysisPlanSlide>
                {
                    new AnalysisPlanSlide { Id = "headline", Layout = "scorecard",
                        Title = "June revenue at a glance",
                        Subtitle = new List<AnalysisPlanText> {
                            new AnalysisPlanText { Text = "Verified revenue: " },
                            new AnalysisPlanText { FactId = juneRevenue.FactId } },
                        Cards = new List<AnalysisPlanCard> {
                            new AnalysisPlanCard { Heading = "June revenue",
                                Points = new List<AnalysisPlanText> { new AnalysisPlanText { FactId = juneRevenue.FactId } } },
                            new AnalysisPlanCard { Heading = "June cost",
                                Points = new List<AnalysisPlanText> { new AnalysisPlanText { FactId = juneCost.FactId } } } } },
                    new AnalysisPlanSlide { Id = "trend", Layout = "two_pane",
                        Title = "Two period comparison",
                        TableHeaders = new List<string> { "Metric", "May", "June" },
                        TableRows = new List<AnalysisPlanRow> {
                            row(new[] { label("Revenue EUR"), reference(mayRevenue), reference(juneRevenue) }),
                            row(new[] { label("Cost EUR"), reference(mayCost), reference(juneCost) }) },
                        Chart = new AnalysisPlanChart { Type = "column", Title = "Revenue EUR",
                            Categories = new List<string> { "2026-05", "2026-06" },
                            Series = new List<AnalysisPlanSeries> { new AnalysisPlanSeries {
                                Name = "Revenue EUR", FactIds = new List<string> { mayRevenue.FactId, juneRevenue.FactId } } } } },
                    new AnalysisPlanSlide { Id = "detail", Layout = "table",
                        Title = "Revenue and cost detail",
                        TableHeaders = new List<string> { "Metric", "May", "June" },
                        TableRows = new List<AnalysisPlanRow> {
                            row(new[] { label("Revenue EUR"), reference(mayRevenue), reference(juneRevenue) }),
                            row(new[] { label("Cost EUR"), reference(mayCost), reference(juneCost) }) } },
                    new AnalysisPlanSlide { Id = "provenance", Layout = "cards",
                        Title = "What the ledger supports",
                        Cards = new List<AnalysisPlanCard> { new AnalysisPlanCard {
                            Heading = "Evidence",
                            Points = new List<AnalysisPlanText> {
                                new AnalysisPlanText { Text = "Verified workbook range and live formulas" } } } } }
                }
            };
            return Tuple.Create(artifact, plan);
        }
    }
}
