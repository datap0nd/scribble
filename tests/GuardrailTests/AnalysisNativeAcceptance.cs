using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Web.Script.Serialization;
using System.Xml.Linq;
using Scribble.Chat;
using Scribble.Office;

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
            dynamic excel = null, workbook = null, powerPoint = null, deck = null;
            var failure = string.Empty;
            var workbookPassed = false;
            var slidesPassed = false;
            var sourcePreserved = false;
            var recoveryPassed = false;
            var typedReviewPassed = false;
            var rendererRepairPassed = false;
            var images = new List<string>();
            var stage = "setup";
            var output = Path.GetDirectoryName(Path.GetFullPath(reportPath));
            Directory.CreateDirectory(output);
            var priorFlag = Environment.GetEnvironmentVariable(
                AnalysisDocumentPilot.FeatureFlag);
            try
            {
                Environment.SetEnvironmentVariable(
                    AnalysisDocumentPilot.FeatureFlag, "1");
                stage = "excel_start";
                excel = Activator.CreateInstance(Type.GetTypeFromProgID(
                    "Excel.Application", true));
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
                readTask.AfterTool(readCall, readResult);
                var bound = readTask.LoadAnalysis();
                Check(bound != null && bound.Facts.Count == 4 &&
                    bound.Facts.Any(fact => fact.Metric == "RevenueEUR" &&
                        fact.Period == "2026-05" && fact.Value == "85519" &&
                        fact.Locators[0].Cell == "I2") &&
                    readStore.Load(readTask.State.Id).AnalysisArtifactEvidenceId ==
                        readTask.State.AnalysisArtifactEvidenceId,
                    "The task did not durably retain facts from native Excel cells.");
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
                stage = "excel_source_freshness";
                ledger.Range("I2").Value2 = 85520d;
                var staleRejected = false;
                try
                {
                    AnalysisDocumentPilot.WriteWorkbook((object)excel,
                        fixture.Item1, fixture.Item2);
                }
                catch (InvalidOperationException error)
                {
                    staleRejected = error.Message.Contains(
                        "ANALYSIS_SOURCE_CHANGED");
                }
                finally { ledger.Range("I2").Value2 = 85519d; }
                Check(staleRejected && (int)workbook.Worksheets.Count == 1,
                    "A changed source cell created a draft before freshness validation.");
                stage = "excel_write_and_readback";
                AnalysisDocumentPilot.WriteWorkbook((object)excel,
                    fixture.Item1, fixture.Item2);
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
                stage = "powerpoint_write";
                AnalysisDocumentPilot.WritePresentation((object)powerPoint,
                    fixture.Item1, fixture.Item2);
                deck = powerPoint.ActivePresentation;
                Check((int)deck.Slides.Count == 4,
                    "The native deck did not contain exactly four slides.");
                for (var index = 1; index <= 4; index++)
                {
                    dynamic slide = deck.Slides[index];
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
                var nativeReview = AnalysisDocumentPilot.ReserveNativeReview(
                    reviewTask, (object)deck, fixture.Item1,
                    fixture.Item2, true);
                var pages = nativeReview.Context.Pages;
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
                Check(staleTableReviewRejected &&
                    AnalysisDocumentPilot.CompleteNativeReview((object)deck,
                        nativeReview, cleanVerdict).Approved,
                    "A native table-cell edit inherited an earlier review approval.");
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
                Check(staleChartReviewRejected &&
                    AnalysisDocumentPilot.CompleteNativeReview((object)deck,
                        nativeReview, cleanVerdict).Approved,
                    "A native chart-series edit inherited an earlier review approval.");
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
                nativeChart.Left = (float)table.Left + (float)table.Width - 40f;
                var collidedPages = AnalysisDocumentPilot.CapturePresentationPages(
                    (object)deck, fixture.Item1, fixture.Item2);
                var collision = AnalysisDocumentPilot.CaptureNativeMeasurements(
                    (object)deck, collidedPages).Single(item =>
                        item.Code == "COLLISION" &&
                        item.NativeSlideId == collidedPages[1].NativeSlideId &&
                        item.TargetId == "shape:" + (int)nativeChart.Id &&
                        item.OtherTargetId == "shape:" + (int)table.Id);
                repairReceipt = AnalysisDocumentPilot.RepairNativeMeasurement(
                    (object)deck, collidedPages, collision, repairReceipt);
                var clearedPages = AnalysisDocumentPilot.CapturePresentationPages(
                    (object)deck, fixture.Item1, fixture.Item2);
                Check(AnalysisDocumentPilot.CaptureNativeMeasurements(
                    (object)deck, clearedPages).Count == 0 &&
                    AnalysisRepairBudget.Read(repairReceipt).PatchedTargets.Count == 3,
                    "The renderer did not clear and read back a chart/table collision.");
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
                    Check(headline.Contains(
                            fixture.Item1.Snapshots[0].SourceInstanceId),
                        "The rendered slide lost its host-derived citation.");
                }
                slidesPassed = true;
            }
            catch (Exception error) { failure = stage + ": " + error; }
            finally
            {
                Environment.SetEnvironmentVariable(
                    AnalysisDocumentPilot.FeatureFlag, priorFlag);
                if ((object)deck != null) try { deck.Close(); } catch { }
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
                rendered_images = images,
                full_acceptance_passed = false,
                note = "Hand-authored structural pilot only; no model, visual attestation, or recovery qualification.",
                failure
            };
            var json = new JavaScriptSerializer().Serialize(report);
            File.WriteAllText(reportPath, json);
            Console.WriteLine(json);
            return workbookPassed && slidesPassed && sourcePreserved &&
                recoveryPassed && typedReviewPassed && rendererRepairPassed
                ? 0 : 1;
        }

        private static string SourceFingerprint(dynamic sheet)
        {
            return string.Join("|", new[] { "B2", "I2", "J2", "B3", "I3", "J3" }
                .Select(cell => Convert.ToString(sheet.Range(cell).Value2,
                    CultureInfo.InvariantCulture)));
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
