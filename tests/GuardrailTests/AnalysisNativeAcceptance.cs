using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
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
                var fixture = Fixture();
                var compiled = AnalysisDocumentCompiler.Compile(
                    fixture.Item1, fixture.Item2);
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
                var chartCount = 0;
                var captured = new List<Dictionary<string, object>>();
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
                    stage = "powerpoint_capture_" + index;
                    captured.Add(PresentationInspection.Capture((object)slide));
                    foreach (dynamic shape in slide.Shapes)
                        if (Convert.ToInt32(shape.HasChart) != 0)
                            chartCount++;
                }
                Check(chartCount >= 1,
                    "The comparison slide did not contain a native chart.");
                stage = "powerpoint_oracle";
                var headline = PresentationInspection.CitationTextFromCaptured(
                    captured[0]);
                var comparison = PresentationInspection.CitationTextFromCaptured(
                    captured[1]);
                Check(headline.Contains("82,992") &&
                    headline.Contains("36,714") &&
                    comparison.Contains("85,519") &&
                    comparison.Contains("82,992"),
                    "The rendered slide content differed from the independent fact oracle.");
                var nativeCharts = ((IEnumerable)captured[1]["shapes"])
                    .Cast<object>().OfType<Dictionary<string, object>>()
                    .Where(shape => shape.ContainsKey("chart"))
                    .Select(shape => new JavaScriptSerializer().Serialize(
                        shape["chart"])).ToArray();
                Check(nativeCharts.Any(chart => chart.Contains("85519") &&
                    chart.Contains("82992")),
                    "The native chart data did not match the verified periods.");
                var firstNotes = PresentationInspection.Notes((object)deck.Slides[1]);
                Check(firstNotes.Contains("WB01"),
                    "The rendered slide lost its host-derived citation.");
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
                rendered_images = images,
                full_acceptance_passed = false,
                note = "Hand-authored structural pilot only; no model, visual attestation, or recovery qualification.",
                failure
            };
            var json = new JavaScriptSerializer().Serialize(report);
            File.WriteAllText(reportPath, json);
            Console.WriteLine(json);
            return workbookPassed && slidesPassed && sourcePreserved ? 0 : 1;
        }

        private static string SourceFingerprint(dynamic sheet)
        {
            return string.Join("|", new[] { "B2", "I2", "J2", "B3", "I3", "J3" }
                .Select(cell => Convert.ToString(sheet.Range(cell).Value2,
                    CultureInfo.InvariantCulture)));
        }

        private static Tuple<AnalysisArtifact, AnalysisDocumentPlan> Fixture()
        {
            var locator = new SourceLocator
            {
                Kind = "excel_range", SourceInstanceId = "WB01",
                WorksheetIdentity = "Ledger", Range = "B1:J3"
            };
            var snapshot = AnalysisContract.CreateSnapshot("WB01",
                "excel_workbook", "native-pilot-1", "complete_range",
                "recalculated", new[] { locator }, new TableDataset[0]);
            Func<string, string, string, VerifiedFact> fact = (metric, value, period) =>
                AnalysisContract.CreateObservedFact(snapshot.SnapshotId,
                    metric, AnalysisContract.DecimalValue, value, value,
                    "currency", "EUR", period, new Dictionary<string, string>(),
                    new[] { locator }, AnalysisContract.Verified);
            var mayRevenue = fact("RevenueEUR", "85519", "2026-05");
            var juneRevenue = fact("RevenueEUR", "82992", "2026-06");
            var mayCost = fact("CostEUR", "36702", "2026-05");
            var juneCost = fact("CostEUR", "36714", "2026-06");
            var artifact = AnalysisContract.CreateArtifact(new[] { snapshot },
                new[] { mayRevenue, juneRevenue, mayCost, juneCost },
                new AnalysisCalculation[0], new string[0], new string[0]);
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
                WorkbookRows = new List<AnalysisPlanRow>
                {
                    row(new[] { label("Metric"), label("May"), label("June") }),
                    row(new[] { label("Revenue EUR"),
                        new AnalysisPlanCell { Formula = "=SUMIF(Ledger!$B$2:$B$3,\"2026-05\",Ledger!$I$2:$I$3)", ExpectedFactId = mayRevenue.FactId },
                        new AnalysisPlanCell { Formula = "=SUMIF(Ledger!$B$2:$B$3,\"2026-06\",Ledger!$I$2:$I$3)", ExpectedFactId = juneRevenue.FactId } }),
                    row(new[] { label("Cost EUR"),
                        new AnalysisPlanCell { Formula = "=SUMIF(Ledger!$B$2:$B$3,\"2026-05\",Ledger!$J$2:$J$3)", ExpectedFactId = mayCost.FactId },
                        new AnalysisPlanCell { Formula = "=SUMIF(Ledger!$B$2:$B$3,\"2026-06\",Ledger!$J$2:$J$3)", ExpectedFactId = juneCost.FactId } })
                },
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
