using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Office;

namespace GuardrailTests
{
    internal static class AnalysisContractTests
    {
        private static void Check(bool condition, string message)
        {
            if (!condition) throw new Exception(message);
        }

        public static void ExplicitTableBindingsIssueOnlyVerifiedFacts()
        {
            Func<TableDataset, SourceSnapshot> snapshot = table =>
                AnalysisContract.CreateSnapshot("workbook-a",
                    "excel_workbook", "revision-1", "complete_range",
                    "literal_values", new[] { Locator("workbook-a",
                        "Ledger", "A1:D3", "") }, new[] { table });
            var binding = new AnalysisTableBinding
            {
                TableId = "ledger", PeriodHeader = "Period",
                DimensionHeaders = new List<string> { "Group" },
                Metrics = new List<AnalysisMetricColumnBinding>
                {
                    new AnalysisMetricColumnBinding { Header = "RevenueEUR",
                        Metric = "RevenueEUR", Unit = "currency",
                        Currency = "EUR" },
                    new AnalysisMetricColumnBinding { Header = "CostEUR",
                        Metric = "CostEUR", Unit = "currency",
                        Currency = "EUR" }
                }
            };
            var artifact = AnalysisTableArtifactBuilder.Build(
                snapshot(MappedTable()), binding);
            Check(artifact.Facts.Count == 4 &&
                artifact.Facts.Single(fact => fact.Metric == "RevenueEUR" &&
                    fact.Period == "2026-06").Value == "82992" &&
                artifact.Facts.Single(fact => fact.Metric == "CostEUR" &&
                    fact.Period == "2026-05").Locators.Single().Cell == "D2" &&
                artifact.Facts.All(fact => fact.Status ==
                    AnalysisContract.Verified &&
                    fact.Dimensions["Group"] == "North"),
                "Typed table binding lost period, dimension, source cell, or value.");
            RejectTableBinding(() => AnalysisWorkbookPlanBuilder.Build(artifact),
                "ANALYSIS_WORKBOOK_FACTS_UNSUPPORTED");
            var plainBinding = new AnalysisTableBinding
            {
                TableId = "ledger", PeriodHeader = "Period",
                Metrics = binding.Metrics
            };
            var plainArtifact = AnalysisTableArtifactBuilder.Build(
                snapshot(MappedTable()), plainBinding);
            var workbookRows = AnalysisWorkbookPlanBuilder.Build(
                plainArtifact);
            Check(workbookRows.Count == 3 &&
                workbookRows[0].Cells[1].Text == "2026-05" &&
                workbookRows[1].Cells[1].Formula ==
                    "=SUMIF('Ledger'!$A$2:$A$3,B$3,'Ledger'!$C$2:$C$3)" &&
                workbookRows[1].Cells[1].ExpectedFactId ==
                    plainArtifact.Facts.Single(fact =>
                        fact.Metric == "RevenueEUR" &&
                        fact.Period == "2026-05").FactId,
                "Host formula plan lost the source range, period header, or fact binding.");
            var zero = MappedTable();
            var zeroCell = zero.Cells.Single(cell => cell.Reference == "C3");
            zeroCell.Value = "0";
            zeroCell.RawValue = "0";
            zeroCell.DisplayText = "0";
            Check(AnalysisTableArtifactBuilder.Build(snapshot(zero), binding)
                .Facts.Single(fact => fact.Metric == "RevenueEUR" &&
                    fact.Period == "2026-06").Value == "0",
                "A verified zero was treated as a missing metric.");
            var mislabeled = new AnalysisTableBinding
            {
                TableId = "ledger", PeriodHeader = "Period",
                Metrics = new List<AnalysisMetricColumnBinding>
                {
                    new AnalysisMetricColumnBinding { Header = "RevenueEUR",
                        Metric = "CostEUR", Unit = "currency",
                        Currency = "EUR" }
                }
            };
            RejectTableBinding(() => AnalysisTableArtifactBuilder.Build(
                snapshot(MappedTable()), mislabeled),
                "ANALYSIS_TABLE_BINDING_INVALID");
            binding.Metrics[0].Currency = "USD";
            RejectTableBinding(() => AnalysisTableArtifactBuilder.Build(
                snapshot(MappedTable()), binding),
                "ANALYSIS_TABLE_BINDING_INVALID");
            binding.Metrics[0].Currency = "EUR";
            binding.PeriodHeader = "Group";
            RejectTableBinding(() => AnalysisTableArtifactBuilder.Build(
                snapshot(MappedTable()), binding),
                "ANALYSIS_TABLE_BINDING_INVALID");
            binding.PeriodHeader = "Period";
            var duplicate = MappedTable();
            duplicate.Cells.Single(cell => cell.Reference == "A3").Value =
                "2026-05";
            RejectTableBinding(() => AnalysisTableArtifactBuilder.Build(
                snapshot(duplicate), binding), "ANALYSIS_TABLE_FACT_DUPLICATE");
            var cachedFormula = MappedTable();
            var formulaCell = cachedFormula.Cells.Single(cell =>
                cell.Reference == "C3");
            formulaCell.Formula = "=SUM(C2:C2)";
            formulaCell.Status = AnalysisContract.Unresolved;
            RejectTableBinding(() => AnalysisTableArtifactBuilder.Build(
                snapshot(cachedFormula), binding),
                "ANALYSIS_TABLE_VALUE_UNVERIFIED");
            var blank = MappedTable();
            var blankCell = blank.Cells.Single(cell => cell.Reference == "D3");
            blankCell.ValueType = AnalysisContract.MissingValue;
            blankCell.Value = string.Empty;
            blankCell.Status = AnalysisContract.Unresolved;
            RejectTableBinding(() => AnalysisTableArtifactBuilder.Build(
                snapshot(blank), binding),
                "ANALYSIS_TABLE_VALUE_UNVERIFIED");
            var ambiguous = MappedTable();
            ambiguous.Cells.Single(cell => cell.Reference == "D1").Value =
                "RevenueEUR";
            RejectTableBinding(() => AnalysisTableArtifactBuilder.Build(
                snapshot(ambiguous), binding),
                "ANALYSIS_TABLE_HEADER_AMBIGUOUS");
            var partial = AnalysisContract.CreateSnapshot("workbook-a",
                "excel_workbook", "revision-1", "partial_page",
                "literal_values", new[] { Locator("workbook-a",
                    "Ledger", "A1:D3", "") }, new[] { MappedTable() });
            RejectTableBinding(() => AnalysisTableArtifactBuilder.Build(
                partial, binding), "ANALYSIS_TABLE_SOURCE_UNSUPPORTED");
        }

        private static void RejectTableBinding(Action action, string code)
        {
            try { action(); }
            catch (InvalidOperationException error)
            {
                if (error.Message.Contains(code)) return;
                throw;
            }
            throw new Exception("Expected table binding failure " + code);
        }

        public static void RepeatedSourceIdentityIsCountedOnce()
        {
            var table = new TableDataset {
                TableId = "ledger-with-ids", Name = "Ledger", Rows = 4,
                Columns = 4, Cells = new List<DatasetCell> {
                    Cell(0, 0, "A1", AnalysisContract.TextValue, "RowID"),
                    Cell(0, 1, "B1", AnalysisContract.TextValue, "Period"),
                    Cell(0, 2, "C1", AnalysisContract.TextValue, "HoursWorked"),
                    Cell(0, 3, "D1", AnalysisContract.TextValue, "HoursAvailable"),
                    Cell(1, 0, "A2", AnalysisContract.TextValue, "record-one"),
                    Cell(1, 1, "B2", AnalysisContract.TextValue, "2026-05"),
                    Cell(1, 2, "C2", AnalysisContract.DecimalValue, "10"),
                    Cell(1, 3, "D2", AnalysisContract.DecimalValue, "12"),
                    Cell(2, 0, "A3", AnalysisContract.TextValue, "record-two"),
                    Cell(2, 1, "B3", AnalysisContract.TextValue, "2026-06"),
                    Cell(2, 2, "C3", AnalysisContract.DecimalValue, "20"),
                    Cell(2, 3, "D3", AnalysisContract.DecimalValue, "25"),
                    Cell(3, 0, "A4", AnalysisContract.TextValue, "record-two"),
                    Cell(3, 1, "B4", AnalysisContract.TextValue, "2026-06"),
                    Cell(3, 2, "C4", AnalysisContract.DecimalValue, "20"),
                    Cell(3, 3, "D4", AnalysisContract.DecimalValue, "25") } };
            Func<TableDataset, SourceSnapshot> snapshot = source =>
                AnalysisContract.CreateSnapshot("workbook-with-ids",
                    "excel_workbook", "revision-1", "complete_range",
                    "literal_values", new[] { Locator("workbook-with-ids",
                        "Ledger", "A1:D4", "") }, new[] { source });
            var binding = new AnalysisTableBinding {
                TableId = table.TableId, PeriodHeader = "Period",
                Metrics = new List<AnalysisMetricColumnBinding> {
                    new AnalysisMetricColumnBinding { Header = "HoursWorked",
                        Metric = "HoursWorked" },
                    new AnalysisMetricColumnBinding { Header = "HoursAvailable",
                        Metric = "HoursAvailable" } } };
            var artifact = AnalysisTableArtifactBuilder.BuildGrouped(
                snapshot(table), binding);
            var june = artifact.Facts.Single(fact =>
                fact.Metric == "HoursWorked" && fact.Period == "2026-06");
            var rows = AnalysisWorkbookPlanBuilder.Build(artifact);
            var receipt = WorkbookGroupedTotals.Compute(new[] {
                (IReadOnlyList<string>)new[] { "RowID", "Period",
                    "HoursWorked", "HoursAvailable" },
                new[] { "record-one", "2026-05", "10", "12" },
                new[] { "record-two", "2026-06", "20", "25" },
                new[] { "record-two", "2026-06", "20", "25" }
            }, new[] { "Period" }, new[] { "HoursWorked" },
                null, null);
            Check(june.Value == "20" && june.Locators.Count == 1 &&
                june.Locators.Single().Cell == "C3" &&
                receipt.Table.Contains("2026-06\t1\t20") &&
                rows[1].Cells[2].Formula ==
                    "=SUMIF('Ledger'!$B$2:$B$4,C$3,'Ledger'!$C$2:$C$4)-SUM('Ledger'!$C$4)" &&
                rows[2].Cells[2].ExpectedFactId == artifact.Facts.Single(
                    fact => fact.Metric == "HoursAvailable" &&
                    fact.Period == "2026-06").FactId,
                "A repeated RowID was counted twice in facts or live formulas.");
            var deck = AnalysisDeckPlanBuilder.Build(artifact,
                new Dictionary<string, object> {
                    { "AnalysisId", artifact.AnalysisId } }, 3,
                "Compare May vs June with a primary chart and hours in the title.");
            Check(deck.Slides.Single(slide => slide.Chart != null)
                .Chart.Title.Contains("hours"),
                "A requested noncurrency chart unit was lost.");
            Check(deck.Slides.SelectMany(slide => slide.Cards)
                .SelectMany(card => card.Points)
                .Any(point => point.Text != null &&
                    point.Text.Contains("identical IDs count once")) &&
                !deck.Slides.SelectMany(slide => slide.Cards)
                    .SelectMany(card => card.Points)
                    .Any(point => point.Text != null &&
                        point.Text.Contains("totals count each row")),
                "The source-limit narrative must describe deduplicated totals accurately.");
            RejectTableBinding(() => AnalysisRequestPlan.Resolve(artifact,
                "Compare May vs June with a primary chart and EUR in the title."),
                "ANALYSIS_REQUEST_CHART_UNIT_UNBOUND");
            table.Cells.Single(cell => cell.Reference == "D4").Value = "26";
            RejectTableBinding(() => AnalysisTableArtifactBuilder.BuildGrouped(
                snapshot(table), binding), "ANALYSIS_TABLE_ROW_ID_CONFLICT");
        }

        public static void GroupedTypedFactsKeepSourceAndFormulaBinding()
        {
            var table = MappedTable();
            table.Rows = 4;
            table.Cells.Add(Cell(3, 0, "A4",
                AnalysisContract.TextValue, "2026-05"));
            table.Cells.Add(Cell(3, 1, "B4",
                AnalysisContract.TextValue, "South"));
            table.Cells.Add(Cell(3, 2, "C4",
                AnalysisContract.DecimalValue, "1"));
            table.Cells.Add(Cell(3, 3, "D4",
                AnalysisContract.DecimalValue, "2"));
            var snapshot = AnalysisContract.CreateSnapshot("workbook-a",
                "excel_workbook", "revision-1", "complete_range",
                "literal_values", new[] { Locator("workbook-a",
                    "Ledger", "A1:D4", "") }, new[] { table });
            var binding = new AnalysisTableBinding {
                TableId = "ledger", PeriodHeader = "Period",
                Metrics = new List<AnalysisMetricColumnBinding> {
                    new AnalysisMetricColumnBinding { Header = "RevenueEUR",
                        Metric = "RevenueEUR", Unit = "currency",
                        Currency = "EUR" },
                    new AnalysisMetricColumnBinding { Header = "CostEUR",
                        Metric = "CostEUR", Unit = "currency",
                        Currency = "EUR" } } };
            var artifact = AnalysisTableArtifactBuilder.BuildGrouped(
                snapshot, binding);
            var may = artifact.Facts.Single(fact =>
                fact.Metric == "RevenueEUR" &&
                fact.Period == "2026-05");
            Check(artifact.Facts.Count == 4 && may.Value == "85520" &&
                may.Locators.Count == 2 && may.Locators.Any(locator =>
                    locator.Cell == "C4"),
                "Grouped facts lost an additive row or native locator.");
            var rows = AnalysisWorkbookPlanBuilder.Build(artifact);
            Check(rows[1].Cells[1].Formula ==
                "=SUMIF('Ledger'!$A$2:$A$4,B$3,'Ledger'!$C$2:$C$4)" &&
                rows[1].Cells[1].ExpectedFactId == may.FactId,
                "The grouped report lost its source-bound formula.");
            var historicalTable = MappedTable();
            historicalTable.Rows = 5;
            historicalTable.Cells.Add(Cell(3, 0, "A4",
                AnalysisContract.TextValue, "2026-04"));
            historicalTable.Cells.Add(Cell(3, 1, "B4",
                AnalysisContract.TextValue, "North"));
            historicalTable.Cells.Add(Cell(3, 2, "C4",
                AnalysisContract.DecimalValue, "80000"));
            historicalTable.Cells.Add(Cell(3, 3, "D4",
                AnalysisContract.DecimalValue, "30000"));
            historicalTable.Cells.Add(Cell(4, 0, "A5",
                AnalysisContract.TextValue, "2026-03"));
            historicalTable.Cells.Add(Cell(4, 1, "B5",
                AnalysisContract.TextValue, "North"));
            historicalTable.Cells.Add(Cell(4, 2, "C5",
                AnalysisContract.DecimalValue, "75000"));
            historicalTable.Cells.Add(Cell(4, 3, "D5",
                AnalysisContract.DecimalValue, "28000"));
            var historicalSnapshot = AnalysisContract.CreateSnapshot(
                "workbook-b", "excel_workbook", "revision-1",
                "complete_range", "literal_values", new[] {
                    Locator("workbook-b", "Ledger", "A1:D5", "") },
                new[] { historicalTable });
            var historicalArtifact = AnalysisTableArtifactBuilder.BuildGrouped(
                historicalSnapshot, binding);
            var comparisonRows = AnalysisWorkbookPlanBuilder.Build(
                historicalArtifact);
            Check(comparisonRows[0].Cells.Count == 3 &&
                comparisonRows[0].Cells[1].Text == "2026-05" &&
                comparisonRows[0].Cells[2].Text == "2026-06" &&
                comparisonRows[1].Cells[1].ExpectedFactId ==
                    historicalArtifact.Facts.Single(fact =>
                        fact.Metric == "RevenueEUR" &&
                        fact.Period == "2026-05").FactId &&
                comparisonRows[1].Cells[2].ExpectedFactId ==
                    historicalArtifact.Facts.Single(fact =>
                        fact.Metric == "RevenueEUR" &&
                        fact.Period == "2026-06").FactId,
                "The comparison did not put the latest verified periods in B and C.");
            var requestedSelection = AnalysisRequestPlan.Resolve(
                historicalArtifact,
                "Compare March vs April and chart with both series.");
            var priorPilot = AnalysisDocumentPilot.Enabled;
            try
            {
                AnalysisDocumentPilot.SetEnabled(true);
                var draftRequest = DocumentChatRequestFactory.Create(
                    "test-model", "excel", "Ledger", new List<ChatTurn>(),
                    "Create a new draft worksheet from the verified Ledger and preserve its source.",
                    true);
                var draftBoundary = Convert.ToString(draftRequest.messages
                    .OfType<ChatCompletionInputMessage>()
                    .First(message => message.role == "system").content);
                Check(draftBoundary.Contains("read_grouped_totals") &&
                    draftBoundary.Contains("complete bound source") &&
                    !draftBoundary.Contains("read it to the END"),
                    "Typed aggregate reads should cover the bound source without requiring row-by-row enumeration.");
                Check(!draftRequest.tools.Any(tool => tool.function.name ==
                    WorkbookToolCatalog.WriteCells) &&
                    draftRequest.tools.Any(tool =>
                        tool.function.name == WorkbookToolCatalog.WriteDraftSheet &&
                        ((Dictionary<string, object>)tool.function.parameters)[
                            "required"] is string[] &&
                        ((string[])((Dictionary<string, object>)tool.function.parameters)[
                            "required"]).Contains("analysis_id")),
                    "The pilot exposed a generic writer before source binding.");
                DocumentChatRequestFactory.ApplyAnalysisPilot(draftRequest,
                    historicalArtifact, "excel");
                Check(!draftRequest.tools.Any(tool => tool.function.name ==
                    WorkbookToolCatalog.WriteCells) &&
                    draftRequest.tools.Any(tool => tool.function.name ==
                        WorkbookToolCatalog.WriteDraftSheet),
                    "The typed analysis route still exposed a source-bound cell edit.");
            }
            finally { AnalysisDocumentPilot.SetEnabled(priorPilot); }
            var requestedRows = AnalysisWorkbookPlanBuilder.Build(
                historicalArtifact, requestedSelection);
            Check(requestedSelection.ComparePeriod == "2026-03" &&
                requestedSelection.FocusPeriod == "2026-04" &&
                requestedRows[0].Cells[1].Text == "2026-03" &&
                requestedRows[0].Cells[2].Text == "2026-04" &&
                requestedRows[1].Cells[1].ExpectedFactId ==
                    historicalArtifact.Facts.Single(fact =>
                        fact.Metric == "RevenueEUR" &&
                        fact.Period == "2026-03").FactId &&
                requestedRows[1].Cells[2].ExpectedFactId ==
                    historicalArtifact.Facts.Single(fact =>
                        fact.Metric == "RevenueEUR" &&
                        fact.Period == "2026-04").FactId,
                "The user-requested non-latest comparison was not bound to the source.");
            var requestedDeck = AnalysisDeckPlanBuilder.Build(
                historicalArtifact, new Dictionary<string, object> {
                    { "AnalysisId", historicalArtifact.AnalysisId },
                    { "ComparePeriod", "2026-03" },
                    { "FocusPeriod", "2026-04" },
                    { "ChartSeries", new[] { "RevenueEUR", "CostEUR" } }
                }, 3, "Compare March vs April and chart with both series.");
            var requestedChart = requestedDeck.Slides.Single(slide =>
                slide.Chart != null).Chart;
            Check(requestedDeck.ComparePeriod == "2026-03" &&
                requestedDeck.FocusPeriod == "2026-04" &&
                requestedChart.Categories.SequenceEqual(new[] {
                    "2026-03", "2026-04" }) &&
                requestedChart.Series.Count == 2 &&
                requestedChart.Series[0].FactIds[0] ==
                    historicalArtifact.Facts.Single(fact =>
                        fact.Metric == "RevenueEUR" &&
                        fact.Period == "2026-03").FactId &&
                requestedChart.Series[1].FactIds[1] ==
                    historicalArtifact.Facts.Single(fact =>
                        fact.Metric == "CostEUR" &&
                        fact.Period == "2026-04").FactId,
                "The user-requested second chart series was dropped.");
            RejectTableBinding(() => AnalysisRequestPlan.Resolve(
                historicalArtifact, "Compare February vs April."),
                "ANALYSIS_REQUEST_PERIOD_UNBOUND");
            RejectTableBinding(() => AnalysisRequestPlan.Resolve(
                historicalArtifact, "Review March results."),
                "ANALYSIS_REQUEST_COMPARISON_UNBOUND");
            RejectTableBinding(() => AnalysisRequestPlan.Resolve(
                historicalArtifact, "Compare March vs April.",
                new Dictionary<string, object> {
                    { "FocusPeriod", "2026-06" } }),
                "ANALYSIS_REQUEST_PERIOD_MISMATCH");
            RejectTableBinding(() => AnalysisRequestPlan.Resolve(
                historicalArtifact,
                "Compare March vs April and chart with both series.",
                new Dictionary<string, object> {
                    { "ChartSeries", new[] { "RevenueEUR", "Unknown" } } }),
                "ANALYSIS_REQUEST_CHART_SERIES_UNBOUND");
            var extraMetricTable = new TableDataset {
                TableId = "extra-metrics", Name = "Ledger", Rows = 3,
                Columns = 4, Cells = new List<DatasetCell> {
                    Cell(0, 0, "A1", AnalysisContract.TextValue, "Period"),
                    Cell(0, 1, "B1", AnalysisContract.TextValue, "Units"),
                    Cell(0, 2, "C1", AnalysisContract.TextValue, "RevenueEUR"),
                    Cell(0, 3, "D1", AnalysisContract.TextValue, "CostEUR"),
                    Cell(1, 0, "A2", AnalysisContract.TextValue, "2026-05"),
                    Cell(1, 1, "B2", AnalysisContract.DecimalValue, "943"),
                    Cell(1, 2, "C2", AnalysisContract.DecimalValue, "85519"),
                    Cell(1, 3, "D2", AnalysisContract.DecimalValue, "36702"),
                    Cell(2, 0, "A3", AnalysisContract.TextValue, "2026-06"),
                    Cell(2, 1, "B3", AnalysisContract.DecimalValue, "971"),
                    Cell(2, 2, "C3", AnalysisContract.DecimalValue, "82992"),
                    Cell(2, 3, "D3", AnalysisContract.DecimalValue, "36714") } };
            var extraMetricSnapshot = AnalysisContract.CreateSnapshot(
                "workbook-extra", "excel_workbook", "revision-1",
                "complete_range", "literal_values", new[] {
                    Locator("workbook-extra", "Ledger", "A1:D3", "") },
                new[] { extraMetricTable });
            var extraMetricBinding = new AnalysisTableBinding {
                TableId = extraMetricTable.TableId,
                PeriodHeader = "Period",
                Metrics = new List<AnalysisMetricColumnBinding> {
                    new AnalysisMetricColumnBinding { Header = "RevenueEUR",
                        Metric = "RevenueEUR", Unit = "currency",
                        Currency = "EUR" },
                    new AnalysisMetricColumnBinding { Header = "CostEUR",
                        Metric = "CostEUR", Unit = "currency",
                        Currency = "EUR" },
                    new AnalysisMetricColumnBinding { Header = "Units",
                        Metric = "Units", Unit = "units" } } };
            var extraMetricArtifact = AnalysisTableArtifactBuilder.BuildGrouped(
                extraMetricSnapshot, extraMetricBinding);
            var extraMetricSelection = AnalysisRequestPlan.Resolve(
                extraMetricArtifact,
                "Create a June-to-May audit table. Put Revenue EUR in A4 and Cost EUR in A5.");
            var extraMetricRows = AnalysisWorkbookPlanBuilder.Build(
                extraMetricArtifact, extraMetricSelection);
            Check(extraMetricRows.Count == 3 &&
                extraMetricSelection.ReportMetrics.SequenceEqual(new[] {
                    "RevenueEUR", "CostEUR" }) &&
                extraMetricRows[1].Cells[0].Text == "RevenueEUR" &&
                extraMetricRows[2].Cells[0].Text == "CostEUR" &&
                extraMetricRows[1].Cells[1].ExpectedFactId ==
                    extraMetricArtifact.Facts.Single(fact =>
                        fact.Metric == "RevenueEUR" &&
                        fact.Period == "2026-05").FactId,
                "The report replaced requested output metrics with source-column order.");
            var extraMetricDeck = AnalysisDeckPlanBuilder.Build(
                extraMetricArtifact, new Dictionary<string, object> {
                    { "AnalysisId", extraMetricArtifact.AnalysisId } }, 3,
                "Create a period comparison with a native chart. The chart must use only primary values for May and June.");
            var extraMetricChart = extraMetricDeck.Slides.Single(slide =>
                slide.Chart != null).Chart;
            Check(extraMetricDeck.ReportMetrics.SequenceEqual(new[] {
                    "RevenueEUR", "CostEUR" }) &&
                extraMetricChart.Series.Count == 1 &&
                extraMetricChart.Series[0].FactIds[0] ==
                    extraMetricArtifact.Facts.Single(fact =>
                        fact.Metric == "RevenueEUR" &&
                        fact.Period == "2026-05").FactId,
                "The primary chart series followed source-column order instead of the typed binding.");
            var reorderedDraft = AnalysisDeckPlanBuilder.Build(
                extraMetricArtifact, new Dictionary<string, object> {
                    { "AnalysisId", extraMetricArtifact.AnalysisId } }, 3,
                "Create a May vs June chart using only primary values.",
                new[] { "Cost EUR", "Revenue EUR" });
            Check(reorderedDraft.ReportMetrics.SequenceEqual(new[] {
                    "CostEUR", "RevenueEUR" }) &&
                reorderedDraft.ChartSeries.SequenceEqual(new[] {
                    "CostEUR" }) &&
                reorderedDraft.Slides.Single(slide => slide.Chart != null)
                    .Chart.Series[0].FactIds[0] ==
                    extraMetricArtifact.Facts.Single(fact =>
                        fact.Metric == "CostEUR" &&
                        fact.Period == "2026-05").FactId,
                "The primary chart ignored the verified draft metric order.");
            RejectTableBinding(() => AnalysisRequestPlan.Resolve(
                extraMetricArtifact, "Compare May vs June with a chart.",
                null, new[] { "Unknown metric" }),
                "ANALYSIS_REQUEST_REPORT_METRIC_UNBOUND");
            binding.DimensionHeaders.Add("Group");
            var dimensioned = AnalysisTableArtifactBuilder.BuildGrouped(
                snapshot, binding, "Period", "2026-05");
            Check(dimensioned.Facts.Count == 6 &&
                dimensioned.Facts.Any(fact => fact.Metric == "CostEUR" &&
                    fact.Dimensions.ContainsKey("Group") &&
                    fact.Dimensions["Group"] == "South" &&
                    fact.Value == "2") &&
                dimensioned.Facts.Any(fact => fact.Metric == "CostEUR" &&
                    fact.Dimensions.Count == 0 && fact.Value == "36704"),
                "Grouped facts lost a dimension or verified period total.");
            var allGroups = AnalysisTableArtifactBuilder.BuildGrouped(
                snapshot, binding);
            var groupRequest = AnalysisRequestPlan.Resolve(allGroups,
                "Compare May vs June with a primary chart.");
            Check(allGroups.Facts.Count == 10 &&
                groupRequest.ComparePeriod == "2026-05" &&
                groupRequest.FocusPeriod == "2026-06" &&
                allGroups.Facts.Any(fact => fact.Metric == "RevenueEUR" &&
                    fact.Dimensions.Count == 0 &&
                    fact.Period == "2026-05" && fact.Value == "85520"),
                "A dimensional read did not retain complete period totals for planning.");
            var subsetBinding = new AnalysisTableBinding {
                TableId = binding.TableId, PeriodHeader = "Period",
                Metrics = new List<AnalysisMetricColumnBinding> {
                    binding.Metrics[1] } };
            var subset = AnalysisTableArtifactBuilder.BuildGrouped(snapshot,
                subsetBinding, "Group", "South");
            Check(subset.Facts.Count == 1 &&
                subset.Facts.Single().Dimensions["Group"] == "South" &&
                subset.Facts.Single().Value == "2",
                "A filtered subtotal was mislabeled as a complete period total.");
            var latestGroups = AnalysisTableArtifactBuilder.BuildGrouped(
                snapshot, binding, "Period", "2026-06");
            var combined = AnalysisContract.CreateArtifact(
                new[] { snapshot }, artifact.Facts.Concat(
                    latestGroups.Facts).GroupBy(fact => fact.FactId,
                        StringComparer.Ordinal).Select(group => group.First()),
                    new AnalysisCalculation[0],
                new string[0], new string[0]);
            var deck = AnalysisDeckPlanBuilder.Build(combined,
                new Dictionary<string, object> {
                    { "AnalysisId", combined.AnalysisId },
                    { "Title", "Verified sales review" },
                    { "Lead", "Period results from source rows" } }, 4);
            Check(deck.Slides.Count == 4 &&
                deck.Slides.Any(slide => slide.Chart != null &&
                    slide.Chart.Series.Count == 1 &&
                    slide.Chart.Series[0].FactIds.SequenceEqual(new[] {
                        combined.Facts.Single(fact =>
                            fact.Dimensions.Count == 0 &&
                            fact.Metric == "RevenueEUR" &&
                            fact.Period == "2026-05").FactId,
                        combined.Facts.Single(fact =>
                            fact.Dimensions.Count == 0 &&
                            fact.Metric == "RevenueEUR" &&
                            fact.Period == "2026-06").FactId })) &&
                deck.Slides.Any(slide => slide.TableRows.Count > 0 &&
                    slide.TableRows.Any(item => item.Cells.Any(cell =>
                        cell.FactId != null))) &&
                AnalysisDocumentCompiler.Compile(combined, deck).Slides.Count
                    == deck.Slides.Count,
                "Host deck planning lost native chart, group evidence or source-bound facts.");
            var missing = MappedTable();
            missing.Cells.Single(cell => cell.Reference == "C3")
                .Status = AnalysisContract.Unresolved;
            var missingSnapshot = AnalysisContract.CreateSnapshot(
                "workbook-a", "excel_workbook", "revision-1",
                "complete_range", "literal_values", new[] {
                    Locator("workbook-a", "Ledger", "A1:D3", "") },
                new[] { missing });
            RejectTableBinding(() =>
                AnalysisTableArtifactBuilder.BuildGrouped(missingSnapshot,
                    binding), "ANALYSIS_TABLE_VALUE_UNVERIFIED");

            var partial = MappedTable();
            partial.Rows = 4;
            partial.Cells.Add(Cell(3, 0, "A4",
                AnalysisContract.TextValue, "2026-05"));
            partial.Cells.Add(Cell(3, 1, "B4",
                AnalysisContract.TextValue, "South"));
            partial.Cells.Add(Cell(3, 2, "C4",
                AnalysisContract.DecimalValue, "1"));
            var blankCost = Cell(3, 3, "D4",
                AnalysisContract.MissingValue, string.Empty);
            blankCost.Status = AnalysisContract.Unresolved;
            partial.Cells.Add(blankCost);
            var partialSnapshot = AnalysisContract.CreateSnapshot(
                "workbook-a", "excel_workbook", "partial-revision",
                "complete_range", "literal_values", new[] {
                    Locator("workbook-a", "Ledger", "A1:D4", "") },
                new[] { partial });
            var known = AnalysisTableArtifactBuilder.BuildGrouped(
                partialSnapshot, binding);
            Check(known.Facts.Single(fact => fact.Metric == "CostEUR" &&
                    fact.Period == "2026-05" &&
                    fact.Dimensions.Count == 0).Value == "36702" &&
                known.UnresolvedConflicts.Any(item => item.Contains(
                    "Ledger!D4")),
                "A blank source value was imputed or the known subtotal was not disclosed.");
            var knownRows = AnalysisWorkbookPlanBuilder.Build(known);
            var knownPlan = new AnalysisDocumentPlan {
                AnalysisId = known.AnalysisId, WorkbookTitle = "Known subtotals",
                WorkbookRows = knownRows };
            Check(knownRows.Last().Cells[0].Text == "Data quality" &&
                knownRows.Last().Cells[1].Text.Contains("CostEUR: known subtotal") &&
                AnalysisDocumentCompiler.Compile(known, knownPlan, false)
                    .ExpectedFormulaFacts.Count == 4,
                "A verified known subtotal could not produce a disclosed live-formula draft.");
            var blocking = AnalysisContract.CreateArtifact(
                known.Snapshots, known.Facts,
                new AnalysisCalculation[0], new string[0],
                new[] { "Unverified source value remains." });
            RejectTableBinding(() => AnalysisWorkbookPlanBuilder.Build(blocking),
                "ANALYSIS_WORKBOOK_SOURCE_UNSUPPORTED");

            var formulaTable = new TableDataset { TableId = "formula-ledger",
                Name = "Ledger", Rows = 2, Columns = 4,
                Cells = new List<DatasetCell> {
                    Cell(0, 0, "A1", AnalysisContract.TextValue, "Period"),
                    Cell(0, 1, "B1", AnalysisContract.TextValue, "Units"),
                    Cell(0, 2, "C1", AnalysisContract.TextValue, "Price"),
                    Cell(0, 3, "D1", AnalysisContract.TextValue, "RevenueEUR"),
                    Cell(1, 0, "A2", AnalysisContract.TextValue, "2026-06"),
                    Cell(1, 1, "B2", AnalysisContract.DecimalValue, "3"),
                    Cell(1, 2, "C2", AnalysisContract.DecimalValue, "4"),
                    Cell(1, 3, "D2", AnalysisContract.DecimalValue, "12") } };
            var formulaCell = formulaTable.Cells.Last();
            formulaCell.Formula = "=IF(C2=\"\",\"\",B2*C2)";
            formulaCell.Status = AnalysisContract.Unresolved;
            var formulaBinding = new AnalysisTableBinding {
                TableId = formulaTable.TableId, PeriodHeader = "Period",
                Metrics = new List<AnalysisMetricColumnBinding> {
                    new AnalysisMetricColumnBinding { Header = "RevenueEUR",
                        Metric = "RevenueEUR", Currency = "EUR",
                        Unit = "currency" } } };
            Func<SourceSnapshot> captureFormula = () =>
                AnalysisContract.CreateSnapshot("workbook-a",
                    "excel_workbook", "formula-revision", "complete_range",
                    "cached_formula_values_unverified", new[] {
                        Locator("workbook-a", "Ledger", "A1:D2", "") },
                    new[] { formulaTable });
            var verifiedFormula = AnalysisTableArtifactBuilder.BuildGrouped(
                captureFormula(), formulaBinding);
            Check(verifiedFormula.Facts.Single().Value == "12",
                "A recomputed source formula was not bound to a typed fact.");
            formulaCell.Value = "13";
            RejectTableBinding(() => AnalysisTableArtifactBuilder.BuildGrouped(
                captureFormula(), formulaBinding),
                "ANALYSIS_TABLE_VALUE_UNVERIFIED");
            formulaCell.Value = "12";
            formulaCell.Formula = "=SUM(B2:C2)";
            RejectTableBinding(() => AnalysisTableArtifactBuilder.BuildGrouped(
                captureFormula(), formulaBinding),
                "ANALYSIS_TABLE_VALUE_UNVERIFIED");

            var balanceTable = new TableDataset { TableId = "balance-ledger",
                Name = "Ledger", Rows = 2, Columns = 7,
                Cells = new List<DatasetCell> {
                    Cell(0, 0, "A1", AnalysisContract.TextValue, "Period"),
                    Cell(0, 1, "B1", AnalysisContract.TextValue, "Opening"),
                    Cell(0, 2, "C1", AnalysisContract.TextValue, "Received"),
                    Cell(0, 3, "D1", AnalysisContract.TextValue, "Shipped"),
                    Cell(0, 4, "E1", AnalysisContract.TextValue, "Closing"),
                    Cell(0, 5, "F1", AnalysisContract.TextValue, "PriceEUR"),
                    Cell(0, 6, "G1", AnalysisContract.TextValue, "ValueEUR"),
                    Cell(1, 0, "A2", AnalysisContract.TextValue, "2026-06"),
                    Cell(1, 1, "B2", AnalysisContract.DecimalValue, "10"),
                    Cell(1, 2, "C2", AnalysisContract.DecimalValue, "4"),
                    Cell(1, 3, "D2", AnalysisContract.DecimalValue, "2"),
                    Cell(1, 4, "E2", AnalysisContract.DecimalValue, "12"),
                    Cell(1, 5, "F2", AnalysisContract.DecimalValue, "4"),
                    Cell(1, 6, "G2", AnalysisContract.DecimalValue, "48") } };
            var closingCell = balanceTable.Cells.Single(cell =>
                cell.Reference == "E2");
            closingCell.Formula = "=IF(C2=\"\",\"\",B2+C2-D2)";
            closingCell.Status = AnalysisContract.Unresolved;
            var valueCell = balanceTable.Cells.Single(cell =>
                cell.Reference == "G2");
            valueCell.Formula = "=IF(E2=\"\",\"\",E2*F2)";
            valueCell.Status = AnalysisContract.Unresolved;
            var balanceBinding = new AnalysisTableBinding {
                TableId = balanceTable.TableId, PeriodHeader = "Period",
                Metrics = new List<AnalysisMetricColumnBinding> {
                    new AnalysisMetricColumnBinding { Header = "Closing",
                        Metric = "Closing" },
                    new AnalysisMetricColumnBinding { Header = "ValueEUR",
                        Metric = "ValueEUR", Currency = "EUR",
                        Unit = "currency" } } };
            Func<SourceSnapshot> captureBalance = () =>
                AnalysisContract.CreateSnapshot("workbook-a",
                    "excel_workbook", "balance-revision", "complete_range",
                    "cached_formula_values_unverified", new[] {
                        Locator("workbook-a", "Ledger", "A1:G" +
                            balanceTable.Rows.ToString(
                                CultureInfo.InvariantCulture), "") },
                    new[] { balanceTable });
            var balanced = AnalysisTableArtifactBuilder.BuildGrouped(
                captureBalance(), balanceBinding);
            Check(balanced.Facts.Single(fact => fact.Metric == "Closing")
                    .Value == "12" && balanced.Facts.Single(fact =>
                    fact.Metric == "ValueEUR").Value == "48",
                "Chained same-row formulas were not verified against literal inputs.");
            closingCell.Value = "13";
            RejectTableBinding(() => AnalysisTableArtifactBuilder.BuildGrouped(
                captureBalance(), balanceBinding),
                "ANALYSIS_TABLE_VALUE_UNVERIFIED");
            closingCell.Value = "12";
            closingCell.Formula = "=IF(C3=\"\",\"\",B2+C2-D2)";
            RejectTableBinding(() => AnalysisTableArtifactBuilder.BuildGrouped(
                captureBalance(), balanceBinding),
                "ANALYSIS_TABLE_VALUE_UNVERIFIED");
            closingCell.Formula = "=IF(B2=\"\",\"\",B2-D2)";
            closingCell.Value = "8";
            valueCell.Value = "32";
            Check(AnalysisTableArtifactBuilder.BuildGrouped(captureBalance(),
                balanceBinding).Facts.Single(fact => fact.Metric ==
                    "Closing").Value == "8",
                "A verified same-row difference was rejected.");
            closingCell.Formula =
                "=IF(B2=\"\",\"\",IF(D2=0,\"\",B2/D2))";
            closingCell.Value = "5";
            valueCell.Value = "20";
            Check(AnalysisTableArtifactBuilder.BuildGrouped(captureBalance(),
                balanceBinding).Facts.Single(fact => fact.Metric ==
                    "Closing").Value == "5",
                "A guarded same-row ratio was rejected.");
            closingCell.Value = "6";
            RejectTableBinding(() => AnalysisTableArtifactBuilder.BuildGrouped(
                captureBalance(), balanceBinding),
                "ANALYSIS_TABLE_VALUE_UNVERIFIED");
            closingCell.Formula = "=IF(C2=\"\",\"\",B2+C2-D2)";
            closingCell.Value = "12";
            valueCell.Value = "48";
            balanceTable.Rows = 3;
            balanceTable.Cells.Add(Cell(2, 0, "A3",
                AnalysisContract.TextValue, "2026-06"));
            balanceTable.Cells.Add(Cell(2, 1, "B3",
                AnalysisContract.DecimalValue, "10"));
            balanceTable.Cells.Add(Cell(2, 2, "C3",
                AnalysisContract.MissingValue, string.Empty));
            balanceTable.Cells.Add(Cell(2, 3, "D3",
                AnalysisContract.DecimalValue, "2"));
            var missingClosing = Cell(2, 4, "E3",
                AnalysisContract.TextValue, string.Empty);
            missingClosing.Formula =
                "=IF(C3=\"\",\"\",B3+C3-D3)";
            missingClosing.Status = AnalysisContract.Unresolved;
            balanceTable.Cells.Add(missingClosing);
            balanceTable.Cells.Add(Cell(2, 5, "F3",
                AnalysisContract.DecimalValue, "4"));
            var missingValue = Cell(2, 6, "G3",
                AnalysisContract.TextValue, string.Empty);
            missingValue.Formula = "=IF(E3=\"\",\"\",E3*F3)";
            missingValue.Status = AnalysisContract.Unresolved;
            balanceTable.Cells.Add(missingValue);
            var knownBalance = AnalysisTableArtifactBuilder.BuildGrouped(
                captureBalance(), balanceBinding);
            Check(knownBalance.Facts.Single(fact => fact.Metric ==
                    "Closing").Value == "12" &&
                knownBalance.Facts.Single(fact => fact.Metric ==
                    "ValueEUR").Value == "48" &&
                knownBalance.UnresolvedConflicts.Count == 2,
                "A guarded blank formula was imputed or its known subtotal was not disclosed.");

            var many = new TableDataset { TableId = "many-rows",
                Name = "Ledger", Rows = 49, Columns = 4,
                Cells = new List<DatasetCell> {
                    Cell(0, 0, "A1", AnalysisContract.TextValue, "Period"),
                    Cell(0, 1, "B1", AnalysisContract.TextValue, "RevenueEUR"),
                    Cell(0, 2, "C1", AnalysisContract.TextValue, "CostEUR"),
                    Cell(0, 3, "D1", AnalysisContract.TextValue, "RowID") } };
            for (var row = 1; row < many.Rows; row++)
            {
                var addressRow = (row + 1).ToString(
                    CultureInfo.InvariantCulture);
                many.Cells.Add(Cell(row, 0, "A" + addressRow,
                    AnalysisContract.TextValue, row <= 24 ?
                        "2026-05" : "2026-06"));
                many.Cells.Add(Cell(row, 1, "B" + addressRow,
                    AnalysisContract.DecimalValue, "100"));
                many.Cells.Add(Cell(row, 2, "C" + addressRow,
                    AnalysisContract.DecimalValue, "50"));
                many.Cells.Add(Cell(row, 3, "D" + addressRow,
                    AnalysisContract.TextValue, "WB01-" +
                    row.ToString("D4", CultureInfo.InvariantCulture)));
            }
            var manySnapshot = AnalysisContract.CreateSnapshot("workbook-a",
                "excel_workbook", "many-rows-revision", "complete_range",
                "literal_values", new[] { Locator("workbook-a",
                    "Ledger", "A1:D49", "") }, new[] { many });
            var manyBinding = new AnalysisTableBinding {
                TableId = many.TableId, PeriodHeader = "Period",
                Metrics = new List<AnalysisMetricColumnBinding> {
                    new AnalysisMetricColumnBinding { Header = "RevenueEUR",
                        Metric = "RevenueEUR", Currency = "EUR",
                        Unit = "currency" },
                    new AnalysisMetricColumnBinding { Header = "CostEUR",
                        Metric = "CostEUR", Currency = "EUR",
                        Unit = "currency" } } };
            var manyArtifact = AnalysisTableArtifactBuilder.BuildGrouped(
                manySnapshot, manyBinding);
            var manyPlan = AnalysisNativeAcceptance.Fixture(
                manyArtifact).Item2;
            manyPlan.Slides[0].Subtitle[0].Text =
                "Source WB01, June 2026. ";
            manyPlan.Slides[0].Cards[0].Points[0].Text =
                "Verified revenue: ";
            var manyCompiled = AnalysisDocumentCompiler.Compile(
                manyArtifact, manyPlan);
            var headline = manyCompiled.Slides[0];
            Check(((string)headline["sources"]).Contains("B26:B49") &&
                ((string)headline["sources"]).Length < 2000 &&
                ((string)headline["subtitle"]).Contains("WB01") &&
                ((string)headline["subtitle"]).Contains("2026") &&
                ((object[])((Dictionary<string, object>)
                    ((object[])headline["cards"])[0])["points"])[0]
                    .ToString().Contains("Verified revenue"),
                "Verified labels or exact grouped-cell citations did not survive deck compilation.");
        }

        public static void SlideQualityUsesBoundFindingsAndSafeFooters()
        {
            var binding = new AnalysisTableBinding {
                TableId = "ledger", PeriodHeader = "Period",
                DimensionHeaders = new List<string> { "Group" },
                Metrics = new List<AnalysisMetricColumnBinding> {
                    new AnalysisMetricColumnBinding { Header = "RevenueEUR",
                        Metric = "RevenueEUR", Unit = "currency",
                        Currency = "EUR" },
                    new AnalysisMetricColumnBinding { Header = "CostEUR",
                        Metric = "CostEUR", Unit = "currency",
                        Currency = "EUR" } } };
            Func<string, TableDataset, AnalysisArtifact> capture =
                (path, table) => AnalysisTableArtifactBuilder.BuildGrouped(
                    AnalysisContract.CreateSnapshot(path,
                        "excel_workbook", "revision-1", "complete_range",
                        "literal_values", new[] { Locator(path,
                            "Ledger", "A1:D3", "") },
                        new[] { table }), binding);
            var first = capture(@"excel:C:\Users\example\Sales.xlsx",
                MappedTable());
            var changed = MappedTable();
            changed.Cells.Single(cell => cell.Reference == "C3")
                .Value = "100000";
            var second = capture(@"excel:C:\Users\example\Other.xlsx",
                changed);
            Func<AnalysisArtifact, AnalysisDocumentPlan> build = artifact =>
                AnalysisDeckPlanBuilder.Build(artifact,
                    new Dictionary<string, object> {
                        { "AnalysisId", artifact.AnalysisId } }, 4,
                    "Compare May vs June with a primary revenue chart.");
            var a = build(first);
            var b = build(second);
            var json = new JavaScriptSerializer();
            var compiledA = AnalysisDocumentCompiler.Compile(first, a);
            var compiledB = AnalysisDocumentCompiler.Compile(second, b);
            var titlesA = compiledA.Slides.Select(slide =>
                Convert.ToString(slide["title"])).ToArray();
            var titlesB = compiledB.Slides.Select(slide =>
                Convert.ToString(slide["title"])).ToArray();
            Check(!titlesA.SequenceEqual(titlesB) &&
                titlesA[0].Contains("82,992") &&
                titlesB[0].Contains("100,000") &&
                titlesA[0].Contains("down") &&
                titlesB[0].Contains("up"),
                "Different workbooks produced the same headline findings.");
            var pages = ((System.Collections.IEnumerable)json.DeserializeObject(
                json.Serialize(SamsungPresentationReview.InspectPlan(
                    json.Serialize(compiledA.Slides)))))
                .Cast<Dictionary<string, object>>().ToArray();
            var visible = pages.SelectMany(page =>
                ((System.Collections.IEnumerable)page["elements"])
                    .Cast<Dictionary<string, object>>())
                .Select(element => Convert.ToString(element["text"]))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Replace("[Scribble draft]", ""))
                .ToArray();
            Check(!visible.Any(value => Regex.IsMatch(value,
                    @"\b(?:verified|captured|bound|native|draft)\b|evidence boundary",
                    RegexOptions.IgnoreCase)),
                "System narration leaked onto a slide.");
            Check(!visible.Any(value => Regex.IsMatch(value,
                    @"[A-Za-z]:[\\/]|[/\\]Users[/\\]",
                    RegexOptions.IgnoreCase)) &&
                visible.Any(value => value.Contains(
                    "Sales.xlsx / Ledger!")),
                "A slide footer lost the workbook name or exposed a local path.");
        }

        private static TableDataset MappedTable()
        {
            return new TableDataset
            {
                TableId = "ledger", Name = "Ledger", Rows = 3,
                Columns = 4, Cells = new List<DatasetCell>
                {
                    Cell(0, 0, "A1", AnalysisContract.TextValue, "Period"),
                    Cell(0, 1, "B1", AnalysisContract.TextValue, "Group"),
                    Cell(0, 2, "C1", AnalysisContract.TextValue, "RevenueEUR"),
                    Cell(0, 3, "D1", AnalysisContract.TextValue, "CostEUR"),
                    Cell(1, 0, "A2", AnalysisContract.TextValue, "2026-05"),
                    Cell(1, 1, "B2", AnalysisContract.TextValue, "North"),
                    Cell(1, 2, "C2", AnalysisContract.DecimalValue, "85519"),
                    Cell(1, 3, "D2", AnalysisContract.DecimalValue, "36702"),
                    Cell(2, 0, "A3", AnalysisContract.TextValue, "2026-06"),
                    Cell(2, 1, "B3", AnalysisContract.TextValue, "North"),
                    Cell(2, 2, "C3", AnalysisContract.DecimalValue, "82992"),
                    Cell(2, 3, "D3", AnalysisContract.DecimalValue, "36714")
                }
            };
        }

        public static void SnapshotIdentityInvalidationAndSerialization()
        {
            var table = Table();
            var locatorA = Locator("workbook-a", "Ledger", "A1:D2", "");
            var first = AnalysisContract.CreateSnapshot(
                "workbook-a",
                "excel_workbook",
                "revision-1",
                "complete_range",
                "recalculated",
                new[] { locatorA },
                new[] { table });
            var repeat = AnalysisContract.CreateSnapshot(
                "workbook-a",
                "excel_workbook",
                "revision-1",
                "complete_range",
                "recalculated",
                new[] { locatorA },
                new[] { Table() });
            var identicalOtherSource = AnalysisContract.CreateSnapshot(
                "workbook-b",
                "excel_workbook",
                "revision-1",
                "complete_range",
                "recalculated",
                new[] { Locator("workbook-b", "Ledger", "A1:D2", "") },
                new[] { Table() });
            var changedRevision = AnalysisContract.CreateSnapshot(
                "workbook-a",
                "excel_workbook",
                "revision-2",
                "complete_range",
                "recalculated",
                new[] { locatorA },
                new[] { Table() });

            Check(first.ContentHash == repeat.ContentHash &&
                first.SnapshotId == repeat.SnapshotId,
                "A stable source revision did not produce a stable snapshot binding.");
            Check(first.ContentHash == identicalOtherSource.ContentHash &&
                first.SnapshotId != identicalOtherSource.SnapshotId,
                "Identical workbook content collapsed two source instances.");
            Check(first.SnapshotId != changedRevision.SnapshotId &&
                !AnalysisContract.IsCurrent(first, "workbook-a", "revision-2",
                    changedRevision.ContentHash),
                "A changed source revision did not invalidate the prior snapshot.");
            Check(table.Cells.Count == 8 &&
                table.Cells.Single(cell => cell.Reference == "B2").ValueType ==
                    AnalysisContract.MissingValue,
                "A blank table cell was dropped instead of retaining its column position.");

            var revenue = Fact(first, "RevenueEUR", "82992", "EUR");
            var cost = Fact(first, "CostEUR", "36714", "EUR");
            VerifiedFact margin;
            var calculation = AnalysisCalculator.Calculate(
                AnalysisCalculator.Margin,
                new[] { revenue, cost },
                "GrossMargin",
                "2026-06",
                out margin);
            var artifact = AnalysisContract.CreateArtifact(
                new[] { first },
                new[] { revenue, cost, margin },
                new[] { calculation },
                new[] { "Revenue and cost use the same period." },
                new string[0]);
            var serialized = AnalysisContract.Serialize(artifact);
            var restored = AnalysisContract.Deserialize(serialized);
            Check(restored.AnalysisId == artifact.AnalysisId &&
                restored.Facts.Single(fact => fact.Metric == "GrossMargin")
                    .Value.StartsWith("0.557", StringComparison.Ordinal) &&
                restored.Snapshots[0].Tables[0].Cells.Count == 8,
                "The typed analysis did not survive persistence without value loss.");

            var root = Path.Combine(Path.GetTempPath(),
                "scribble-analysis-contract-" + Guid.NewGuid().ToString("N"));
            try
            {
                var request = new ChatCompletionRequest
                {
                    model = "offline-test",
                    messages = new List<object>
                    {
                        new ChatCompletionInputMessage
                        {
                            role = "user",
                            content = "Persist a typed analysis"
                        }
                    }
                };
                var task = new TaskContextManager(request, "excel",
                    "Persist a typed analysis", new TaskCheckpointStore(root));
                var evidenceId = task.PersistAnalysis(artifact);
                var loaded = task.LoadAnalysis();
                Check(task.State.AnalysisContractVersion ==
                        AnalysisContract.Version &&
                    task.State.AnalysisArtifactEvidenceId == evidenceId &&
                    loaded != null && loaded.AnalysisId == artifact.AnalysisId,
                    "Task persistence did not bind the typed analysis version and protected evidence.");
                var extendedFact = Fact(first, "OperatingEUR", "10", "EUR");
                var extended = AnalysisContract.CreateArtifact(
                    new[] { first },
                    artifact.Facts.Concat(new[] { extendedFact }),
                    artifact.Calculations, artifact.Assumptions,
                    artifact.UnresolvedConflicts);
                task.PersistAnalysis(extended);
                Check(task.AcceptsAnalysisId(artifact.AnalysisId, extended) &&
                    task.AcceptsAnalysisId(extended.AnalysisId, extended) &&
                    !task.AcceptsAnalysisId("analysis_fabricated", extended),
                    "A prior ID for the same source was not retained safely after an additive read.");
                var changed = AnalysisContract.CreateArtifact(
                    new[] { changedRevision }, new[] {
                        Fact(changedRevision, "RevenueEUR", "82992", "EUR") },
                    new AnalysisCalculation[0], new string[0],
                    new string[0]);
                task.PersistAnalysis(changed);
                Check(!task.AcceptsAnalysisId(artifact.AnalysisId, changed) &&
                    !task.AcceptsAnalysisId(extended.AnalysisId, changed),
                    "A prior analysis ID survived a source revision change.");

                var sources = new TaskSources(task);
                var firstText = sources.Add("Attached document", "Same source text", "attachment:0");
                var secondText = sources.Add("Attached document", "Same source text", "attachment:1");
                var repeatedFirst = sources.Add("Attached document", "Same source text", "attachment:0");
                var changedFirst = sources.Add("Attached document", "Changed source text", "attachment:0");
                Check(firstText.Count == 1 && secondText.Count == 1 &&
                    firstText[0] != secondText[0] &&
                    repeatedFirst.SequenceEqual(firstText) &&
                    changedFirst[0] != firstText[0] &&
                    sources.Resolve(firstText) == "Same source text" &&
                    sources.Resolve(secondText) == "Same source text" &&
                    sources.Spans().Count == 3,
                    "Identical attachments lost distinct provenance, or changed content reused an old source span.");
            }
            finally
            {
                if (Directory.Exists(root)) Directory.Delete(root, true);
            }

            restored.Facts.Single(fact => fact.Metric == "RevenueEUR").Value =
                "99999";
            var tamperRejected = false;
            try { AnalysisContract.Serialize(restored); }
            catch (InvalidOperationException error)
            {
                tamperRejected = error.Message.Contains("FACT_ID_MISMATCH");
            }
            Check(tamperRejected,
                "A fact value changed in place without invalidating its host-issued ID.");
        }

        public static void DeterministicCalculationsPreserveAuthority()
        {
            var snapshot = AnalysisContract.CreateSnapshot(
                "workbook-calculations",
                "excel_workbook",
                "revision-1",
                "complete_range",
                "recalculated",
                new[] { Locator("workbook-calculations", "Ledger", "A1:D2", "") },
                new[] { Table() });
            var june = Fact(snapshot, "RevenueEUR", "82992", "EUR");
            var may = Fact(snapshot, "RevenueEUR", "85519", "EUR");
            var groupA = Fact(snapshot, "RevenueEUR", "22675", "EUR");
            var groupB = Fact(snapshot, "RevenueEUR", "22044", "EUR");

            VerifiedFact difference;
            AnalysisCalculator.Calculate(AnalysisCalculator.Difference,
                new[] { june, may }, "RevenueChange", "2026-06", out difference);
            VerifiedFact growth;
            AnalysisCalculator.Calculate(AnalysisCalculator.Growth,
                new[] { june, may }, "RevenueGrowth", "2026-06", out growth);
            Check(difference.Value == "-2527" &&
                growth.Value.StartsWith("-0.0295", StringComparison.Ordinal),
                "Host arithmetic did not preserve the exact May-to-June direction.");

            var ranking = AnalysisCalculator.RankDescending(
                new[] { groupB, groupA });
            Check(ranking.Single(item => item.FactId == groupA.FactId).Rank == 1 &&
                ranking.Single(item => item.FactId == groupB.FactId).Rank == 2,
                "Deterministic ranking accepted the incorrect 22,044 > 22,675 objection.");

            var zero = Fact(snapshot, "RevenueEUR", "0", "EUR");
            var zeroRejected = false;
            try
            {
                VerifiedFact ignored;
                AnalysisCalculator.Calculate(AnalysisCalculator.Ratio,
                    new[] { june, zero }, "Ratio", "2026-06", out ignored);
            }
            catch (InvalidOperationException error)
            {
                zeroRejected = error.Message.Contains("ZERO_DENOMINATOR");
            }
            Check(zeroRejected, "A zero denominator produced a verified ratio.");

            var usd = Fact(snapshot, "Revenue", "100", "USD");
            var mismatchRejected = false;
            try
            {
                VerifiedFact ignored;
                AnalysisCalculator.Calculate(AnalysisCalculator.Sum,
                    new[] { june, usd }, "MixedRevenue", "2026-06", out ignored);
            }
            catch (InvalidOperationException error)
            {
                mismatchRejected = error.Message.Contains("UNIT_MISMATCH");
            }
            Check(mismatchRejected,
                "A mixed-currency calculation was promoted to a verified fact.");

            var missing = AnalysisContract.CreateObservedFact(
                snapshot.SnapshotId,
                "RevenueEUR",
                AnalysisContract.MissingValue,
                string.Empty,
                string.Empty,
                "currency",
                "EUR",
                "2026-07",
                new Dictionary<string, string>(),
                snapshot.Locators,
                AnalysisContract.Unresolved);
            var missingRejected = false;
            try
            {
                VerifiedFact ignored;
                AnalysisCalculator.Calculate(AnalysisCalculator.Sum,
                    new[] { june, missing }, "RevenueEUR", "2026-H2", out ignored);
            }
            catch (InvalidOperationException error)
            {
                missingRejected = error.Message.Contains("NOT_VERIFIED_NUMERIC");
            }
            Check(missingRejected,
                "A missing source value was silently treated as zero.");

            var values = new object[2, 4]
            {
                { "RowID", "Period", "Revenue", "Approved" },
                { "R1", 45808d, 82992d, true }
            };
            var formulas = new object[2, 4]
            {
                { "RowID", "Period", "Revenue", "Approved" },
                { "R1", 45808d, "=SUM(Source!C2:C10)", true }
            };
            var formats = new object[2, 4]
            {
                { "General", "General", "General", "General" },
                { "General", "yyyy-mm", "#,##0", "General" }
            };
            var displayed = new object[2, 4]
            {
                { "RowID", "Period", "Revenue", "Approved" },
                { "R1", "2025-05", "82,992", "TRUE" }
            };
            var typed = WorkbookTypedCapture.Capture("live", "Ledger",
                values, formulas, formats, displayed, 2, 4, 1, 1);
            var period = typed.Cells.Single(cell => cell.Reference == "B2");
            var formula = typed.Cells.Single(cell => cell.Reference == "C2");
            var approved = typed.Cells.Single(cell => cell.Reference == "D2");
            Check(period.ValueType == AnalysisContract.DateValue &&
                period.RawValue == "45808" &&
                period.DisplayText == "2025-05" &&
                formula.Formula == "=SUM(Source!C2:C10)" &&
                formula.RawValue == "82992" &&
                formula.Status == AnalysisContract.Unresolved &&
                approved.ValueType == AnalysisContract.BooleanValue &&
                approved.Value == "true",
                "Typed workbook capture lost raw values, display text, formulas, formats, or booleans.");

            var aggregateValues = new object[3, 3]
            {
                { "Group", "Period", "RevenueEUR" },
                { "A", "2026-06", 10d },
                { "A", "2026-06", 20d }
            };
            var aggregateFormulas = new object[3, 3]
            {
                { null, null, null },
                { null, null, null },
                { null, null, "=SUM(Source!C2:C10)" }
            };
            var aggregate = WorkbookTypedCapture.Capture(
                "typed-groups", "Ledger", aggregateValues,
                aggregateFormulas, "General", null, 3, 3, 1, 1);
            var totals = WorkbookGroupedTotals.Compute(aggregate,
                new[] { "Group" }, new[] { "RevenueEUR" },
                "Period", "2026-06");
            Check(totals.SourceRows == 2 && totals.MatchedRows == 2 &&
                totals.Groups == 1 && totals.SkippedCells == 1 &&
                totals.Table.Contains("2026-06\tA\t2\t10\t1"),
                "Typed grouped totals trusted an unresolved formula cache or lost its skipped-cell disclosure.");
        }

        public static void OpenXmlCaptureRetainsTypedCells()
        {
            var root = Path.Combine(Path.GetTempPath(),
                "scribble-openxml-analysis-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                var path = Path.Combine(root, "typed.xlsx");
                using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
                {
                    Entry(archive, "xl/workbook.xml",
                        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>" +
                        "<sheet name=\"Ledger\" sheetId=\"2\" r:id=\"rId2\"/>" +
                        "<sheet name=\"Notes\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                    Entry(archive, "xl/_rels/workbook.xml.rels",
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Target=\"worksheets/sheet1.xml\" " +
                        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\"/>" +
                        "<Relationship Id=\"rId2\" Target=\"worksheets/sheet2.xml\" " +
                        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\"/>" +
                        "</Relationships>");
                    Entry(archive, "xl/sharedStrings.xml",
                        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                        "<si><t>Period</t></si><si><t>Revenue</t></si><si><t>Approved</t></si></sst>");
                    Entry(archive, "xl/styles.xml",
                        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                        "<numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"#,##0 &quot;EUR&quot;\"/></numFmts>" +
                        "<cellXfs count=\"3\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\"/><xf numFmtId=\"164\"/></cellXfs>" +
                        "</styleSheet>");
                    Entry(archive, "xl/worksheets/sheet1.xml",
                        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" +
                        "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Source notes</t></is></c></row>" +
                        "</sheetData></worksheet>");
                    Entry(archive, "xl/worksheets/sheet2.xml",
                        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" +
                        "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c>" +
                        "<c r=\"D1\" t=\"s\"><v>2</v></c></row>" +
                        "<row r=\"2\"><c r=\"A2\" s=\"1\"><v>45808</v></c>" +
                        "<c r=\"B2\" s=\"2\"><f>SUM(Source!B2:B10)</f><v>82992</v></c>" +
                        "<c r=\"C2\"/><c r=\"D2\" t=\"b\"><v>1</v></c></row>" +
                        "<row r=\"3\"><c r=\"A3\" t=\"b\"><v>maybe</v></c></row>" +
                        "</sheetData></worksheet>");
                }
                var snapshot = OpenXmlWorkbookSnapshotReader.Capture(path,
                    "attachment:typed", "revision-1", CancellationToken.None);
                var table = snapshot.Tables[0];
                var date = table.Cells.Single(cell => cell.Reference == "A2");
                var formula = table.Cells.Single(cell => cell.Reference == "B2");
                var blank = table.Cells.Single(cell => cell.Reference == "C2");
                var approved = table.Cells.Single(cell => cell.Reference == "D2");
                var malformed = table.Cells.Single(cell => cell.Reference == "A3");
                Check(snapshot.Tables.Count == 2 && table.Name == "Ledger" &&
                    snapshot.Tables[1].Name == "Notes" && table.Rows == 3 &&
                    table.Columns == 4 && date.ValueType == AnalysisContract.DateValue &&
                    date.RawValue == "45808" && date.NumberFormat == "m/d/yy" &&
                    formula.Formula == "=SUM(Source!B2:B10)" &&
                    formula.RawCellType == "n" && formula.RawValue == "82992" &&
                    formula.NumberFormat.Contains("EUR") &&
                    formula.Status == AnalysisContract.Unresolved &&
                    blank.ValueType == AnalysisContract.MissingValue &&
                    approved.ValueType == AnalysisContract.BooleanValue &&
                    malformed.ValueType == AnalysisContract.ErrorValue &&
                    malformed.RawCellType == "b" &&
                    malformed.Status == AnalysisContract.Unresolved &&
                    snapshot.CalculationState == "cached_formula_values_unverified",
                    "OpenXML typed capture lost a sheet name, blank column, date serial, formula, format, boolean, malformed value, or cache status.");
            }
            finally
            {
                if (Directory.Exists(root)) Directory.Delete(root, true);
            }
        }

        public static void MixedExcelFormatsResolveDates()
        {
            var values = new object[,] { { "Period", "Amount" },
                { 45808d, 120d } };
            var columnReads = 0;
            var cellReads = 0;
            var formats = WorkbookTypedCapture.ResolveMixedNumberFormats(
                DBNull.Value, 2, 2,
                column => { columnReads++; return column == 0
                    ? DBNull.Value : (object)"#,##0"; },
                (row, column) => { cellReads++; return row == 1
                    ? (object)"m/d/yy" : "General"; });
            var table = WorkbookTypedCapture.Capture("mixed", "Ledger",
                values, null, formats, null, 2, 2, 1, 1);
            Check(columnReads == 2 && cellReads == 2 &&
                table.Cells.Single(cell => cell.Reference == "A2")
                    .ValueType == AnalysisContract.DateValue &&
                table.Cells.Single(cell => cell.Reference == "B2")
                    .ValueType == AnalysisContract.DecimalValue,
                "Excel DBNull mixed formats did not resolve the real date column.");
            Check(WorkbookTypedCapture.ResolveMixedNumberFormats(
                    DBNull.Value, 501, 1, _ => "General",
                    (_, __) => "General") == null,
                "Oversized mixed-format range should remain incomplete.");
            var rejected = false;
            try { WorkbookTypedCapture.Capture("mixed", "Ledger",
                values, null, DBNull.Value, null, 2, 2, 1, 1); }
            catch (InvalidOperationException exception)
            { rejected = exception.Message.Contains("ANALYSIS_NUMBER_FORMAT_INCOMPLETE"); }
            Check(rejected, "Unresolved DBNull formats were silently typed as numbers.");
        }

        private static void Entry(
            ZipArchive archive,
            string name,
            string content)
        {
            var entry = archive.CreateEntry(name);
            using (var writer = new StreamWriter(entry.Open(),
                new UTF8Encoding(false))) writer.Write(content);
        }

        private static TableDataset Table()
        {
            return new TableDataset
            {
                TableId = "ledger",
                Name = "Ledger",
                Rows = 2,
                Columns = 4,
                Cells = new List<DatasetCell>
                {
                    Cell(0, 0, "A1", AnalysisContract.TextValue, "RowID"),
                    Cell(0, 1, "B1", AnalysisContract.TextValue, "Group"),
                    Cell(0, 2, "C1", AnalysisContract.TextValue, "RevenueEUR"),
                    Cell(0, 3, "D1", AnalysisContract.TextValue, "Period"),
                    Cell(1, 0, "A2", AnalysisContract.TextValue, "R1"),
                    Cell(1, 1, "B2", AnalysisContract.MissingValue, string.Empty),
                    Cell(1, 2, "C2", AnalysisContract.DecimalValue, "0"),
                    Cell(1, 3, "D2", AnalysisContract.TextValue, "2026-06")
                }
            };
        }

        private static DatasetCell Cell(
            int row,
            int column,
            string reference,
            string valueType,
            string value)
        {
            return new DatasetCell
            {
                Row = row,
                Column = column,
                Reference = reference,
                ValueType = valueType,
                RawCellType = valueType == AnalysisContract.MissingValue
                    ? "blank"
                    : valueType == AnalysisContract.DecimalValue ||
                      valueType == AnalysisContract.IntegerValue
                        ? "number"
                        : "text",
                RawValue = value,
                Value = value,
                DisplayText = value,
                Formula = string.Empty,
                NumberFormat = "General",
                Status = valueType == AnalysisContract.MissingValue
                    ? AnalysisContract.Unresolved
                    : AnalysisContract.Verified
            };
        }

        private static SourceLocator Locator(
            string instance,
            string sheet,
            string range,
            string cell)
        {
            return new SourceLocator
            {
                Kind = "excel_range",
                SourceInstanceId = instance,
                WorksheetIdentity = sheet,
                Range = range,
                Cell = cell
            };
        }

        private static VerifiedFact Fact(
            SourceSnapshot snapshot,
            string metric,
            string value,
            string currency)
        {
            return AnalysisContract.CreateObservedFact(
                snapshot.SnapshotId,
                metric,
                AnalysisContract.DecimalValue,
                value,
                value,
                "currency",
                currency,
                "2026-06",
                new Dictionary<string, string>(),
                snapshot.Locators,
                AnalysisContract.Verified);
        }
    }
}
