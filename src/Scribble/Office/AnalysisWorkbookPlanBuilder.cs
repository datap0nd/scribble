using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Scribble.Office
{
    // The pilot report is host-owned: a model can arrange narrative slides,
    // but cannot author formulas or type replacement numeric results.
    public static class AnalysisWorkbookPlanBuilder
    {
        private static readonly Regex CellAddress = new Regex(
            @"^([A-Z]+)([1-9][0-9]*)$", RegexOptions.CultureInvariant);

        public static List<AnalysisPlanRow> Build(AnalysisArtifact artifact)
        {
            return Build(artifact, null);
        }

        public static List<AnalysisPlanRow> Build(AnalysisArtifact artifact,
            AnalysisRequestPlan selection)
        {
            AnalysisContract.Serialize(artifact);
            if (artifact.Snapshots.Count != 1 ||
                artifact.Snapshots[0].Tables.Count != 1 ||
                artifact.Facts.Count == 0 ||
                artifact.Calculations.Count != 0 ||
                artifact.Assumptions.Count != 0 ||
                AnalysisContract.HasBlockingConflicts(artifact))
                throw new InvalidOperationException(
                    "ANALYSIS_WORKBOOK_SOURCE_UNSUPPORTED");
            var snapshot = artifact.Snapshots[0];
            var table = snapshot.Tables[0];
            if (snapshot.Coverage != "complete_range" ||
                table.Rows < 2 || table.Columns < 2)
                throw new InvalidOperationException(
                    "ANALYSIS_WORKBOOK_SOURCE_UNSUPPORTED");
            var headerCells = table.Cells.Where(cell => cell.Row == 0 &&
                    !string.IsNullOrWhiteSpace(cell.Value)).ToArray();
            if (headerCells.GroupBy(cell => cell.Value,
                    StringComparer.Ordinal).Any(group => group.Count() != 1))
                throw new InvalidOperationException(
                    "ANALYSIS_WORKBOOK_HEADER_AMBIGUOUS");
            var headers = headerCells.ToDictionary(cell => cell.Value,
                    cell => cell.Column,
                    StringComparer.Ordinal);
            int periodColumn;
            if (!headers.TryGetValue("Period", out periodColumn) ||
                string.IsNullOrWhiteSpace(table.Name) ||
                table.Name.IndexOfAny(new[] { '[', ']', '\\' }) >= 0)
                throw new InvalidOperationException(
                    "ANALYSIS_WORKBOOK_PERIOD_UNSUPPORTED");
            var sourcePeriods = table.Cells.Where(cell =>
                    cell.Row > 0 && cell.Column == periodColumn)
                .OrderBy(cell => cell.Row).ToArray();
            if (sourcePeriods.Length != table.Rows - 1)
                throw new InvalidOperationException(
                    "ANALYSIS_WORKBOOK_PERIOD_UNSUPPORTED");
            var allPeriods = sourcePeriods.GroupBy(cell => cell.Value,
                    StringComparer.Ordinal).Select(group => group.First())
                .OrderBy(cell => cell.Value, StringComparer.Ordinal).ToArray();
            var reportFacts = artifact.Facts.Where(fact =>
                fact.Dimensions.Count == 0).ToArray();
            var metrics = reportFacts.Select(fact => fact.Metric)
                .Distinct(StringComparer.Ordinal).ToArray();
            if (metrics.Length == 0 ||
                metrics.Length + 1 > WorkbookDraftWriter.MaxDraftRows ||
                metrics.Any(metric => !headers.ContainsKey(metric)) ||
                reportFacts.Length != metrics.Length * allPeriods.Length ||
                reportFacts.Any(fact => fact.Status !=
                    AnalysisContract.Verified ||
                    fact.SnapshotId != snapshot.SnapshotId))
                throw new InvalidOperationException(
                    "ANALYSIS_WORKBOOK_FACTS_UNSUPPORTED");
            var binding = new AnalysisTableBinding
            {
                TableId = table.TableId,
                PeriodHeader = "Period",
                Metrics = metrics.Select(metric =>
                {
                    var fact = reportFacts.First(item =>
                        item.Metric == metric);
                    return new AnalysisMetricColumnBinding
                    {
                        Header = metric, Metric = metric,
                        Unit = fact.Unit, Currency = fact.Currency
                    };
                }).ToList()
            };
            if (reportFacts.All(fact => fact.Kind ==
                    AnalysisContract.SourceObserved) &&
                AnalysisTableArtifactBuilder.Build(snapshot,
                    binding).AnalysisId != artifact.AnalysisId)
                throw new InvalidOperationException(
                    "ANALYSIS_WORKBOOK_FACTS_UNSUPPORTED");
            selection = selection ?? AnalysisRequestPlan.Resolve(artifact,
                null);
            selection.Validate(artifact);
            var periods = new[] { selection.ComparePeriod,
                selection.FocusPeriod }.Select(period => allPeriods.SingleOrDefault(
                    cell => cell.Value == period)).ToArray();
            if (periods.Any(cell => cell == null) ||
                periods.Length + 1 > WorkbookDraftWriter.MaxDraftColumns)
                throw new InvalidOperationException(
                    "ANALYSIS_WORKBOOK_PERIOD_UNSUPPORTED");
            var sourcePeriod = SourceColumnRange(table, periodColumn);
            var sheet = "'" + table.Name.Replace("'", "''") + "'!";
            var duplicateRows = AnalysisTableArtifactBuilder
                .DuplicateIdentityRows(table);
            int identityColumn;
            var hasIdentity = headers.TryGetValue("RowID", out identityColumn);
            if (duplicateRows.Count > 0 && table.Cells.Where(cell =>
                    cell.Row > 0 && cell.Column == identityColumn)
                .GroupBy(cell => cell.Value,
                    StringComparer.OrdinalIgnoreCase)
                .Any(group => group.Select(cell => cell.Value)
                    .Distinct(StringComparer.Ordinal).Count() > 1))
                throw new InvalidOperationException(
                    "ANALYSIS_WORKBOOK_ROW_ID_CASE_AMBIGUOUS");
            var identityRange = hasIdentity
                ? SourceColumnRange(table, identityColumn) : null;
            var firstIdentity = hasIdentity ? table.Cells.Single(cell =>
                cell.Row == 1 && cell.Column == identityColumn).Reference : null;
            var firstIdentityMatch = CellAddress.Match(firstIdentity ?? "");
            var firstIdentityAddress = hasIdentity
                ? "$" + firstIdentityMatch.Groups[1].Value + "$" +
                    firstIdentityMatch.Groups[2].Value : null;
            var rows = new List<AnalysisPlanRow>();
            rows.Add(Row(new[] { Label("Metric") }.Concat(
                periods.Select(cell => Label(cell.Value)))));
            foreach (var metric in selection.ReportMetrics)
            {
                var metricColumn = SourceColumnRange(table,
                    headers[metric]);
                var cells = new List<AnalysisPlanCell> { Label(metric) };
                for (var index = 0; index < periods.Length; index++)
                {
                    var period = periods[index].Value;
                    var fact = reportFacts.Single(item =>
                        item.Metric == metric && item.Period == period);
                    var reportColumn = ColumnName(index + 2);
                    var formula = "=SUMIF(" + sheet + sourcePeriod + "," +
                        reportColumn + "$3," + sheet + metricColumn + ")";
                    if (duplicateRows.Count > 0)
                        formula = "=SUMPRODUCT((" + sheet + sourcePeriod +
                            "=" + reportColumn + "$3)*(MATCH(" + sheet +
                            identityRange + "," + sheet + identityRange +
                            ",0)=ROW(" + sheet + identityRange +
                            ")-ROW(" + sheet + firstIdentityAddress +
                            ")+1)*IFERROR(1*" + sheet + metricColumn +
                            ",0))";
                    if (formula.Length > 500 ||
                        (hasIdentity && !firstIdentityMatch.Success))
                        throw new InvalidOperationException(
                            "ANALYSIS_WORKBOOK_FORMULA_UNSUPPORTED");
                    cells.Add(new AnalysisPlanCell
                    {
                        Formula = formula,
                        ExpectedFactId = fact.FactId
                    });
                }
                rows.Add(Row(cells));
            }
            var knownSubtotalNotes = periods.Select(period => {
                var notes = selection.ReportMetrics.Where(metric =>
                    AnalysisContract.IsKnownSubtotal(artifact, metric,
                        period.Value)).Select(metric => metric +
                    ": known subtotal; blank source values excluded")
                    .ToList();
                var repeated = duplicateRows.Count(row =>
                    sourcePeriods.Single(cell => cell.Row == row).Value ==
                        period.Value);
                if (repeated > 0)
                    notes.Add(repeated.ToString(CultureInfo.InvariantCulture) +
                        " repeated RowID(s) counted once");
                return string.Join("; ", notes);
            }).ToArray();
            if (knownSubtotalNotes.Any(note => note.Length != 0))
            {
                if (rows.Count + 1 > WorkbookDraftWriter.MaxDraftRows)
                    throw new InvalidOperationException(
                        "ANALYSIS_WORKBOOK_ROWS_UNSUPPORTED");
                rows.Add(Row(new[] { Label("Data quality") }.Concat(
                    knownSubtotalNotes.Select(note => Label(note.Length == 0 ?
                        "No blank inputs in selected metrics" : note)))));
            }
            return rows;
        }

        private static string SourceColumnRange(TableDataset table, int column)
        {
            var first = table.Cells.Single(cell =>
                cell.Row == 1 && cell.Column == column);
            var last = table.Cells.Single(cell =>
                cell.Row == table.Rows - 1 && cell.Column == column);
            var start = CellAddress.Match(first.Reference ?? string.Empty);
            var end = CellAddress.Match(last.Reference ?? string.Empty);
            if (!start.Success || !end.Success ||
                start.Groups[1].Value != end.Groups[1].Value ||
                int.Parse(end.Groups[2].Value, CultureInfo.InvariantCulture) -
                int.Parse(start.Groups[2].Value, CultureInfo.InvariantCulture)
                    != table.Rows - 2)
                throw new InvalidOperationException(
                    "ANALYSIS_WORKBOOK_RANGE_UNSUPPORTED");
            var letters = start.Groups[1].Value;
            return "$" + letters + "$" + start.Groups[2].Value +
                ":$" + letters + "$" + end.Groups[2].Value;
        }

        private static string ColumnName(int column)
        {
            var name = string.Empty;
            while (column > 0)
            {
                column--;
                name = (char)('A' + column % 26) + name;
                column /= 26;
            }
            return name;
        }

        private static AnalysisPlanCell Label(string text)
        {
            return new AnalysisPlanCell { Text = text };
        }

        private static AnalysisPlanRow Row(IEnumerable<AnalysisPlanCell> cells)
        {
            return new AnalysisPlanRow { Cells = cells.ToList() };
        }
    }
}
