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
            AnalysisContract.Serialize(artifact);
            if (artifact.Snapshots.Count != 1 ||
                artifact.Snapshots[0].Tables.Count != 1 ||
                artifact.Facts.Count == 0 ||
                artifact.Calculations.Count != 0 ||
                artifact.Assumptions.Count != 0 ||
                artifact.UnresolvedConflicts.Count != 0)
                throw new InvalidOperationException(
                    "ANALYSIS_WORKBOOK_SOURCE_UNSUPPORTED");
            var snapshot = artifact.Snapshots[0];
            var table = snapshot.Tables[0];
            if (snapshot.Coverage != "complete_range" ||
                table.Rows < 2 || table.Columns < 2 ||
                table.Rows > WorkbookDraftWriter.MaxDraftColumns ||
                table.Columns > WorkbookDraftWriter.MaxDraftColumns)
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
            var periods = table.Cells.Where(cell =>
                    cell.Row > 0 && cell.Column == periodColumn)
                .OrderBy(cell => cell.Row).ToArray();
            if (periods.Length != table.Rows - 1 ||
                periods.Select(cell => cell.Value)
                    .Distinct(StringComparer.Ordinal).Count() != periods.Length)
                throw new InvalidOperationException(
                    "ANALYSIS_WORKBOOK_PERIOD_UNSUPPORTED");
            var metrics = artifact.Facts.Select(fact => fact.Metric)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(metric => headers.ContainsKey(metric)
                    ? headers[metric] : int.MaxValue).ToArray();
            if (metrics.Length + 1 > WorkbookDraftWriter.MaxDraftRows ||
                metrics.Any(metric => !headers.ContainsKey(metric)) ||
                artifact.Facts.Any(fact => fact.Kind !=
                    AnalysisContract.SourceObserved ||
                    fact.Status != AnalysisContract.Verified ||
                    fact.Dimensions.Count != 0))
                throw new InvalidOperationException(
                    "ANALYSIS_WORKBOOK_FACTS_UNSUPPORTED");
            var binding = new AnalysisTableBinding
            {
                TableId = table.TableId,
                PeriodHeader = "Period",
                Metrics = metrics.Select(metric =>
                {
                    var fact = artifact.Facts.First(item =>
                        item.Metric == metric);
                    return new AnalysisMetricColumnBinding
                    {
                        Header = metric, Metric = metric,
                        Unit = fact.Unit, Currency = fact.Currency
                    };
                }).ToList()
            };
            var rebound = AnalysisTableArtifactBuilder.Build(snapshot,
                binding);
            if (rebound.AnalysisId != artifact.AnalysisId)
                throw new InvalidOperationException(
                    "ANALYSIS_WORKBOOK_FACTS_UNSUPPORTED");
            var sourcePeriod = SourceColumnRange(table, periodColumn);
            var sheet = "'" + table.Name.Replace("'", "''") + "'!";
            var rows = new List<AnalysisPlanRow>();
            rows.Add(Row(new[] { Label("Metric") }.Concat(
                periods.Select(cell => Label(cell.Value)))));
            foreach (var metric in metrics)
            {
                var metricColumn = SourceColumnRange(table,
                    headers[metric]);
                var cells = new List<AnalysisPlanCell> { Label(metric) };
                for (var index = 0; index < periods.Length; index++)
                {
                    var period = periods[index].Value;
                    var fact = artifact.Facts.Single(item =>
                        item.Metric == metric && item.Period == period);
                    var reportColumn = ColumnName(index + 2);
                    cells.Add(new AnalysisPlanCell
                    {
                        Formula = "=SUMIF(" + sheet + sourcePeriod + "," +
                            reportColumn + "$3," + sheet + metricColumn + ")",
                        ExpectedFactId = fact.FactId
                    });
                }
                rows.Add(Row(cells));
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
