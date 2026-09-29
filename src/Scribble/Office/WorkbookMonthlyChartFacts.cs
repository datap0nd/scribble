using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Threading;

namespace Scribble.Office
{
    // Read-only monthly chart binding. Values come from verified typed
    // source aggregates; a chart's old cache is never factual authority.
    public static class WorkbookMonthlyChartFacts
    {
        public sealed class BoundSeries
        {
            public string SourceSha256 { get; set; }
            public string SnapshotId { get; set; }
            public string[] Categories { get; set; }
            public string[] Names { get; set; }
            public decimal[][] Values { get; set; }
            public string[] Limitations { get; set; }
            public VerifiedFact[] Facts { get; set; }
        }

        // Chart series are selected by their visible names, then bound to
        // unique workbook headers. The typed aggregate recomputes source
        // formulas and duplicate identities before any chart write.
        public static BoundSeries ReadBoundSeries(string path,
            IReadOnlyList<string> seriesNames,
            CancellationToken cancellationToken)
        {
            if (seriesNames == null || seriesNames.Count == 0 ||
                seriesNames.Count > 6 ||
                seriesNames.Any(string.IsNullOrWhiteSpace) ||
                seriesNames.Select(NormalizeName).Distinct(
                    StringComparer.Ordinal).Count() != seriesNames.Count)
                throw new InvalidOperationException(
                    "MONTHLY_CHART_SERIES_INVALID");
            var before = Hash(path);
            var snapshot = OpenXmlWorkbookSnapshotReader.Capture(path,
                Path.GetFullPath(path), before, cancellationToken);
            var matches = new List<Tuple<TableDataset, string[]>>();
            foreach (var table in snapshot.Tables)
            {
                var headers = table.Cells.Where(cell => cell.Row == 0 &&
                    cell.Status == AnalysisContract.Verified &&
                    cell.ValueType == AnalysisContract.TextValue &&
                    !string.IsNullOrWhiteSpace(cell.Value))
                    .Select(cell => cell.Value.Trim()).ToArray();
                if (headers.Count(name => name == "Period") != 1 ||
                    headers.Count(name => name == "RowID") != 1)
                    continue;
                var bound = new List<string>();
                foreach (var name in seriesNames)
                {
                    var candidates = headers.Where(header =>
                        NormalizeName(header) == NormalizeName(name))
                        .ToArray();
                    if (candidates.Length != 1)
                    { bound.Clear(); break; }
                    bound.Add(candidates[0]);
                }
                if (bound.Count == seriesNames.Count)
                    matches.Add(Tuple.Create(table, bound.ToArray()));
            }
            if (matches.Count != 1)
                throw new InvalidOperationException(
                    matches.Count == 0 ? "MONTHLY_CHART_SOURCE_MISSING" :
                        "MONTHLY_CHART_SOURCE_AMBIGUOUS");
            var selected = matches[0];
            var scoped = AnalysisContract.CreateSnapshot(
                snapshot.SourceInstanceId, snapshot.SourceType,
                snapshot.CaptureRevision, "complete_range",
                snapshot.CalculationState, snapshot.Locators,
                new[] { selected.Item1 });
            var binding = new AnalysisTableBinding {
                TableId = selected.Item1.TableId,
                PeriodHeader = "Period",
                Metrics = selected.Item2.Select(header =>
                    new AnalysisMetricColumnBinding {
                        Header = header, Metric = header,
                        Unit = header.EndsWith("EUR",
                            StringComparison.OrdinalIgnoreCase)
                            ? "currency" : string.Empty,
                        Currency = header.EndsWith("EUR",
                            StringComparison.OrdinalIgnoreCase)
                            ? "EUR" : string.Empty
                    }).ToList()
            };
            var facts = AnalysisTableArtifactBuilder.BuildGrouped(scoped,
                binding);
            var categories = facts.Facts.Where(fact =>
                fact.Dimensions.Count == 0).Select(fact => fact.Period)
                .Distinct(StringComparer.Ordinal).OrderBy(period => period,
                    StringComparer.Ordinal).ToArray();
            if (categories.Length < 2 ||
                categories.Any(category => !Regex.IsMatch(category,
                    @"^\d{4}-(?:0[1-9]|1[0-2])$")))
                throw new InvalidOperationException(
                    "MONTHLY_CHART_PERIODS_INVALID");
            var values = new List<decimal[]>();
            foreach (var header in selected.Item2)
            {
                var series = new List<decimal>();
                foreach (var category in categories)
                {
                    var matching = facts.Facts.Where(fact =>
                        fact.Metric == header && fact.Period == category &&
                        fact.Dimensions.Count == 0).ToArray();
                    if (matching.Length != 1)
                        throw new InvalidOperationException(
                            "MONTHLY_CHART_FACT_AMBIGUOUS");
                    series.Add(AnalysisContract.Decimal(matching[0]));
                }
                values.Add(series.ToArray());
            }
            if (before != Hash(path))
                throw new InvalidOperationException(
                    "MONTHLY_CHART_SOURCE_CHANGED");
            return new BoundSeries {
                SourceSha256 = before,
                SnapshotId = scoped.SnapshotId,
                Categories = categories,
                Names = seriesNames.ToArray(),
                Values = values.ToArray(),
                Limitations = facts.UnresolvedConflicts.ToArray(),
                Facts = facts.Facts.ToArray()
            };
        }

        private static string NormalizeName(string value)
        {
            return Regex.Replace(value ?? string.Empty, @"[^a-z0-9]",
                string.Empty, RegexOptions.IgnoreCase).ToLowerInvariant();
        }

        private static string Hash(string path)
        {
            using (var stream = File.OpenRead(path))
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(stream))
                    .Replace("-", "").ToLowerInvariant();
        }
    }
}
