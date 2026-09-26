using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Scribble.Office
{
    public sealed class AnalysisMetricColumnBinding
    {
        public string Header { get; set; }
        public string Metric { get; set; }
        public string Unit { get; set; }
        public string Currency { get; set; }
    }

    public sealed class AnalysisTableBinding
    {
        public string TableId { get; set; }
        public string PeriodHeader { get; set; }
        public List<string> DimensionHeaders { get; set; } =
            new List<string>();
        public List<AnalysisMetricColumnBinding> Metrics { get; set; } =
            new List<AnalysisMetricColumnBinding>();
    }

    // Bind a complete typed worksheet table to semantic facts. Column roles
    // are explicit; the host reads values and issues IDs. A model cannot
    // register a number or a stale formula cache as verified data.
    public static class AnalysisTableArtifactBuilder
    {
        public const int MaxMappedFacts = 20000;

        // A grouped read names the grouping columns and additive metrics.
        // Recompute every aggregate from the typed source cells, retaining
        // the exact contributing cell locators in each host-issued fact.
        public static AnalysisArtifact BuildGrouped(SourceSnapshot snapshot,
            AnalysisTableBinding binding, string filterColumn = null,
            string filterEquals = null)
        {
            if (snapshot == null || snapshot.Coverage != "complete_range" ||
                snapshot.Tables == null || snapshot.Tables.Count != 1 ||
                binding == null || binding.PeriodHeader != "Period" ||
                binding.Metrics == null || binding.Metrics.Count == 0 ||
                binding.DimensionHeaders == null)
                throw new InvalidOperationException(
                    "ANALYSIS_GROUP_BINDING_INVALID");
            var table = snapshot.Tables[0];
            if (table.TableId != binding.TableId || table.Rows < 2 ||
                (long)table.Rows * binding.Metrics.Count > MaxMappedFacts ||
                binding.Metrics.Any(metric => metric == null))
                throw new InvalidOperationException(
                    "ANALYSIS_GROUP_BINDING_INVALID");
            var cells = table.Cells.ToDictionary(cell =>
                cell.Row + ":" + cell.Column, StringComparer.Ordinal);
            var headers = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var column = 0; column < table.Columns; column++)
            {
                DatasetCell header;
                if (!cells.TryGetValue("0:" + column, out header) ||
                    string.IsNullOrWhiteSpace(header.Value)) continue;
                if (header.Status != AnalysisContract.Verified ||
                    header.ValueType != AnalysisContract.TextValue ||
                    headers.ContainsKey(header.Value.Trim()))
                    throw new InvalidOperationException(
                        "ANALYSIS_TABLE_HEADER_AMBIGUOUS");
                headers.Add(header.Value.Trim(), column);
            }
            var names = new[] { binding.PeriodHeader }
                .Concat(binding.DimensionHeaders)
                .Concat(binding.Metrics.Select(metric => metric.Header))
                .Concat(string.IsNullOrEmpty(filterColumn)
                    ? new string[0] : new[] { filterColumn }).ToArray();
            if (names.Any(string.IsNullOrWhiteSpace) ||
                names.Take(names.Length - (string.IsNullOrEmpty(filterColumn)
                    ? 0 : 1)).Distinct(StringComparer.Ordinal).Count() !=
                    names.Length - (string.IsNullOrEmpty(filterColumn)
                        ? 0 : 1) || names.Any(name => !headers.ContainsKey(name)) ||
                binding.Metrics.Any(metric => metric == null ||
                    metric.Metric != metric.Header ||
                    (!string.IsNullOrEmpty(metric.Currency) &&
                     (metric.Unit != "currency" ||
                      !metric.Header.EndsWith(metric.Currency,
                          StringComparison.Ordinal)))))
                throw new InvalidOperationException(
                    "ANALYSIS_GROUP_BINDING_INVALID");
            var groups = new Dictionary<string, GroupedValue>(
                StringComparer.Ordinal);
            for (var row = 1; row < table.Rows; row++)
            {
                if (!string.IsNullOrEmpty(filterColumn) &&
                    Required(cells, row, headers[filterColumn]).Value !=
                    filterEquals) continue;
                var period = Required(cells, row,
                    headers[binding.PeriodHeader]);
                if (period.ValueType != AnalysisContract.TextValue ||
                    !Regex.IsMatch(period.Value,
                        @"^\d{4}-(?:0[1-9]|1[0-2])$"))
                    throw new InvalidOperationException(
                        "ANALYSIS_TABLE_PERIOD_UNVERIFIED");
                var dimensions = new Dictionary<string, string>(
                    StringComparer.Ordinal);
                foreach (var name in binding.DimensionHeaders)
                {
                    var dimension = Required(cells, row, headers[name]);
                    if (dimension.ValueType != AnalysisContract.TextValue &&
                        dimension.ValueType != AnalysisContract.DateValue)
                        throw new InvalidOperationException(
                            "ANALYSIS_TABLE_DIMENSION_UNVERIFIED");
                    dimensions.Add(name, dimension.Value);
                }
                foreach (var metric in binding.Metrics)
                {
                    var cell = Required(cells, row,
                        headers[metric.Header]);
                    if ((cell.ValueType != AnalysisContract.DecimalValue &&
                         cell.ValueType != AnalysisContract.IntegerValue) ||
                        !string.IsNullOrEmpty(cell.Formula))
                        throw new InvalidOperationException(
                            "ANALYSIS_TABLE_VALUE_UNVERIFIED");
                    decimal value;
                    if (!decimal.TryParse(cell.Value, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out value))
                        throw new InvalidOperationException(
                            "ANALYSIS_TABLE_VALUE_UNVERIFIED");
                    var key = metric.Header + "\0" + period.Value + "\0" +
                        string.Join("\0", binding.DimensionHeaders.Select(
                            name => dimensions[name]));
                    GroupedValue group;
                    if (!groups.TryGetValue(key, out group))
                    {
                        group = new GroupedValue { Metric = metric,
                            Period = period.Value,
                            Dimensions = dimensions };
                        groups.Add(key, group);
                    }
                    group.Sum = checked(group.Sum + value);
                    group.Locators.Add(new SourceLocator {
                        Kind = "excel_cell",
                        SourceInstanceId = snapshot.SourceInstanceId,
                        WorksheetIdentity = table.Name,
                        Cell = cell.Reference });
                }
            }
            if (groups.Count == 0)
                throw new InvalidOperationException(
                    "ANALYSIS_GROUP_EMPTY");
            var facts = groups.Values.Select(group =>
            {
                var fact = AnalysisContract.CreateObservedFact(
                    snapshot.SnapshotId, group.Metric.Header,
                    AnalysisContract.DecimalValue,
                    AnalysisContract.Format(group.Sum),
                    AnalysisContract.Format(group.Sum),
                    group.Metric.Unit, group.Metric.Currency,
                    group.Period, group.Dimensions, group.Locators,
                    AnalysisContract.Verified);
                fact.Kind = AnalysisContract.DerivedCalculation;
                fact.FactId = AnalysisContract.ExpectedFactId(fact);
                return fact;
            }).ToArray();
            return AnalysisContract.CreateArtifact(new[] { snapshot }, facts,
                new AnalysisCalculation[0], new string[0], new string[0]);
        }

        private sealed class GroupedValue
        {
            public AnalysisMetricColumnBinding Metric;
            public string Period;
            public Dictionary<string, string> Dimensions;
            public decimal Sum;
            public List<SourceLocator> Locators = new List<SourceLocator>();
        }

        public static AnalysisArtifact Build(SourceSnapshot snapshot,
            AnalysisTableBinding binding)
        {
            AnalysisContract.CreateArtifact(new[] { snapshot },
                new VerifiedFact[0], new AnalysisCalculation[0],
                new string[0], new string[0]);
            if (!snapshot.SourceType.StartsWith("excel", StringComparison.Ordinal) ||
                (snapshot.Coverage != "complete_range" &&
                 snapshot.Coverage !=
                    "package_cached_values_formulas_formats") ||
                snapshot.Tables == null)
                throw new InvalidOperationException(
                    "ANALYSIS_TABLE_SOURCE_UNSUPPORTED");
            if (binding == null || string.IsNullOrWhiteSpace(binding.TableId) ||
                binding.PeriodHeader != "Period" ||
                binding.Metrics == null || binding.Metrics.Count == 0 ||
                binding.DimensionHeaders == null ||
                binding.DimensionHeaders.Any(string.IsNullOrWhiteSpace) ||
                binding.Metrics.Any(metric => metric == null ||
                    string.IsNullOrWhiteSpace(metric.Header) ||
                    string.IsNullOrWhiteSpace(metric.Metric) ||
                    metric.Metric != metric.Header ||
                    (string.IsNullOrWhiteSpace(metric.Currency)
                        ? !string.IsNullOrWhiteSpace(metric.Unit)
                        : metric.Unit != "currency" ||
                            !metric.Header.EndsWith(metric.Currency,
                                StringComparison.Ordinal) ||
                            metric.Currency.Length != 3)))
                throw new InvalidOperationException("ANALYSIS_TABLE_BINDING_INVALID");
            var table = snapshot.Tables.SingleOrDefault(item =>
                item.TableId == binding.TableId);
            if (table == null || table.Rows < 2 || table.Columns < 2 ||
                (long)(table.Rows - 1) * binding.Metrics.Count >
                    MaxMappedFacts)
                throw new InvalidOperationException(
                    "ANALYSIS_TABLE_BINDING_UNSUPPORTED");
            var cells = table.Cells.ToDictionary(item =>
                item.Row + ":" + item.Column, StringComparer.Ordinal);
            var headers = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var column = 0; column < table.Columns; column++)
            {
                DatasetCell cell;
                if (!cells.TryGetValue("0:" + column, out cell) ||
                    string.IsNullOrWhiteSpace(cell.Value)) continue;
                if (cell.ValueType != AnalysisContract.TextValue ||
                    cell.Status != AnalysisContract.Verified ||
                    headers.ContainsKey(cell.Value.Trim()))
                    throw new InvalidOperationException(
                        "ANALYSIS_TABLE_HEADER_AMBIGUOUS");
                headers.Add(cell.Value.Trim(), column);
            }
            var bound = new[] { binding.PeriodHeader }
                .Concat(binding.DimensionHeaders)
                .Concat(binding.Metrics.Select(metric => metric.Header))
                .ToArray();
            if (bound.Distinct(StringComparer.Ordinal).Count() !=
                    bound.Length || bound.Any(header =>
                    !headers.ContainsKey(header)))
                throw new InvalidOperationException(
                    "ANALYSIS_TABLE_HEADER_AMBIGUOUS");
            var facts = new List<VerifiedFact>();
            var keys = new HashSet<string>(StringComparer.Ordinal);
            for (var row = 1; row < table.Rows; row++)
            {
                var period = Required(cells, row,
                    headers[binding.PeriodHeader]);
                if (period.ValueType != AnalysisContract.TextValue ||
                    !Regex.IsMatch(period.Value,
                        @"^\d{4}-(?:0[1-9]|1[0-2])$"))
                    throw new InvalidOperationException(
                        "ANALYSIS_TABLE_PERIOD_UNVERIFIED");
                var dimensions = new Dictionary<string, string>(
                    StringComparer.Ordinal);
                foreach (var header in binding.DimensionHeaders)
                {
                    var dimension = Required(cells, row, headers[header]);
                    if (dimension.ValueType != AnalysisContract.TextValue &&
                        dimension.ValueType != AnalysisContract.DateValue)
                        throw new InvalidOperationException(
                            "ANALYSIS_TABLE_DIMENSION_UNVERIFIED");
                    dimensions.Add(header, dimension.Value);
                }
                foreach (var metric in binding.Metrics)
                {
                    var cell = Required(cells, row,
                        headers[metric.Header]);
                    if ((cell.ValueType != AnalysisContract.DecimalValue &&
                         cell.ValueType != AnalysisContract.IntegerValue) ||
                        string.IsNullOrWhiteSpace(cell.Formula) == false)
                        throw new InvalidOperationException(
                            "ANALYSIS_TABLE_VALUE_UNVERIFIED");
                    var key = metric.Metric + "|" + period.Value + "|" +
                        string.Join("|", dimensions.OrderBy(item => item.Key,
                            StringComparer.Ordinal).Select(item =>
                                item.Key + "=" + item.Value));
                    if (!keys.Add(key))
                        throw new InvalidOperationException(
                            "ANALYSIS_TABLE_FACT_DUPLICATE: " + key);
                    facts.Add(AnalysisContract.CreateObservedFact(
                        snapshot.SnapshotId, metric.Metric,
                        cell.ValueType, cell.Value, cell.DisplayText,
                        metric.Unit, metric.Currency, period.Value,
                        dimensions, new[] { new SourceLocator
                        {
                            Kind = "excel_cell",
                            SourceInstanceId = snapshot.SourceInstanceId,
                            WorksheetIdentity = table.Name,
                            Cell = cell.Reference
                        } }, AnalysisContract.Verified));
                }
            }
            return AnalysisContract.CreateArtifact(new[] { snapshot },
                facts, new AnalysisCalculation[0], new string[0],
                new string[0]);
        }

        private static DatasetCell Required(
            IDictionary<string, DatasetCell> cells, int row, int column)
        {
            DatasetCell cell;
            if (!cells.TryGetValue(row + ":" + column, out cell) ||
                cell.Status != AnalysisContract.Verified ||
                string.IsNullOrWhiteSpace(cell.Reference) ||
                string.IsNullOrWhiteSpace(cell.Value))
                throw new InvalidOperationException(
                    "ANALYSIS_TABLE_VALUE_UNVERIFIED");
            return cell;
        }
    }
}
