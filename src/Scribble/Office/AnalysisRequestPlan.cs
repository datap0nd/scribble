using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Scribble.Office
{
    // The user's comparison and chart intent is resolved before any native
    // write. Model-supplied selections can confirm it, but cannot redirect it.
    public sealed class AnalysisRequestPlan
    {
        private static readonly string[] Months = {
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November",
            "December" };
        private const string PeriodToken =
            @"(?:20[0-9]{2}-(?:0[1-9]|1[0-2])|(?:January|February|March|April|May|June|July|August|September|October|November|December)(?:\s+20[0-9]{2})?)";
        private static readonly Regex Comparison = new Regex(
            @"\b(?<first>" + PeriodToken + @")(?:\s*(?:-|–|—|/|to|vs\.?|versus|against|and|with)\s*)+(?<second>" +
            PeriodToken + @")\b", RegexOptions.IgnoreCase |
            RegexOptions.CultureInvariant);
        private static readonly Regex PeriodMention = new Regex(
            @"\b" + PeriodToken + @"\b", RegexOptions.IgnoreCase |
            RegexOptions.CultureInvariant);
        private static readonly Regex CanonicalPeriod = new Regex(
            @"^20[0-9]{2}-(?:0[1-9]|1[0-2])$",
            RegexOptions.CultureInvariant);

        public string ComparePeriod { get; set; }
        public string FocusPeriod { get; set; }
        public List<string> ChartSeries { get; set; } = new List<string>();

        public static AnalysisRequestPlan Resolve(AnalysisArtifact artifact,
            string objective, IDictionary<string, object> supplied = null)
        {
            AnalysisContract.Serialize(artifact);
            if (artifact.Snapshots.Count != 1 ||
                artifact.Snapshots[0].Tables.Count != 1)
                throw new InvalidOperationException(
                    "ANALYSIS_REQUEST_SOURCE_UNSUPPORTED");
            var periods = artifact.Facts.Where(fact =>
                    fact.Dimensions.Count == 0 &&
                    CanonicalPeriod.IsMatch(fact.Period ?? string.Empty))
                .Select(fact => fact.Period).Distinct(StringComparer.Ordinal)
                .OrderBy(period => period, StringComparer.Ordinal).ToArray();
            if (periods.Length < 2)
                throw new InvalidOperationException(
                    "ANALYSIS_REQUEST_PERIODS_MISSING");
            var pair = Comparison.Match(objective ?? string.Empty);
            if (!pair.Success && PeriodMention.IsMatch(objective ?? string.Empty))
                throw new InvalidOperationException(
                    "ANALYSIS_REQUEST_COMPARISON_UNBOUND");
            var requested = pair.Success ? new[] {
                ResolvePeriod(pair.Groups["first"].Value, periods),
                ResolvePeriod(pair.Groups["second"].Value, periods) }
                .OrderBy(period => period, StringComparer.Ordinal).ToArray()
                : periods.Skip(periods.Length - 2).ToArray();
            var plan = new AnalysisRequestPlan {
                ComparePeriod = requested[0],
                FocusPeriod = requested[1]
            };
            var metrics = AvailableMetrics(artifact, plan.ComparePeriod,
                plan.FocusPeriod);
            if (metrics.Length == 0)
                throw new InvalidOperationException(
                    "ANALYSIS_REQUEST_METRICS_MISSING");
            var chartText = string.Join(" ", Regex.Matches(objective ?? "",
                @"\bchart\b[^.!?]*", RegexOptions.IgnoreCase |
                RegexOptions.CultureInvariant).Cast<Match>()
                .Select(match => match.Value));
            var both = Regex.IsMatch(chartText,
                @"\bboth\s+(?:metric|metrics|series)\b|\bprimary\s+and\s+secondary\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            var primaryOnly = Regex.IsMatch(chartText,
                @"\bonly\s+(?:the\s+)?primary\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (both && primaryOnly)
                throw new InvalidOperationException(
                    "ANALYSIS_REQUEST_CHART_AMBIGUOUS");
            var named = metrics.Where(metric =>
                chartText.IndexOf(metric, StringComparison.OrdinalIgnoreCase)
                    >= 0 ||
                chartText.IndexOf(ReadableMetric(metric),
                    StringComparison.OrdinalIgnoreCase) >= 0).ToArray();
            plan.ChartSeries = both ? metrics.Take(2).ToList() :
                primaryOnly ? metrics.Take(1).ToList() :
                named.Length > 0 ? named.ToList() :
                metrics.Take(1).ToList();
            if (both && metrics.Length != 2)
                throw new InvalidOperationException(
                    "ANALYSIS_REQUEST_CHART_SERIES_UNBOUND");
            ValidateSupplied(plan, supplied, metrics);
            plan.Validate(artifact);
            return plan;
        }

        public void Validate(AnalysisArtifact artifact)
        {
            var periods = artifact.Facts.Where(fact =>
                    fact.Dimensions.Count == 0 &&
                    fact.Status == AnalysisContract.Verified)
                .Select(fact => fact.Period).Distinct(StringComparer.Ordinal)
                .ToArray();
            if (!periods.Contains(ComparePeriod, StringComparer.Ordinal) ||
                !periods.Contains(FocusPeriod, StringComparer.Ordinal) ||
                string.CompareOrdinal(ComparePeriod, FocusPeriod) >= 0)
                throw new InvalidOperationException(
                    "ANALYSIS_REQUEST_PERIOD_UNBOUND");
            var available = AvailableMetrics(artifact, ComparePeriod,
                FocusPeriod);
            if (ChartSeries == null || ChartSeries.Count == 0 ||
                ChartSeries.Count > PresentationDraftWriter.MaxChartSeries ||
                ChartSeries.Distinct(StringComparer.Ordinal).Count() !=
                    ChartSeries.Count ||
                ChartSeries.Any(metric => !available.Contains(metric,
                    StringComparer.Ordinal)))
                throw new InvalidOperationException(
                    "ANALYSIS_REQUEST_CHART_SERIES_UNBOUND");
        }

        private static string[] AvailableMetrics(AnalysisArtifact artifact,
            string compare, string focus)
        {
            var facts = artifact.Facts.Where(fact =>
                fact.Dimensions.Count == 0 &&
                fact.Status == AnalysisContract.Verified &&
                (fact.ValueType == AnalysisContract.DecimalValue ||
                 fact.ValueType == AnalysisContract.IntegerValue)).ToArray();
            var headers = artifact.Snapshots[0].Tables[0].Cells.Where(cell =>
                cell.Row == 0).ToArray();
            return facts.Where(fact => fact.Period == focus &&
                    facts.Any(other => other.Period == compare &&
                        other.Metric == fact.Metric &&
                        other.Unit == fact.Unit &&
                        other.Currency == fact.Currency))
                .Select(fact => fact.Metric).Distinct(StringComparer.Ordinal)
                .OrderBy(metric => headers.Where(cell => cell.Value == metric)
                    .Select(cell => cell.Column).DefaultIfEmpty(int.MaxValue)
                    .First()).ToArray();
        }

        private static string ResolvePeriod(string token, string[] periods)
        {
            token = token.Trim();
            if (CanonicalPeriod.IsMatch(token)) return token;
            var parts = token.Split(new[] { ' ' },
                StringSplitOptions.RemoveEmptyEntries);
            var month = Array.FindIndex(Months, value =>
                string.Equals(value, parts[0],
                    StringComparison.OrdinalIgnoreCase)) + 1;
            if (month == 0)
                throw new InvalidOperationException(
                    "ANALYSIS_REQUEST_PERIOD_UNBOUND");
            var suffix = "-" + month.ToString("00",
                CultureInfo.InvariantCulture);
            var matches = periods.Where(period => period.EndsWith(suffix,
                StringComparison.Ordinal) &&
                (parts.Length == 1 || period.StartsWith(parts[1] + "-",
                    StringComparison.Ordinal))).ToArray();
            if (matches.Length != 1)
                throw new InvalidOperationException(matches.Length == 0 ?
                    "ANALYSIS_REQUEST_PERIOD_UNBOUND" :
                    "ANALYSIS_REQUEST_PERIOD_AMBIGUOUS");
            return matches[0];
        }

        private static void ValidateSupplied(AnalysisRequestPlan plan,
            IDictionary<string, object> supplied, string[] available)
        {
            if (supplied == null) return;
            var compare = StringOption(supplied, "ComparePeriod",
                "compare_period");
            var focus = StringOption(supplied, "FocusPeriod",
                "focus_period");
            if ((compare != null && compare != plan.ComparePeriod) ||
                (focus != null && focus != plan.FocusPeriod))
                throw new InvalidOperationException(
                    "ANALYSIS_REQUEST_PERIOD_MISMATCH");
            object raw;
            if (!supplied.TryGetValue("ChartSeries", out raw) &&
                !supplied.TryGetValue("chart_series", out raw)) return;
            var items = raw as IEnumerable;
            if (items == null || raw is string)
                throw new InvalidOperationException(
                    "ANALYSIS_REQUEST_CHART_SERIES_UNBOUND");
            var requested = new List<string>();
            foreach (var item in items)
            {
                var name = item as string;
                var metric = available.FirstOrDefault(value =>
                    string.Equals(value, name,
                        StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(ReadableMetric(value), name,
                        StringComparison.OrdinalIgnoreCase));
                if (metric == null)
                    throw new InvalidOperationException(
                        "ANALYSIS_REQUEST_CHART_SERIES_UNBOUND");
                requested.Add(metric);
            }
            if (!requested.SequenceEqual(plan.ChartSeries,
                StringComparer.Ordinal))
                throw new InvalidOperationException(
                    "ANALYSIS_REQUEST_CHART_SERIES_MISMATCH");
        }

        private static string StringOption(IDictionary<string, object> values,
            string primary, string alternate)
        {
            object raw;
            if (!values.TryGetValue(primary, out raw) &&
                !values.TryGetValue(alternate, out raw)) return null;
            var text = raw as string;
            if (string.IsNullOrWhiteSpace(text))
                throw new InvalidOperationException(
                    "ANALYSIS_REQUEST_PERIOD_UNBOUND");
            return text;
        }

        private static string ReadableMetric(string metric)
        {
            return Regex.Replace(metric ?? string.Empty,
                @"([a-z])([A-Z])", "$1 $2");
        }
    }
}
