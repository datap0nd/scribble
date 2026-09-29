using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace Scribble.Office
{
    // The model chooses references; only the host turns verified workbook
    // facts into visible values, metric labels, units, periods and locators.
    internal sealed class RevisionFactCatalog
    {
        private static readonly Regex Reference = new Regex(
            @"\[\[fact:(fact_[0-9a-f]{24}):(value|percent|metric|period|unit|locator|dimension:[A-Za-z0-9_]+)\]\]",
            RegexOptions.Compiled);
        private static readonly Regex PeriodName = new Regex(
            @"\b(January|February|March|April|May|June|July|August|September|October|November|December)\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private readonly Dictionary<string, VerifiedFact> _facts;

        private RevisionFactCatalog(IEnumerable<VerifiedFact> facts)
        {
            _facts = facts.GroupBy(fact => fact.FactId,
                StringComparer.Ordinal).ToDictionary(group => group.Key,
                    group => group.First(), StringComparer.Ordinal);
            if (_facts.Count == 0 || _facts.Count > 250 ||
                _facts.Values.Any(fact => fact.Status !=
                    AnalysisContract.Verified ||
                    fact.FactId != AnalysisContract.ExpectedFactId(fact)))
                throw new InvalidOperationException(
                    "REVISION_FACT_CATALOG_INVALID");
        }

        internal static RevisionFactCatalog FromBindings(
            IEnumerable<PresentationDraftCopy.MonthlyChartBinding> bindings)
        {
            var observed = (bindings ??
                new PresentationDraftCopy.MonthlyChartBinding[0])
                .SelectMany(binding => binding.Facts?.Facts ??
                    new VerifiedFact[0]).GroupBy(fact => fact.FactId,
                    StringComparer.Ordinal).Select(group => group.First())
                .ToList();
            var derived = new List<VerifiedFact>();
            var totals = observed.Where(fact =>
                fact.Dimensions.Count == 0).ToArray();
            foreach (var series in totals.GroupBy(fact => fact.Metric,
                StringComparer.Ordinal))
            {
                var periods = series.OrderBy(fact => fact.Period,
                    StringComparer.Ordinal).ToArray();
                for (var index = 1; index < periods.Length; index++)
                {
                    VerifiedFact growth;
                    try
                    {
                        AnalysisCalculator.Calculate(
                            AnalysisCalculator.Growth,
                            new[] { periods[index], periods[index - 1] },
                            series.Key + " change", periods[index].Period,
                            out growth);
                        derived.Add(growth);
                    }
                    catch (InvalidOperationException) { }
                }
            }
            foreach (var period in totals.GroupBy(fact => fact.Period,
                StringComparer.Ordinal))
            {
                var revenue = period.SingleOrDefault(fact =>
                    Regex.IsMatch(fact.Metric, @"\brevenue\b",
                        RegexOptions.IgnoreCase));
                var cost = period.SingleOrDefault(fact =>
                    Regex.IsMatch(fact.Metric, @"\bcost\b",
                        RegexOptions.IgnoreCase));
                if (revenue == null || cost == null) continue;
                VerifiedFact margin;
                try
                {
                    AnalysisCalculator.Calculate(
                        AnalysisCalculator.Margin,
                        new[] { revenue, cost }, "Gross margin",
                        period.Key, out margin);
                    derived.Add(margin);
                }
                catch (InvalidOperationException) { }
            }
            return new RevisionFactCatalog(observed.Concat(derived));
        }

        internal object[] PublicFacts()
        {
            return _facts.Values.OrderBy(fact => fact.Metric,
                StringComparer.Ordinal).ThenBy(fact => fact.Period,
                StringComparer.Ordinal).ThenBy(fact =>
                string.Join("/", fact.Dimensions.OrderBy(pair =>
                    pair.Key).Select(pair => pair.Value)))
                .Select(fact => (object)new {
                    fact_id = fact.FactId,
                    metric = fact.Metric,
                    period = fact.Period,
                    dimensions = fact.Dimensions,
                    value = DisplayValue(fact),
                    unit = DisplayUnit(fact),
                    locator = SafeLocator(fact)
                }).ToArray();
        }

        internal object[] BindOperations(object[] operations)
        {
            if (operations == null || operations.Length == 0)
                throw new InvalidOperationException(
                    "REVISION_FACT_OPERATIONS_INVALID");
            var serializer = new JavaScriptSerializer {
                MaxJsonLength = 16000000
            };
            var copy = serializer.DeserializeObject(
                serializer.Serialize(operations)) as object[];
            if (copy == null)
                throw new InvalidOperationException(
                    "REVISION_FACT_OPERATIONS_INVALID");
            foreach (var raw in copy)
            {
                var operation = raw as Dictionary<string, object>;
                if (operation == null)
                    throw new InvalidOperationException(
                        "REVISION_FACT_OPERATIONS_INVALID");
                var kind = Convert.ToString(operation["kind"]);
                if (kind == "replace_text" || kind == "table_cell")
                    operation["text"] = Render(Convert.ToString(
                        operation["text"]));
                else if (kind == "notes_append")
                    operation["notes"] = Render(Convert.ToString(
                        operation["notes"]));
                else if (kind == "replace_slide")
                    BindSlide(operation["slide"]);
            }
            return copy;
        }

        private void BindSlide(object value)
        {
            var map = value as Dictionary<string, object>;
            if (map == null)
                throw new InvalidOperationException(
                    "REVISION_FACT_SLIDE_INVALID");
            foreach (var key in map.Keys.ToArray())
            {
                if (new[] { "id", "layout", "purpose", "content_kind",
                    "source_spans", "sources", "evidence", "image_names" }
                    .Contains(key)) continue;
                var text = map[key] as string;
                if (text != null) map[key] = Render(text);
                else if (map[key] is Dictionary<string, object>)
                    BindSlide(map[key]);
                else if (map[key] is IEnumerable &&
                    !(map[key] is string))
                {
                    var list = ((IEnumerable)map[key]).Cast<object>()
                        .ToArray();
                    for (var index = 0; index < list.Length; index++)
                    {
                        if (list[index] is string)
                            list[index] = Render((string)list[index]);
                        else if (list[index] is
                            Dictionary<string, object>)
                            BindSlide(list[index]);
                    }
                    map[key] = list;
                }
            }
        }

        internal string Render(string text)
        {
            text = text ?? string.Empty;
            var output = new StringBuilder();
            var offset = 0;
            foreach (Match match in Reference.Matches(text))
            {
                var literal = text.Substring(offset, match.Index - offset);
                RejectLiteral(literal);
                output.Append(literal);
                VerifiedFact fact;
                if (!_facts.TryGetValue(match.Groups[1].Value, out fact))
                    throw new InvalidOperationException(
                        "REVISION_FACT_ID_UNKNOWN");
                output.Append(Resolve(fact, match.Groups[2].Value));
                offset = match.Index + match.Length;
            }
            var tail = text.Substring(offset);
            RejectLiteral(tail);
            output.Append(tail);
            return output.ToString();
        }

        private void RejectLiteral(string literal)
        {
            if (literal.IndexOf("[[fact:",
                    StringComparison.OrdinalIgnoreCase) >= 0 ||
                Regex.IsMatch(literal, @"\d") ||
                PeriodName.IsMatch(literal) ||
                _facts.Values.Select(fact => fact.Metric)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Any(metric => metric.Length > 2 &&
                        Regex.IsMatch(literal,
                            @"\b" + Regex.Escape(metric) + @"\b",
                            RegexOptions.IgnoreCase)) ||
                _facts.Values.Select(DisplayUnit)
                    .Where(unit => unit.Length > 1)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Any(unit => Regex.IsMatch(literal,
                        @"\b" + Regex.Escape(unit) + @"\b",
                        RegexOptions.IgnoreCase)))
                throw new InvalidOperationException(
                    "REVISION_FACT_LITERAL_UNBOUND");
        }

        private static string Resolve(VerifiedFact fact, string field)
        {
            if (field == "value") return DisplayValue(fact);
            if (field == "percent")
            {
                if (fact.Unit != "ratio")
                    throw new InvalidOperationException(
                        "REVISION_FACT_PERCENT_UNIT_INVALID");
                return AnalysisContract.Decimal(fact).ToString("0.##%",
                    CultureInfo.InvariantCulture);
            }
            if (field == "metric") return fact.Metric;
            if (field == "unit") return DisplayUnit(fact);
            if (field == "locator") return SafeLocator(fact);
            if (field == "period")
            {
                DateTime date;
                return DateTime.TryParseExact(fact.Period, "yyyy-MM",
                    CultureInfo.InvariantCulture, DateTimeStyles.None,
                    out date) ? date.ToString("MMMM yyyy",
                        CultureInfo.InvariantCulture) : fact.Period;
            }
            if (field.StartsWith("dimension:",
                    StringComparison.Ordinal))
            {
                string dimension;
                if (!fact.Dimensions.TryGetValue(field.Substring(10),
                        out dimension))
                    throw new InvalidOperationException(
                        "REVISION_FACT_DIMENSION_UNKNOWN");
                return dimension;
            }
            throw new InvalidOperationException(
                "REVISION_FACT_FIELD_INVALID");
        }

        private static string DisplayValue(VerifiedFact fact)
        {
            var value = AnalysisContract.Decimal(fact);
            return value.ToString(value == decimal.Truncate(value)
                ? "#,##0" : "#,##0.##", CultureInfo.InvariantCulture);
        }

        private static string DisplayUnit(VerifiedFact fact)
        {
            return string.IsNullOrWhiteSpace(fact.Currency)
                ? fact.Unit ?? string.Empty : fact.Currency;
        }

        private static string SafeLocator(VerifiedFact fact)
        {
            var locator = fact.Locators.FirstOrDefault();
            if (locator == null)
                throw new InvalidOperationException(
                    "REVISION_FACT_LOCATOR_MISSING");
            var sheet = locator.WorksheetIdentity ?? string.Empty;
            var cell = !string.IsNullOrWhiteSpace(locator.Range)
                ? locator.Range : locator.Cell;
            if (sheet.Length == 0 || string.IsNullOrWhiteSpace(cell))
                throw new InvalidOperationException(
                    "REVISION_FACT_LOCATOR_MISSING");
            return sheet + "!" + cell;
        }
    }
}
