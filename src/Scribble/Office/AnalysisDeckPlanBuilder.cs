using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Scribble.Office
{
    // The model chooses concise narrative; the host chooses the evidence
    // structure, fact IDs and supported native layouts.
    public static class AnalysisDeckPlanBuilder
    {
        public static AnalysisDocumentPlan Build(AnalysisArtifact artifact,
            IDictionary<string, object> choices, int requestedSlides,
            string objective = null,
            IEnumerable<string> draftMetricOrder = null)
        {
            AnalysisContract.Serialize(artifact);
            if (artifact.Snapshots.Count != 1 ||
                artifact.Snapshots[0].Tables.Count != 1 ||
                requestedSlides < 1 ||
                choices == null || !choices.ContainsKey("AnalysisId"))
                throw new InvalidOperationException(
                    "ANALYSIS_DECK_SOURCE_UNSUPPORTED");
            var totals = artifact.Facts.Where(fact =>
                fact.Status == AnalysisContract.Verified &&
                fact.Dimensions.Count == 0 &&
                (fact.ValueType == AnalysisContract.DecimalValue ||
                 fact.ValueType == AnalysisContract.IntegerValue))
                .ToArray();
            var selection = AnalysisRequestPlan.Resolve(artifact,
                objective, choices, draftMetricOrder);
            var focus = selection.FocusPeriod;
            var compare = selection.ComparePeriod;
            var metrics = selection.ChartSeries.Concat(
                    selection.ReportMetrics).Distinct(StringComparer.Ordinal)
                .ToArray();
            if (metrics.Length == 0)
                throw new InvalidOperationException(
                    "ANALYSIS_DECK_METRICS_MISSING");
            Func<string, string, VerifiedFact> total = (metric, period) =>
                totals.Single(fact => fact.Metric == metric &&
                    fact.Period == period);
            var title = "Period results";
            var leadMetric = selection.ChartSeries[0];
            var leadFocus = total(leadMetric, focus);
            var leadCompare = total(leadMetric, compare);
            var headlineFacts = metrics.Take(2).Select(metric => total(metric,
                focus)).ToArray();
            var chartFact = total(selection.ChartSeries[0], focus);
            var chartUnit = !string.IsNullOrEmpty(selection.ChartUnit) ?
                selection.ChartUnit : !string.IsNullOrEmpty(chartFact.Currency) ?
                    chartFact.Currency : chartFact.Unit;
            var chartMetrics = string.Join(" + ",
                selection.ChartSeries.Select(Label));
            var chartTitle = chartMetrics +
                (string.IsNullOrEmpty(chartUnit) ||
                 chartMetrics.EndsWith(chartUnit,
                     StringComparison.OrdinalIgnoreCase)
                    ? string.Empty : " (" + chartUnit + ")");
            if (headlineFacts.Length == 1)
                headlineFacts = new[] { headlineFacts[0],
                    total(metrics[0], compare) };
            var headline = new AnalysisPlanSlide {
                Id = "analysis-headline", Layout = "scorecard",
                TitleParts = new List<AnalysisPlanText> {
                    Period(leadFocus), new AnalysisPlanText {
                        Text = " " + Label(leadMetric) + " " },
                    new AnalysisPlanText { FactId = leadFocus.FactId },
                    new AnalysisPlanText { Text = ", " },
                    Change(leadFocus, leadCompare) },
                Subtitle = new List<AnalysisPlanText> {
                    new AnalysisPlanText { Text = "Compared with " },
                    Period(leadCompare) },
                Cards = headlineFacts.Select(fact => new AnalysisPlanCard {
                    Heading = Label(fact.Metric),
                    Points = new List<AnalysisPlanText> {
                        Reference(fact), fact.Period == focus
                            ? Change(fact, total(fact.Metric, compare),
                                "Change from prior period: ")
                            : new AnalysisPlanText {
                                Text = "Previous period" } } }).ToList(),
                Takeaway = Finding(leadFocus, leadCompare)
            };
            var revenueMetric = metrics.FirstOrDefault(metric =>
                Regex.IsMatch(metric, "revenue|sales|income",
                    RegexOptions.IgnoreCase));
            var costMetric = metrics.FirstOrDefault(metric =>
                Regex.IsMatch(metric, "cost|expense",
                    RegexOptions.IgnoreCase));
            if (revenueMetric != null && costMetric != null &&
                revenueMetric != costMetric && headline.Cards.Count < 4)
            {
                var revenue = total(revenueMetric, focus);
                var cost = total(costMetric, focus);
                var priorRevenue = total(revenueMetric, compare);
                var priorCost = total(costMetric, compare);
                if (!string.IsNullOrWhiteSpace(revenue.Currency) &&
                    revenue.Currency == cost.Currency &&
                    revenue.Currency == priorRevenue.Currency &&
                    revenue.Currency == priorCost.Currency &&
                    decimal.Parse(revenue.Value, NumberStyles.Float,
                        CultureInfo.InvariantCulture) != 0m &&
                    decimal.Parse(priorRevenue.Value, NumberStyles.Float,
                        CultureInfo.InvariantCulture) != 0m)
                {
                    var marginChange = Calculation(
                        "margin_change_points", revenue, cost,
                        priorRevenue, priorCost);
                    marginChange.Text = "Change from prior period:";
                    headline.Cards.Add(new AnalysisPlanCard {
                        Heading = "Gross margin",
                        Points = new List<AnalysisPlanText> {
                            Calculation("margin_percent", revenue, cost),
                            marginChange } });
                }
            }
            var comparison = new AnalysisPlanSlide {
                Id = "analysis-comparison", Layout = "two_pane",
                TitleParts = new List<AnalysisPlanText> {
                    new AnalysisPlanText { Text = Label(leadMetric) + " " },
                    Change(leadFocus, leadCompare),
                    new AnalysisPlanText { Text = " from " },
                    Period(leadCompare),
                    new AnalysisPlanText { Text = " to " },
                    Period(leadFocus) },
                Subtitle = Parts("Values by selected period"),
                TableHeaders = new List<string> { "Metric", compare, focus },
                TableRows = metrics.Select(metric => new AnalysisPlanRow {
                    Cells = new List<AnalysisPlanCell> {
                        new AnalysisPlanCell { Text = Label(metric) },
                        new AnalysisPlanCell { FactId = total(metric,
                            compare).FactId },
                        new AnalysisPlanCell { FactId = total(metric,
                            focus).FactId } } }).ToList(),
                Chart = new AnalysisPlanChart { Type = "column",
                    Title = chartTitle,
                    Categories = new List<string> { compare, focus },
                    Series = selection.ChartSeries.Select(metric =>
                        new AnalysisPlanSeries { Name = Label(metric),
                            FactIds = new List<string> {
                                total(metric, compare).FactId,
                                total(metric, focus).FactId } }).ToList() },
                Takeaway = Finding(leadFocus, leadCompare)
            };
            var slides = new List<AnalysisPlanSlide> { headline, comparison };
            var dimensioned = artifact.Facts.Where(fact =>
                fact.Period == focus && fact.Dimensions.Count == 1 &&
                metrics.Contains(fact.Metric, StringComparer.Ordinal) &&
                fact.Status == AnalysisContract.Verified).ToArray();
            var dimension = dimensioned.SelectMany(fact =>
                fact.Dimensions.Keys).GroupBy(name => name,
                    StringComparer.Ordinal).OrderByDescending(group =>
                    group.Count()).Select(group => group.Key)
                .FirstOrDefault();
            if (dimension != null)
            {
                var groups = dimensioned.Where(fact =>
                    fact.Dimensions.ContainsKey(dimension)).Select(fact =>
                    fact.Dimensions[dimension]).Distinct(StringComparer.Ordinal)
                    .ToArray();
                var groupMetrics = metrics.Where(metric => groups.All(group =>
                    dimensioned.Count(fact => fact.Metric == metric &&
                        fact.Dimensions[dimension] == group) == 1)).ToArray();
                if (groups.Length > 0 && groupMetrics.Length > 0)
                {
                    var rankMetric = groupMetrics[0];
                    var ranked = groups.OrderByDescending(group => decimal.Parse(
                        dimensioned.Single(fact => fact.Metric == rankMetric &&
                            fact.Dimensions[dimension] == group).Value,
                        NumberStyles.Float, CultureInfo.InvariantCulture))
                        .ThenBy(group => group, StringComparer.Ordinal)
                        .ToArray();
                    groups = ranked
                        .Take(20)
                        .ToArray();
                    var top = dimensioned.Single(fact =>
                        fact.Metric == rankMetric &&
                        fact.Dimensions[dimension] == groups[0]);
                    slides.Add(new AnalysisPlanSlide {
                        Id = "analysis-groups", Layout = "table",
                        TitleParts = new List<AnalysisPlanText> {
                            new AnalysisPlanText { Text = groups[0] +
                                " leads " }, Period(top),
                            new AnalysisPlanText { Text = " " +
                                Label(rankMetric) } },
                        Subtitle = new List<AnalysisPlanText> {
                            new AnalysisPlanText { Text = ranked.Length == 1
                                ? "Only group in " : "Bottom: " +
                                    ranked.Last() + " in " },
                            Period(top) },
                        TableHeaders = new[] { dimension }
                            .Concat(groupMetrics.Select(Label)).ToList(),
                        TableRows = groups.Select(group => new AnalysisPlanRow {
                            Cells = new[] { new AnalysisPlanCell {
                                    Text = group } }.Concat(groupMetrics.Select(
                                    metric => new AnalysisPlanCell {
                                        FactId = dimensioned.Single(fact =>
                                            fact.Metric == metric &&
                                            fact.Dimensions[dimension] == group)
                                            .FactId })).ToList() }).ToList(),
                        Takeaway = new List<AnalysisPlanText> {
                            new AnalysisPlanText { Text = groups[0] +
                                " accounts for " },
                            Share(top, total(rankMetric, focus)),
                            new AnalysisPlanText { Text = " of " +
                                Label(rankMetric) + "." } }
                    });
                }
            }
            var limitation = SourceLimitation(artifact);
            slides.Add(new AnalysisPlanSlide {
                Id = "analysis-boundary", Layout = "cards",
                Title = "What these figures cover",
                Subtitle = Parts("Scope and exclusions"),
                Cards = new List<AnalysisPlanCard> {
                    new AnalysisPlanCard { Heading = "Scope",
                        Points = new List<AnalysisPlanText> {
                            new AnalysisPlanText { Text =
                                "Current results cover", PeriodFactId =
                                    leadFocus.FactId } } },
                    new AnalysisPlanCard { Heading = "Exclusions",
                        Points = new List<AnalysisPlanText> {
                            new AnalysisPlanText { Text = limitation } } },
                    new AnalysisPlanCard { Heading = "Comparison",
                        Points = new List<AnalysisPlanText> {
                            new AnalysisPlanText { Text =
                                "Changes compare with", PeriodFactId =
                                    leadCompare.FactId } } } },
                Takeaway = Parts("A period change does not establish a cause.")
            });
            if (slides.Count < requestedSlides)
                throw new InvalidOperationException(
                    "ANALYSIS_DECK_SECTIONS_MISSING");
            var plan = new AnalysisDocumentPlan {
                AnalysisId = artifact.AnalysisId,
                WorkbookTitle = title,
                ComparePeriod = compare, FocusPeriod = focus,
                ReportMetrics = selection.ReportMetrics.ToList(),
                ChartSeries = selection.ChartSeries.ToList(),
                WorkbookRows = AnalysisWorkbookPlanBuilder.Build(artifact,
                    selection),
                Slides = slides.Take(requestedSlides).ToList()
            };
            AnalysisDocumentCompiler.Compile(artifact, plan);
            return plan;
        }

        private static string Label(string metric)
        {
            return Regex.Replace(metric ?? string.Empty,
                @"([a-z])([A-Z])", "$1 $2");
        }

        private static List<AnalysisPlanText> Parts(string value)
        {
            return new List<AnalysisPlanText> {
                new AnalysisPlanText { Text = value } };
        }

        private static AnalysisPlanText Reference(VerifiedFact fact)
        {
            return new AnalysisPlanText { FactId = fact.FactId,
                IncludeUnit = true };
        }

        private static AnalysisPlanText Period(VerifiedFact fact)
        {
            return new AnalysisPlanText { PeriodFactId = fact.FactId };
        }

        private static AnalysisPlanText Change(VerifiedFact current,
            VerifiedFact previous, string prefix = null)
        {
            return new AnalysisPlanText { Text = prefix,
                Calculation = "percent_change", InputFactIds =
                    new List<string> { current.FactId, previous.FactId } };
        }

        private static AnalysisPlanText Share(VerifiedFact group,
            VerifiedFact total)
        {
            return new AnalysisPlanText { Calculation = "share_percent",
                InputFactIds = new List<string> {
                    group.FactId, total.FactId } };
        }

        private static AnalysisPlanText Calculation(string operation,
            params VerifiedFact[] inputs)
        {
            return new AnalysisPlanText { Calculation = operation,
                InputFactIds = inputs.Select(fact => fact.FactId).ToList() };
        }

        private static List<AnalysisPlanText> Finding(
            VerifiedFact current, VerifiedFact previous)
        {
            return new List<AnalysisPlanText> {
                new AnalysisPlanText { Text = Label(current.Metric) +
                    " is " }, Change(current, previous),
                new AnalysisPlanText { Text = " against " },
                Period(previous) };
        }

        private static string SourceLimitation(AnalysisArtifact artifact)
        {
            if (artifact.UnresolvedConflicts.Any(item =>
                item.StartsWith(AnalysisContract.KnownSubtotalPrefix,
                    StringComparison.Ordinal)))
                return "Blank cells contribute no value; totals cover known amounts only.";
            var table = artifact.Snapshots[0].Tables[0];
            var identifier = table.Cells.FirstOrDefault(cell =>
                cell.Row == 0 && cell.Value != null &&
                cell.Value.EndsWith("ID", StringComparison.OrdinalIgnoreCase));
            if (identifier != null)
            {
                var values = table.Cells.Where(cell => cell.Row > 0 &&
                    cell.Column == identifier.Column &&
                    !string.IsNullOrWhiteSpace(cell.Value))
                    .Select(cell => cell.Value).ToArray();
                if (values.Length != values.Distinct(
                        StringComparer.Ordinal).Count())
                    return "Identical repeated IDs count once.";
            }
            if (table.Cells.Any(cell => cell.Row > 0 &&
                cell.Status != AnalysisContract.Verified &&
                string.IsNullOrEmpty(cell.Formula)))
                return "Missing entries are omitted; they are not zero.";
            return "Later workbook edits are outside these figures.";
        }

        private static bool IsKnownSubtotal(AnalysisArtifact artifact,
            VerifiedFact fact)
        {
            return AnalysisContract.IsKnownSubtotal(artifact, fact.Metric,
                fact.Period);
        }
    }
}
