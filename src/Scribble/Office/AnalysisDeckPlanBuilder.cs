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
            string objective = null)
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
                objective, choices);
            var focus = selection.FocusPeriod;
            var compare = selection.ComparePeriod;
            var metrics = selection.ReportMetrics.Concat(
                    selection.ChartSeries).Distinct(StringComparer.Ordinal)
                .ToArray();
            if (metrics.Length == 0)
                throw new InvalidOperationException(
                    "ANALYSIS_DECK_METRICS_MISSING");
            Func<string, string, VerifiedFact> total = (metric, period) =>
                totals.Single(fact => fact.Metric == metric &&
                    fact.Period == period);
            var title = Narrative(choices, "Title", "Verified period results",
                70);
            var lead = Narrative(choices, "Lead",
                "Verified workbook totals for the selected period.", 100);
            var caveat = Narrative(choices, "Caveat",
                "Only captured source rows and verified calculations are included.",
                150);
            var headlineFacts = metrics.Take(2).Select(metric => total(metric,
                focus)).ToArray();
            if (headlineFacts.Length == 1)
                headlineFacts = new[] { headlineFacts[0],
                    total(metrics[0], compare) };
            var headline = new AnalysisPlanSlide {
                Id = "analysis-headline", Layout = "scorecard", Title = title,
                Subtitle = Parts(lead),
                Cards = headlineFacts.Select(fact => new AnalysisPlanCard {
                    Heading = Label(fact.Metric),
                    Points = new List<AnalysisPlanText> {
                        Reference(fact), new AnalysisPlanText {
                            Text = IsKnownSubtotal(artifact, fact) ?
                                "Known subtotal; source has blanks" :
                                "Verified period total" } } }).ToList(),
                Takeaway = Parts("Compare verified periods before drawing a trend conclusion.")
            };
            var comparison = new AnalysisPlanSlide {
                Id = "analysis-comparison", Layout = "two_pane",
                Title = "Period comparison",
                Subtitle = Parts("Native chart and table use verified source values."),
                TableHeaders = new List<string> { "Metric", compare, focus },
                TableRows = metrics.Select(metric => new AnalysisPlanRow {
                    Cells = new List<AnalysisPlanCell> {
                        new AnalysisPlanCell { Text = Label(metric) },
                        new AnalysisPlanCell { FactId = total(metric,
                            compare).FactId },
                        new AnalysisPlanCell { FactId = total(metric,
                            focus).FactId } } }).ToList(),
                Chart = new AnalysisPlanChart { Type = "column",
                    Title = "Verified values (" +
                        (!string.IsNullOrEmpty(total(selection.ChartSeries[0],
                            focus).Currency) ? total(selection.ChartSeries[0],
                                focus).Currency : total(selection.ChartSeries[0],
                                focus).Unit) + ")",
                    Categories = new List<string> { compare, focus },
                    Series = selection.ChartSeries.Select(metric =>
                        new AnalysisPlanSeries { Name = Label(metric),
                            FactIds = new List<string> {
                                total(metric, compare).FactId,
                                total(metric, focus).FactId } }).ToList() },
                Takeaway = Parts("Periods are shown in chronological order.")
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
                    .OrderBy(value => value, StringComparer.Ordinal)
                    .Take(20).ToArray();
                var groupMetrics = metrics.Where(metric => groups.All(group =>
                    dimensioned.Count(fact => fact.Metric == metric &&
                        fact.Dimensions[dimension] == group) == 1)).ToArray();
                if (groups.Length > 0 && groupMetrics.Length > 0)
                    slides.Add(new AnalysisPlanSlide {
                        Id = "analysis-groups", Layout = "table",
                        Title = "Group analysis",
                        Subtitle = Parts("Verified selected-period group totals."),
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
                        Takeaway = Parts("Each group is bound to the captured source rows.")
                    });
            }
            var limitation = SourceLimitation(artifact);
            slides.Add(new AnalysisPlanSlide {
                Id = "analysis-boundary", Layout = "cards",
                Title = "Data quality and limits",
                Subtitle = Parts("Evidence boundary for this draft."),
                Cards = new List<AnalysisPlanCard> {
                    new AnalysisPlanCard { Heading =
                        IsKnownSubtotal(artifact, total(metrics[0], focus)) ?
                            "Known subtotal" : "Verified total",
                        Points = new List<AnalysisPlanText> {
                            Reference(total(metrics[0], focus)) } },
                    new AnalysisPlanCard { Heading = "Source limits",
                        Points = Parts(limitation) },
                    new AnalysisPlanCard { Heading = "Interpretation",
                        Points = Parts(caveat) } },
                Takeaway = Parts("Review the unsaved draft against the source workbook.")
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

        private static string Narrative(IDictionary<string, object> choices,
            string name, string fallback, int maximum)
        {
            object raw;
            var value = choices.TryGetValue(name, out raw) ?
                raw as string : null;
            value = (value ?? string.Empty).Trim();
            return value.Length > 0 && value.Length <= maximum &&
                !Regex.IsMatch(value, @"\d") ? value : fallback;
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

        private static string SourceLimitation(AnalysisArtifact artifact)
        {
            if (artifact.UnresolvedConflicts.Any(item =>
                item.StartsWith("Known subtotal for ",
                    StringComparison.Ordinal)))
                return "Blank source values were excluded, never set to zero; affected figures are known subtotals.";
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
                    return "Repeated source identifiers remain in the captured rows; totals count each row.";
            }
            if (table.Cells.Any(cell => cell.Row > 0 &&
                cell.Status != AnalysisContract.Verified &&
                string.IsNullOrEmpty(cell.Formula)))
                return "Blank source cells remain unresolved and are not imputed.";
            return "Only the captured source range is covered; later edits require a new analysis.";
        }

        private static bool IsKnownSubtotal(AnalysisArtifact artifact,
            VerifiedFact fact)
        {
            return artifact.UnresolvedConflicts.Any(item =>
                item.StartsWith("Known subtotal for " + fact.Metric + " " +
                    fact.Period + " ", StringComparison.Ordinal));
        }
    }
}
