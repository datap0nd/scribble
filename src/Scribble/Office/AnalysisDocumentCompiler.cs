using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Scribble.Security;

namespace Scribble.Office
{
    // A semantic plan supplies wording and structure. Numeric output and
    // citations are resolved from one validated analysis at compile time.
    // This is an adapter to the existing Office writers, not a new renderer.
    public sealed class AnalysisDocumentPlan
    {
        public string AnalysisId { get; set; }
        public string WorkbookTitle { get; set; }
        public List<AnalysisPlanRow> WorkbookRows { get; set; } =
            new List<AnalysisPlanRow>();
        public List<AnalysisPlanSlide> Slides { get; set; } =
            new List<AnalysisPlanSlide>();
    }

    public sealed class AnalysisPlanRow
    {
        public List<AnalysisPlanCell> Cells { get; set; } =
            new List<AnalysisPlanCell>();
    }

    public sealed class AnalysisPlanCell
    {
        public string Text { get; set; }
        public string FactId { get; set; }
        // Formula is host-authored for the hand-authored pilot. Its expected
        // fact is checked against native recalculation after the write.
        public string Formula { get; set; }
        public string ExpectedFactId { get; set; }
    }

    public sealed class AnalysisPlanText
    {
        public string Text { get; set; }
        public string FactId { get; set; }
        public bool IncludeUnit { get; set; }
    }

    public sealed class AnalysisPlanSlide
    {
        public string Id { get; set; }
        public string Layout { get; set; }
        public string Title { get; set; }
        public List<AnalysisPlanText> Subtitle { get; set; } =
            new List<AnalysisPlanText>();
        public List<AnalysisPlanText> Takeaway { get; set; } =
            new List<AnalysisPlanText>();
        public List<string> TableHeaders { get; set; } =
            new List<string>();
        public List<AnalysisPlanRow> TableRows { get; set; } =
            new List<AnalysisPlanRow>();
        public AnalysisPlanChart Chart { get; set; }
        public List<AnalysisPlanCard> Cards { get; set; } =
            new List<AnalysisPlanCard>();
    }

    public sealed class AnalysisPlanChart
    {
        public string Type { get; set; }
        public string Title { get; set; }
        public List<string> Categories { get; set; } =
            new List<string>();
        public List<AnalysisPlanSeries> Series { get; set; } =
            new List<AnalysisPlanSeries>();
    }

    public sealed class AnalysisPlanSeries
    {
        public string Name { get; set; }
        public List<string> FactIds { get; set; } =
            new List<string>();
    }

    public sealed class AnalysisPlanCard
    {
        public string Heading { get; set; }
        public List<AnalysisPlanText> Points { get; set; } =
            new List<AnalysisPlanText>();
    }

    public sealed class CompiledAnalysisDocuments
    {
        public string AnalysisId { get; set; }
        public string WorkbookTitle { get; set; }
        public List<List<string>> WorkbookRows { get; set; } =
            new List<List<string>>();
        public Dictionary<string, string> ExpectedFormulaFacts { get; set; } =
            new Dictionary<string, string>();
        public List<Dictionary<string, object>> Slides { get; set; } =
            new List<Dictionary<string, object>>();
    }

    public static class AnalysisDocumentCompiler
    {
        public static CompiledAnalysisDocuments Compile(
            AnalysisArtifact artifact,
            AnalysisDocumentPlan plan)
        {
            AnalysisContract.Validate(artifact);
            if (plan == null || plan.AnalysisId != artifact.AnalysisId)
                throw new InvalidOperationException(
                    "ANALYSIS_PLAN_BINDING_INVALID: The document plan must name the current analysis revision.");
            if (artifact.UnresolvedConflicts != null &&
                artifact.UnresolvedConflicts.Count > 0)
                throw new InvalidOperationException(
                    "ANALYSIS_CONFLICT_UNRESOLVED: Resolve source conflicts before producing verified outputs.");
            var facts = artifact.Facts.ToDictionary(fact => fact.FactId,
                StringComparer.Ordinal);
            var result = new CompiledAnalysisDocuments
            {
                AnalysisId = artifact.AnalysisId,
                WorkbookTitle = plan.WorkbookTitle ?? string.Empty
            };
            foreach (var row in plan.WorkbookRows ?? new List<AnalysisPlanRow>())
            {
                if (row == null || row.Cells == null)
                    throw new InvalidOperationException("ANALYSIS_PLAN_ROW_INVALID");
                if (row.Cells.Count == 0 ||
                    row.Cells.Count > WorkbookDraftWriter.MaxDraftColumns)
                    throw new InvalidOperationException("ANALYSIS_WORKBOOK_COLUMNS_INVALID");
                var values = new List<string>();
                foreach (var cell in row.Cells)
                {
                    var address = Address(result.WorkbookRows.Count + 3,
                        values.Count + 1);
                    values.Add(Cell(cell, facts, result.ExpectedFormulaFacts,
                        address, null));
                }
                result.WorkbookRows.Add(values);
            }
            if (result.WorkbookRows.Count == 0 || result.WorkbookRows.Count >
                WorkbookDraftWriter.MaxDraftRows)
                throw new InvalidOperationException("ANALYSIS_WORKBOOK_ROWS_INVALID");
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var slide in plan.Slides ?? new List<AnalysisPlanSlide>())
            {
                if (slide == null || string.IsNullOrWhiteSpace(slide.Id) ||
                    !ids.Add(slide.Id) || string.IsNullOrWhiteSpace(slide.Title))
                    throw new InvalidOperationException("ANALYSIS_SLIDE_ID_INVALID");
                var used = new HashSet<string>(StringComparer.Ordinal);
                var map = new Dictionary<string, object>
                {
                    { "id", slide.Id }, { "layout", slide.Layout ?? string.Empty },
                    { "title", slide.Title },
                    { "subtitle", Text(slide.Subtitle, facts, used) },
                    { "takeaway", Text(slide.Takeaway, facts, used) }
                };
                if (slide.TableHeaders != null && slide.TableHeaders.Count > 0)
                {
                    var tableRows = new List<object>();
                    foreach (var row in slide.TableRows ?? new List<AnalysisPlanRow>())
                    {
                        if (row == null || row.Cells == null ||
                            row.Cells.Count != slide.TableHeaders.Count)
                            throw new InvalidOperationException(
                                "ANALYSIS_SLIDE_TABLE_ALIGNMENT_INVALID");
                        tableRows.Add(row.Cells.Select(cell => Cell(cell, facts,
                            null, null, used)).ToArray());
                    }
                    map["table"] = new Dictionary<string, object>
                    {
                        { "headers", slide.TableHeaders.ToArray() },
                        { "rows", tableRows.ToArray() }
                    };
                }
                if (slide.Chart != null)
                    map["chart"] = Chart(slide.Chart, facts, used);
                if (slide.Cards != null && slide.Cards.Count > 0)
                    map["cards"] = slide.Cards.Select(card =>
                        (object)new Dictionary<string, object>
                        {
                            { "heading", card.Heading ?? string.Empty },
                            { "points", (card.Points ??
                                new List<AnalysisPlanText>()).Select(point =>
                                    Text(new[] { point }, facts, used)).ToArray() }
                        }).ToArray();
                var cited = used.SelectMany(id => facts[id].Locators ??
                    new List<SourceLocator>()).Select(Citation).Where(value =>
                    value.Length > 0).Distinct(StringComparer.Ordinal).ToArray();
                var sources = string.Join("; ", cited);
                if (sources.Length > 2000)
                    throw new InvalidOperationException("ANALYSIS_SLIDE_CITATIONS_TOO_LARGE");
                map["sources"] = sources;
                map["footnote"] = sources.Length <= 200 ? sources :
                    sources.Substring(0, 197) + "...";
                map["evidence"] = string.Join("\n", used.Select(id =>
                    facts[id].Metric + " [" + id + "] = " +
                    facts[id].Value + " (" + Citation(facts[id].Locators == null
                        ? null : facts[id].Locators.FirstOrDefault()) + ")"));
                result.Slides.Add(map);
            }
            if (result.Slides.Count == 0 || result.Slides.Count >
                PresentationDraftWriter.MaxDraftSlides)
                throw new InvalidOperationException("ANALYSIS_SLIDE_COUNT_INVALID");
            // Let the current writer validate all layout, density, chart and
            // table constraints before any Office mutation.
            PresentationDraftWriter.ParseSlides(result.Slides.Cast<object>().ToArray());
            return result;
        }

        private static string Cell(AnalysisPlanCell cell,
            IDictionary<string, VerifiedFact> facts,
            IDictionary<string, string> expected,
            string address, ISet<string> used)
        {
            if (cell == null) throw new InvalidOperationException("ANALYSIS_CELL_INVALID");
            var choices = (cell.Text != null ? 1 : 0) +
                (cell.FactId != null ? 1 : 0) +
                (cell.Formula != null ? 1 : 0);
            if (choices != 1) throw new InvalidOperationException(
                "ANALYSIS_CELL_AUTHORITY_AMBIGUOUS");
            if (cell.FactId != null)
            {
                used?.Add(cell.FactId);
                var fact = Fact(facts, cell.FactId);
                // Workbook cells carry invariant raw values; presentation
                // cells carry display formatting. This keeps native Excel
                // values numeric across locales.
                return expected == null ? Display(fact) : fact.Value;
            }
            if (cell.Formula != null)
            {
                if (expected == null || string.IsNullOrWhiteSpace(address) ||
                    !DraftFormulaPolicy.IsAllowedFormula(cell.Formula))
                    throw new InvalidOperationException("ANALYSIS_FORMULA_INVALID");
                Fact(facts, cell.ExpectedFactId);
                expected.Add(address, cell.ExpectedFactId);
                return cell.Formula;
            }
            if (cell.Text.StartsWith("=", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "ANALYSIS_LITERAL_FORMULA_INVALID: Formula cells require an expected fact binding.");
            return cell.Text;
        }

        private static string Text(IEnumerable<AnalysisPlanText> parts,
            IDictionary<string, VerifiedFact> facts, ISet<string> used)
        {
            var rendered = new List<string>();
            foreach (var part in parts ?? new AnalysisPlanText[0])
            {
                if (part == null || (part.Text == null) ==
                    (part.FactId == null))
                    throw new InvalidOperationException("ANALYSIS_TEXT_AUTHORITY_AMBIGUOUS");
                if (part.FactId == null) rendered.Add(part.Text);
                else
                {
                    used.Add(part.FactId);
                    var fact = Fact(facts, part.FactId);
                    rendered.Add(Display(fact) + (part.IncludeUnit &&
                        !string.IsNullOrWhiteSpace(fact.Currency)
                            ? " " + fact.Currency : string.Empty));
                }
            }
            return string.Join("", rendered);
        }

        private static object Chart(AnalysisPlanChart chart,
            IDictionary<string, VerifiedFact> facts, ISet<string> used)
        {
            if (chart.Categories == null || chart.Series == null ||
                chart.Categories.Count == 0 || chart.Series.Count == 0)
                throw new InvalidOperationException("ANALYSIS_CHART_EMPTY");
            var series = new List<object>();
            foreach (var item in chart.Series)
            {
                if (item == null || item.FactIds == null ||
                    item.FactIds.Count != chart.Categories.Count)
                    throw new InvalidOperationException("ANALYSIS_CHART_ALIGNMENT_INVALID");
                var values = new List<object>();
                for (var category = 0; category < item.FactIds.Count;
                    category++)
                {
                    var id = item.FactIds[category];
                    if (string.IsNullOrWhiteSpace(id))
                    {
                        values.Add(null);
                        continue;
                    }
                    used.Add(id);
                    var fact = Fact(facts, id);
                    if (!string.IsNullOrWhiteSpace(fact.Period) &&
                        !string.Equals(fact.Period,
                            chart.Categories[category],
                            StringComparison.Ordinal))
                        throw new InvalidOperationException(
                            "ANALYSIS_CHART_PERIOD_MISMATCH");
                    if (fact.ValueType != AnalysisContract.DecimalValue &&
                        fact.ValueType != AnalysisContract.IntegerValue)
                        throw new InvalidOperationException("ANALYSIS_CHART_FACT_NOT_NUMERIC");
                    var number = double.Parse(fact.Value,
                        CultureInfo.InvariantCulture);
                    if (double.IsNaN(number) || double.IsInfinity(number))
                        throw new InvalidOperationException("ANALYSIS_CHART_VALUE_INVALID");
                    values.Add(number);
                }
                series.Add(new Dictionary<string, object>
                {
                    { "name", item.Name ?? string.Empty },
                    { "values", values.ToArray() }
                });
            }
            return new Dictionary<string, object>
            {
                { "type", chart.Type ?? "column" },
                { "title", chart.Title ?? string.Empty },
                { "categories", chart.Categories.ToArray() },
                { "series", series.ToArray() }
            };
        }

        private static VerifiedFact Fact(
            IDictionary<string, VerifiedFact> facts, string id)
        {
            VerifiedFact fact;
            if (string.IsNullOrWhiteSpace(id) ||
                !facts.TryGetValue(id, out fact) ||
                fact.Status != AnalysisContract.Verified)
                throw new InvalidOperationException(
                    "ANALYSIS_FACT_NOT_VERIFIED: " + id);
            return fact;
        }

        private static string Display(VerifiedFact fact)
        {
            if (fact.ValueType != AnalysisContract.DecimalValue &&
                fact.ValueType != AnalysisContract.IntegerValue)
                return fact.Value;
            var value = decimal.Parse(fact.Value,
                NumberStyles.Float, CultureInfo.InvariantCulture);
            return fact.Unit == "ratio"
                ? (value * 100m).ToString("0.00", CultureInfo.InvariantCulture) + "%"
                : value.ToString("#,##0.##", CultureInfo.InvariantCulture);
        }

        private static string Citation(SourceLocator locator)
        {
            if (locator == null) return string.Empty;
            var source = locator.SourceInstanceId ?? string.Empty;
            var sheet = locator.WorksheetIdentity ?? string.Empty;
            var range = locator.Cell ?? locator.Range ?? string.Empty;
            return source + (sheet.Length == 0 ? string.Empty : " / " + sheet) +
                (range.Length == 0 ? string.Empty : "!" + range);
        }

        private static string Address(int row, int column)
        {
            var letters = string.Empty;
            while (column > 0)
            {
                column--;
                letters = (char)('A' + column % 26) + letters;
                column /= 26;
            }
            return letters + row.ToString(CultureInfo.InvariantCulture);
        }
    }
}
