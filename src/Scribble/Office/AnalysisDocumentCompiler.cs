using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
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
        public string ComparePeriod { get; set; }
        public string FocusPeriod { get; set; }
        public List<string> ChartSeries { get; set; } =
            new List<string>();
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
            AnalysisDocumentPlan plan,
            bool requireSlides = true)
        {
            // Validate the host-issued identity as well as the schema. A
            // mutated in-memory artifact must never retain its old revision.
            AnalysisContract.Serialize(artifact);
            if (plan == null || plan.AnalysisId != artifact.AnalysisId)
                throw new InvalidOperationException(
                    "ANALYSIS_PLAN_BINDING_INVALID: The document plan must name the current analysis revision.");
            if (artifact.UnresolvedConflicts != null &&
                artifact.UnresolvedConflicts.Count > 0)
                throw new InvalidOperationException(
                    "ANALYSIS_CONFLICT_UNRESOLVED: Resolve source conflicts before producing verified outputs.");
            var facts = artifact.Facts.ToDictionary(fact => fact.FactId,
                StringComparer.Ordinal);
            var verifiedLabels = VerifiedNumericLabels(artifact);
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
                        address, null, verifiedLabels));
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
                RejectNumericLiteral(slide.Title, verifiedLabels);
                var used = new HashSet<string>(StringComparer.Ordinal);
                var map = new Dictionary<string, object>
                {
                    { "id", slide.Id }, { "layout", slide.Layout ?? string.Empty },
                    { "title", slide.Title },
                    { "subtitle", Text(slide.Subtitle, facts, used,
                        verifiedLabels) },
                    { "takeaway", Text(slide.Takeaway, facts, used,
                        verifiedLabels) }
                };
                if (slide.TableHeaders != null && slide.TableHeaders.Count > 0)
                {
                    foreach (var header in slide.TableHeaders)
                        RejectNumericLiteral(header, verifiedLabels);
                    var tableRows = new List<object>();
                    foreach (var row in slide.TableRows ?? new List<AnalysisPlanRow>())
                    {
                        if (row == null || row.Cells == null ||
                            row.Cells.Count != slide.TableHeaders.Count)
                            throw new InvalidOperationException(
                                "ANALYSIS_SLIDE_TABLE_ALIGNMENT_INVALID");
                        tableRows.Add(row.Cells.Select(cell => Cell(cell, facts,
                            null, null, used, verifiedLabels)).ToArray());
                    }
                    map["table"] = new Dictionary<string, object>
                    {
                        { "headers", slide.TableHeaders.ToArray() },
                        { "rows", tableRows.ToArray() }
                    };
                }
                if (slide.Chart != null)
                    map["chart"] = Chart(slide.Chart, facts, used,
                        verifiedLabels);
                if (slide.Cards != null && slide.Cards.Count > 0)
                    map["cards"] = slide.Cards.Select(card =>
                        (object)new Dictionary<string, object>
                        {
                            { "heading", CheckedProse(card.Heading,
                                verifiedLabels) },
                            { "points", (card.Points ??
                                new List<AnalysisPlanText>()).Select(point =>
                                    Text(new[] { point }, facts, used,
                                        verifiedLabels)).ToArray() }
                        }).ToArray();
                var cited = CompactCitations(used.SelectMany(id =>
                    facts[id].Locators ?? new List<SourceLocator>()));
                var sources = string.Join("; ", cited);
                if (sources.Length > 2000)
                    throw new InvalidOperationException("ANALYSIS_SLIDE_CITATIONS_TOO_LARGE");
                map["sources"] = sources;
                map["evidence"] = string.Join("\n", used.Select(id =>
                    facts[id].Metric + " [" + id + "] = " +
                    facts[id].Value + " (" + Citation(facts[id].Locators == null
                        ? null : facts[id].Locators.FirstOrDefault()) + ")"));
                result.Slides.Add(map);
            }
            if ((requireSlides && result.Slides.Count == 0) ||
                result.Slides.Count >
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
            string address, ISet<string> used,
            ISet<string> verifiedLabels)
        {
            if (cell == null) throw new InvalidOperationException("ANALYSIS_CELL_INVALID");
            if (expected == null && cell.Text != null &&
                cell.FactId != null && cell.Formula == null)
            {
                var source = Fact(facts, cell.FactId);
                RejectNumericLiteral(cell.Text, verifiedLabels);
                if (source.Dimensions.Values.Contains(cell.Text,
                        StringComparer.Ordinal))
                {
                    used?.Add(cell.FactId);
                    return cell.Text;
                }
            }
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
            if (expected == null)
                RejectNumericLiteral(cell.Text, verifiedLabels);
            return cell.Text;
        }

        private static string Text(IEnumerable<AnalysisPlanText> parts,
            IDictionary<string, VerifiedFact> facts, ISet<string> used,
            ISet<string> verifiedLabels)
        {
            var rendered = new List<string>();
            foreach (var part in parts ?? new AnalysisPlanText[0])
            {
                if (part == null || (part.Text == null &&
                    part.FactId == null))
                    throw new InvalidOperationException("ANALYSIS_TEXT_AUTHORITY_AMBIGUOUS");
                if (part.FactId == null)
                {
                    RejectNumericLiteral(part.Text, verifiedLabels);
                    rendered.Add(part.Text);
                }
                else
                {
                    used.Add(part.FactId);
                    var fact = Fact(facts, part.FactId);
                    if (part.Text != null)
                    {
                        RejectNumericLiteral(part.Text, verifiedLabels);
                        rendered.Add(part.Text +
                            (part.Text.Length == 0 ||
                             char.IsWhiteSpace(part.Text[part.Text.Length - 1])
                                ? string.Empty : " "));
                    }
                    rendered.Add(Display(fact) + (part.IncludeUnit &&
                        !string.IsNullOrWhiteSpace(fact.Currency)
                            ? " " + fact.Currency : string.Empty));
                }
            }
            return string.Join("", rendered);
        }

        // A fact reference cannot launder a separate model-supplied number
        // in the surrounding prose. This is a narrow structural check;
        // qualitative claims still require the Phase 3 evidence review.
        private static void RejectNumericLiteral(string value,
            ISet<string> verifiedLabels)
        {
            var remaining = value ?? string.Empty;
            if (!Regex.IsMatch(remaining, @"\d")) return;
            foreach (var label in verifiedLabels.OrderByDescending(item =>
                item.Length))
                remaining = Regex.Replace(remaining,
                    @"(?<![A-Za-z0-9])" + Regex.Escape(label) +
                    @"(?![A-Za-z0-9])", string.Empty,
                    RegexOptions.CultureInvariant);
            if (Regex.IsMatch(remaining, @"\d"))
                throw new InvalidOperationException(
                    "ANALYSIS_NUMERIC_LITERAL_UNVERIFIED: Use a verified fact reference for numbers in slide prose and tables.");
        }

        private static string CheckedProse(string value,
            ISet<string> verifiedLabels)
        {
            RejectNumericLiteral(value, verifiedLabels);
            return value ?? string.Empty;
        }

        private static object Chart(AnalysisPlanChart chart,
            IDictionary<string, VerifiedFact> facts, ISet<string> used,
            ISet<string> verifiedLabels)
        {
            RejectNumericLiteral(chart.Title, verifiedLabels);
            if (chart.Categories == null || chart.Series == null ||
                chart.Categories.Count == 0 || chart.Series.Count == 0)
                throw new InvalidOperationException("ANALYSIS_CHART_EMPTY");
            var series = new List<object>();
            foreach (var item in chart.Series)
            {
                if (item != null) RejectNumericLiteral(item.Name,
                    verifiedLabels);
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

        private static ISet<string> VerifiedNumericLabels(
            AnalysisArtifact artifact)
        {
            var labels = new HashSet<string>(StringComparer.Ordinal);
            foreach (var cell in artifact.Snapshots.SelectMany(snapshot =>
                snapshot.Tables).SelectMany(table => table.Cells).Where(cell =>
                cell.Status == AnalysisContract.Verified &&
                cell.ValueType == AnalysisContract.TextValue &&
                !string.IsNullOrWhiteSpace(cell.Value) &&
                cell.Value.Length <= 64 &&
                Regex.IsMatch(cell.Value, @"[A-Za-z]") &&
                Regex.IsMatch(cell.Value, @"\d")))
            {
                labels.Add(cell.Value);
                var identifier = Regex.Match(cell.Value,
                    @"^([A-Za-z]+[0-9]+)-[0-9]+$",
                    RegexOptions.CultureInvariant);
                if (identifier.Success)
                    labels.Add(identifier.Groups[1].Value);
            }
            foreach (var period in artifact.Facts.Select(fact =>
                fact.Period).Where(value => Regex.IsMatch(value ?? "",
                    @"^[0-9]{4}-(?:0[1-9]|1[0-2])$")))
            {
                labels.Add(period);
                labels.Add(period.Substring(0, 4));
            }
            return labels;
        }

        // Keep the exact contributing cell locators on each fact. Only the
        // rendered citation line folds consecutive cells into ranges.
        public static string[] CompactCitations(
            IEnumerable<SourceLocator> locators)
        {
            var plain = new HashSet<string>(StringComparer.Ordinal);
            var cells = new Dictionary<string, SortedSet<int>>(
                StringComparer.Ordinal);
            foreach (var locator in locators ??
                new SourceLocator[0])
            {
                if (locator == null) continue;
                var address = Regex.Match(locator.Cell ?? string.Empty,
                    @"^([A-Z]+)([1-9][0-9]*)$",
                    RegexOptions.CultureInvariant);
                int row;
                if (!address.Success ||
                    !int.TryParse(address.Groups[2].Value,
                        NumberStyles.None, CultureInfo.InvariantCulture,
                        out row) || string.IsNullOrEmpty(
                            locator.WorksheetIdentity))
                {
                    var citation = Citation(locator);
                    if (citation.Length > 0) plain.Add(citation);
                    continue;
                }
                var prefix = (locator.SourceInstanceId ?? string.Empty) +
                    " / " + locator.WorksheetIdentity + "!" +
                    address.Groups[1].Value;
                SortedSet<int> rows;
                if (!cells.TryGetValue(prefix, out rows))
                {
                    rows = new SortedSet<int>();
                    cells.Add(prefix, rows);
                }
                rows.Add(row);
            }
            foreach (var group in cells.OrderBy(item => item.Key,
                StringComparer.Ordinal))
            {
                var first = -1;
                var last = -1;
                foreach (var row in group.Value)
                {
                    if (first < 0) { first = last = row; continue; }
                    if (row == last + 1) { last = row; continue; }
                    plain.Add(group.Key + first.ToString(
                        CultureInfo.InvariantCulture) +
                        (last == first ? string.Empty : ":" +
                            group.Key.Split('!').Last() +
                            last.ToString(CultureInfo.InvariantCulture)));
                    first = last = row;
                }
                if (first >= 0)
                    plain.Add(group.Key + first.ToString(
                        CultureInfo.InvariantCulture) +
                        (last == first ? string.Empty : ":" +
                            group.Key.Split('!').Last() +
                            last.ToString(CultureInfo.InvariantCulture)));
            }
            return plain.OrderBy(value => value,
                StringComparer.Ordinal).ToArray();
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
