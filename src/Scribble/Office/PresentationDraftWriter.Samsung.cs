using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using Scribble.Security;

namespace Scribble.Office
{
    internal static partial class PresentationDraftWriter
    {
        internal static void TraceNativeChartStage(string stage)
        {
            var path = Environment.GetEnvironmentVariable(
                "SCRIBBLE_NATIVE_CHART_TRACE");
            if (string.IsNullOrWhiteSpace(path)) return;
            try { File.AppendAllText(path, stage + Environment.NewLine); }
            catch { }
        }
        private static object SamsungValue(IDictionary<string, object> map, string key)
        { object value; return map.TryGetValue(key, out value) ? value : null; }
        private static string SamsungString(IDictionary<string, object> map, string key, int limit)
        { var value = Convert.ToString(SamsungValue(map, key)) ?? ""; if (value.Length > limit) throw new InvalidOperationException(key + " exceeds " + limit + " characters; split the content."); return value; }
        private static IReadOnlyList<int> SamsungIndices(IDictionary<string, object> map)
        {
            var values = SamsungValue(map, "highlight_rows") as System.Collections.IEnumerable;
            if (values == null) return new int[0];
            return values.Cast<object>().Select(Convert.ToInt32).Distinct().ToArray();
        }
        private static void ValidateSamsungInput(IDictionary<string, object> map)
        {
            foreach (var key in new[] { "title", "subtitle", "footnote", "unit" })
                SamsungString(map, key, key == "title" ? MaxTitleCharacters : key == "subtitle" ? MaxSubtitleCharacters : key == "unit" ? MaxUnitCharacters : MaxFootnoteCharacters);
            var layout = SamsungString(map, "layout", 40).ToLowerInvariant();
            if (layout.Length > 0 && !SamsungSlideDesign.Layouts.Contains(layout)) throw new InvalidOperationException("Unknown Samsung layout: " + layout);
            ValidateArray(SamsungValue(map, "bullets"), MaxBulletsPerSlide, MaxBulletCharacters);
            var cards = ValidateArray(SamsungValue(map, "cards"), MaxCards, 0);
            foreach (var card in cards)
            {
                var data = card as IDictionary<string, object>;
                if (data == null) throw new InvalidOperationException("Each card must be an object.");
                SamsungString(data, "heading", 120); ValidateArray(SamsungValue(data, "points"), MaxCardPoints, MaxBulletCharacters);
            }
            foreach (var key in new[] { "table", "secondary_table" })
            {
                var data = SamsungValue(map, key) as IDictionary<string, object>;
                if (SamsungValue(map, key) != null && data == null) throw new InvalidOperationException("Invalid table object.");
                if (data == null) continue;
                var columns = ValidateArray(SamsungValue(data, "headers"), MaxTableColumns, MaxCellCharacters).Length;
                foreach (var row in ValidateArray(SamsungValue(data, "rows"), MaxTableRows, 0))
                {
                    var cells = ValidateArray(row, MaxTableColumns, MaxCellCharacters);
                    if (columns == 0) columns = cells.Length;
                    if (cells.Length != columns) throw new InvalidOperationException("Table rows must preserve exact column alignment.");
                }
            }
            foreach (var key in new[] { "chart", "secondary_chart" })
            {
                var data = SamsungValue(map, key) as IDictionary<string, object>;
                if (SamsungValue(map, key) != null && data == null) throw new InvalidOperationException("Invalid chart object.");
                if (data == null) continue;
                var categories = ValidateArray(SamsungValue(data, "categories"), MaxChartCategories, 80);
                var series = ValidateArray(SamsungValue(data, "series"), MaxChartSeries, 0);
                if (categories.Length == 0 || series.Length == 0) throw new InvalidOperationException("Chart categories and series are required.");
                foreach (var item in series)
                {
                    var entry = item as IDictionary<string, object>;
                    if (entry == null) throw new InvalidOperationException("Invalid chart series.");
                    var values = ValidateArray(SamsungValue(entry, "values"), MaxChartCategories, 0);
                    if (values.Length != categories.Length) throw new InvalidOperationException("Chart values must match every category; missing values cannot become zero.");
                    foreach (var value in values)
                    { if (value == null) continue; double number; if (!double.TryParse(Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out number) || double.IsNaN(number) || double.IsInfinity(number)) throw new InvalidOperationException("Chart values must be finite numbers from the source."); }
                }
            }
        }
        private static object[] ValidateArray(object value, int maximum, int characters)
        {
            if (value == null) return new object[0];
            var array = value as System.Collections.IEnumerable;
            if (array == null || value is string) throw new InvalidOperationException("Expected an array.");
            var items = array.Cast<object>().ToArray();
            if (items.Length > maximum || (characters > 0 && items.Any(i => Convert.ToString(i).Length > characters)))
                throw new InvalidOperationException("Slide content exceeds a batch limit; split it across slides. No values were truncated.");
            return items;
        }
        internal sealed class SamsungElement
        {
            internal RectangleF Box;
            internal string Text = "", Font = "Arial", Fill, Color = "#000000";
            internal float Size = SamsungSlideDesign.BodySize, Minimum = SamsungSlideDesign.BodyMinimum;
            internal float[] ColumnWidths;
            internal bool Bold, Hollow;
            internal bool Circle, Connector;
            internal int Alignment = 1;
            internal DraftTable Table;
            internal DraftChart Chart;
            internal string ImageData;
            internal DraftChart HighlightChart;
            internal int HighlightCategory;
            internal int HighlightSeries;
        }
        internal sealed class SamsungPage
        {
            internal DraftSlide Source;
            internal List<SamsungElement> Elements = new List<SamsungElement>();
            internal string Background = "#FFFFFF";
            internal SamsungElement PageNumber;
        }
        internal sealed class SamsungOutput
        {
            internal object Slide;
            internal SamsungPage Page;
            internal string Owner;
            internal List<int> ShapeIds = new List<int>();
            internal string Image;
        }

        private static SamsungElement TextElement(string text, RectangleF box, float size = 18, float minimum = 14,
            string font = "Arial", bool bold = false, string fill = null, string color = "#000000")
        { return new SamsungElement { Text = text ?? "", Box = box, Size = size, Minimum = minimum, Font = font, Bold = bold, Fill = fill, Color = color }; }

        internal static bool RetryableSamsungChartFailure(string failure)
        {
            var message = failure ?? string.Empty;
            return message.IndexOf("write chart data: COMException 0x800A01A8",
                StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("AddChart2: COMException 0x80004005 The chart data grid is already open",
                    StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("SetSourceData: COMException 0x800A01A8",
                    StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("series readback: COMException 0x800A01A8",
                    StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("ChartData.Workbook: COMException 0x80010001",
                    StringComparison.OrdinalIgnoreCase) >= 0;
        }

        internal static List<SamsungPage> ComposeSamsung(IReadOnlyList<DraftSlide> drafts)
        {
            var pages = new List<SamsungPage>();
            foreach (var draft in drafts)
            {
                var highlightCount = draft.Table != null ? draft.Table.Rows.Count : draft.Chart?.Categories.Count ?? 0;
                if (draft.HighlightRows.Any(i => i < 1 || i > highlightCount)) throw new InvalidOperationException("SLIDE_HIGHLIGHT_INVALID: Reference an existing data row or chart category.");
                if (draft.HighlightRows.Count > 0 && draft.Table == null && draft.Chart != null &&
                    new[] { DraftChartTypes.Pie, DraftChartTypes.Scatter }.Contains(draft.Chart.TypeCode))
                    throw new InvalidOperationException("SLIDE_HIGHLIGHT_UNSUPPORTED: Use a bar, column or line chart for category highlight frames.");
                var rows = draft.Table?.Rows.Count ?? 0;
                var perPage = SamsungSlideDesign.RowsPerPage(draft.Table?.Headers.Count ?? 1);
                var secondaryRows = draft.SecondaryTable?.Rows.Count ?? 0;
                if (draft.Layout == "two_pane") perPage = 7;
                // Estimate wrapped row height at the minimum permitted size,
                // then paginate before COM work instead of clipping long cells.
                foreach (var data in new[] { draft.Table, draft.SecondaryTable }.Where(t => t != null))
                {
                    var area = SamsungSlideDesign.Regions(draft.Layout)[draft.Layout == "two_pane" ? 1 : 0];
                    var cols = Math.Max(1, Math.Max(data.Headers.Count, data.Rows.Select(r => r.Count).DefaultIfEmpty(0).Max()));
                    foreach (var row in data.Rows)
                    {
                        var charsPerLine = Math.Max(1, (int)((area.Width / cols - 8) / 4));
                        var lines = row.Select(v => Math.Max(1, (v.Length + charsPerLine - 1) / charsPerLine)).DefaultIfEmpty(1).Max();
                        perPage = Math.Min(perPage, Math.Max(1, (int)(area.Height / (lines * 9 + 3)) - 1));
                    }
                }
                var count = Math.Max(1, (int)Math.Ceiling(Math.Max(rows, secondaryRows) / (double)perPage));
                for (var part = 0; part < count; part++)
                {
                    var table = SliceTable(draft.Table, part * perPage, perPage);
                    var secondaryTable = SliceTable(draft.SecondaryTable, part * perPage, perPage);
                    var page = ComposeSamsungPage(draft, table, secondaryTable, pages.Count + 1, part, count, perPage);
                    foreach (var element in page.Elements)
                    {
                        if (!SamsungSlideDesign.InBounds(element.Box)) throw new InvalidOperationException("SLIDE_GEOMETRY_INVALID");
                        if (element.Table != null) FitTable(element);
                        else if (element.Chart == null && element.ImageData == null && !element.Hollow && element.Text.Length > 0)
                            element.Size = SamsungSlideDesign.Fit(element.Text, element.Font, element.Box, element.Size, element.Minimum, element.Bold);
                    }
                    pages.Add(page);
                }
            }
            return pages;
        }
        private static DraftTable SliceTable(DraftTable table, int offset, int count)
        { return table == null ? null : new DraftTable(table.Headers, table.Rows.Skip(offset).Take(count).ToArray()); }

        private static SamsungPage ComposeSamsungPage(DraftSlide draft, DraftTable table, DraftTable secondaryTable, int index, int part, int parts, int perPage)
        {
            var page = new SamsungPage { Source = draft };
            var regions = SamsungSlideDesign.Regions(draft.Layout);
            var elements = page.Elements;
            var special = draft.Layout == "cover" || draft.Layout == "divider" || draft.Layout == "closing";
            if (special)
            {
                if (draft.Layout == "closing") page.Background = SamsungSlideDesign.Blue;
                var coverTitleSize = draft.Layout == "cover" ?
                    (draft.Title.Length > 64 ? 42 : 66) : 40;
                elements.Add(TextElement(draft.Title, regions[0], coverTitleSize, 28,
                    MetoTheme.TitleFont, true, null, draft.Layout == "closing" ? "#FFFFFF" : "#000000"));
                if (draft.Subtitle.Length > 0) elements.Add(TextElement(draft.Subtitle, SamsungSlideDesign.Percent(4.6f, 78, 84.4f, 8), 22, 18, color: draft.Layout == "closing" ? "#FFFFFF" : "#000000"));
                if (draft.Layout == "cover") elements.Add(TextElement("", SamsungSlideDesign.Percent(0, 94.7f, 100, 2.1f), fill: SamsungSlideDesign.Blue));
                if (draft.Bullets.Count > 0 || draft.Cards.Count > 0 || table != null || draft.Chart != null || draft.SecondaryTable != null || draft.SecondaryChart != null || draft.ImageNames.Count > 0 || draft.Takeaway.Length > 0 || draft.Caption.Length > 0 || draft.Unit.Length > 0)
                    throw new InvalidOperationException("Cover/divider/closing accepts title and subtitle only; put supporting content on a content slide.");
            }
            else
            {
                elements.Add(TextElement(draft.Title + (parts > 1 ? " (" + (part + 1) + "/" + parts + ")" : ""), SamsungSlideDesign.Title, SamsungSlideDesign.TitleSize, 18, MetoTheme.TitleFont, true));
                elements.Add(TextElement(draft.Subtitle, SamsungSlideDesign.Action, SamsungSlideDesign.ActionSize, SamsungSlideDesign.ActionSize));
                var queue = new Queue<SamsungElement>();
                if (table != null) queue.Enqueue(new SamsungElement { Table = table });
                if (draft.Chart != null) queue.Enqueue(new SamsungElement { Chart = draft.Chart });
                if (secondaryTable != null) queue.Enqueue(new SamsungElement { Table = secondaryTable });
                if (draft.SecondaryChart != null) queue.Enqueue(new SamsungElement { Chart = draft.SecondaryChart });
                foreach (var image in draft.ImageData) queue.Enqueue(new SamsungElement { ImageData = image });
                if (draft.ImageNames.Count != draft.ImageData.Count) throw new InvalidOperationException("SLIDE_IMAGE_UNRESOLVED: Attach the named source images before drafting.");
                if (draft.Cards.Count > 0 && (draft.Layout == "roadmap" || draft.Layout == "stack" || draft.Layout == "cards" || draft.Layout == "scorecard" || draft.Layout == "action_list"))
                {
                    if (queue.Count > 0) throw new InvalidOperationException("Use two_pane or visual_grid to combine card commentary with data.");
                    AddStructuredCards(elements, draft, regions[0]);
                    if (draft.Bullets.Count > 0) throw new InvalidOperationException("Move supporting bullets into the structured card points so nothing is omitted.");
                }
                else
                {
                    var commentary = string.Join("\n", draft.Bullets.Select(b => b.Text).Concat(draft.Cards.Select(c => c.Heading + "\n" + string.Join("\n", c.Points))));
                    if (draft.Layout == "two_pane" && draft.Cards.Count > 0)
                        commentary = string.Join("\n", draft.Bullets.Select(b => b.Text).Concat(draft.Cards[0].Points).Concat(draft.Cards.Skip(1).Select(c => c.Heading + "\n" + string.Join("\n", c.Points))));
                    if (commentary.Length > 0)
                    {
                        var comment = TextElement(commentary, RectangleF.Empty);
                        if (draft.Layout == "two_pane") comment.Fill = SamsungSlideDesign.Gray;
                        if (draft.Layout == "two_pane") queue = new Queue<SamsungElement>(new[] { comment }.Concat(queue));
                        else queue.Enqueue(comment);
                    }
                    if (queue.Count == 0) throw new InvalidOperationException("A content slide needs source-backed content.");
                    if (draft.Layout == "two_pane" && commentary.Length == 0 &&
                        table != null && draft.Chart != null &&
                        secondaryTable == null && draft.SecondaryChart == null &&
                        draft.ImageData.Count == 0)
                        regions = SamsungSlideDesign.ChartTableRegions();
                    if (queue.Count > regions.Length)
                    {
                        // Generic composition: preserve every supplied block. Explicit
                        // complex recipes require the matching number of regions.
                        if (draft.Layout == "table" || draft.Layout == "chart" || draft.Layout == "bullets" || draft.Layout == "large_visual" || draft.Layout == "landscape")
                            regions = SamsungSlideDesign.Regions(queue.Count > 2 ? "visual_grid" : "dual_visual");
                        if (queue.Count > regions.Length) throw new InvalidOperationException("Too many blocks for this layout; split the slide or select visual_grid.");
                    }
                    for (var i = 0; queue.Count > 0; i++)
                    {
                        var element = queue.Dequeue(); element.Box = regions[i];
                        if (draft.Layout == "two_pane" && i == 0 && element.Text.Length > 0 && draft.Cards.Count > 0)
                        {
                            var headingBox = new RectangleF(element.Box.X, element.Box.Y, element.Box.Width, 32.4f);
                            elements.Add(TextElement(draft.Cards[0].Heading, headingBox, 14, 11, "Arial", true, SamsungSlideDesign.Blue, "#FFFFFF"));
                            element.Box = new RectangleF(element.Box.X, element.Box.Y + 32.4f, element.Box.Width, element.Box.Height - 32.4f);
                        }
                        if (element.Table != null &&
                            (draft.Layout == "table" || draft.Layout == "matrix") &&
                            element.Table.Rows.SelectMany(r => r).Concat(element.Table.Headers).All(t => t.Length <= 30))
                        {
                            // A short executive table should read as the visual, not
                            // collapse into a spreadsheet strip surrounded by empty canvas.
                            var desiredHeight = Math.Min(element.Box.Height,
                                Math.Max(168f, (element.Table.Rows.Count + 1) * 45f));
                            element.Box = new RectangleF(element.Box.X,
                                element.Box.Y + (element.Box.Height - desiredHeight) / 2f,
                                element.Box.Width, desiredHeight);
                        }
                        elements.Add(element);
                    }
                }
                var audienceCaption = SamsungAuthoringPolicy.AudienceNote(draft.Caption);
                if (audienceCaption.Length > 0) elements.Add(TextElement(audienceCaption, SamsungSlideDesign.Percent(15.6f, 21f, 64.2f, 3.5f), 14, 14, "Arial Narrow"));
                var audienceUnit = SamsungAuthoringPolicy.AudienceUnit(draft.Unit);
                if (audienceUnit.Length > 0) { var unit = TextElement(audienceUnit, SamsungSlideDesign.Percent(80, 21.5f, 16.2f, 3.1f), 11, 11, "Calibri"); unit.Alignment = 3; elements.Add(unit); }
                var audienceTakeaway = SamsungAuthoringPolicy.AudienceTakeaway(draft.Takeaway);
                if (audienceTakeaway.Length > 0) elements.Add(TextElement(audienceTakeaway,
                    draft.Layout == "scorecard" ? SamsungSlideDesign.ScorecardTakeaway : SamsungSlideDesign.Takeaway,
                    14, 14, "Arial Narrow", true, SamsungSlideDesign.Blue, "#FFFFFF"));
                AddSamsungAnnotations(elements, draft, table, secondaryTable, part, perPage);
                // Semantic row references, never model-supplied coordinates.
                var primary = elements.FirstOrDefault(e => e.Table != null || e.Chart != null);
                foreach (var originalRow in draft.HighlightRows)
                {
                    var row = primary?.Table != null ? originalRow - part * perPage : originalRow;
                    if (primary == null) throw new InvalidOperationException("Highlights require a table or chart.");
                    var total = primary.Table != null ? primary.Table.Rows.Count : primary.Chart.Categories.Count;
                    if (row < 1 || row > total) continue;
                    RectangleF box;
                    if (primary.Table != null)
                    { var h = primary.Box.Height / (total + 1); box = new RectangleF(primary.Box.X, primary.Box.Y + row * h, primary.Box.Width, h); }
                    else
                    { var w = primary.Box.Width * .8f / total; box = new RectangleF(primary.Box.X + primary.Box.Width * .15f + (row - 1) * w, primary.Box.Y + primary.Box.Height * .2f, w, primary.Box.Height * .6f); }
                    elements.Add(new SamsungElement { Box = box, Hollow = true, HighlightChart = primary.Chart,
                        HighlightCategory = row, HighlightSeries = primary.Chart == null ? 0 : 1 });
                }
                foreach (var data in elements.Where(e => e.Table != null).ToArray())
                {
                    var rows = data.Table.Rows; var columns = Math.Max(1, data.Table.Headers.Count);
                    for (var r = 0; r < rows.Count; r++)
                    for (var c = 0; c < rows[r].Count; c++)
                    {
                        // An explicit semantic status is safe to render. Never infer
                        // that an increase/decrease is intrinsically good or bad.
                        var value = rows[r][c].Trim().ToLowerInvariant();
                        if (value != "strong" && value != "weak" && value != "neutral") continue;
                        var cellW = data.Box.Width / columns; var cellH = data.Box.Height / (rows.Count + 1);
                        var statusBox = new RectangleF(data.Box.X + (c + 1) * cellW - 8, data.Box.Y + (r + 1) * cellH + 2, 6, 6);
                        elements.Add(new SamsungElement { Box = statusBox, Circle = true, Fill = value == "strong" ? SamsungSlideDesign.Green : value == "weak" ? SamsungSlideDesign.Red : "#7F7F7F" });
                    }
                }
                if (draft.HighlightRows.Count > 0 && draft.Takeaway.Length > 0 && draft.Layout == "annotated_chart")
                    elements.Add(new SamsungElement { Box = SamsungSlideDesign.Percent(50.8f, 80.2f, 5.3f, 4.5f), Connector = true });
            }
            var audienceFootnote = SamsungAuthoringPolicy.AudienceNote(draft.Footnote);
            var source = string.Join("; ", new[] { audienceFootnote,
                VisibleSourceCaption(draft.Sources) }.Where(s =>
                !string.IsNullOrWhiteSpace(s)));
            // The complete citation always reaches the speaker notes. A cover or
            // divider keeps only a short visible reference, placed clear of the
            // cover's accent bar instead of across it. When the combined note
            // is long, the actual citation takes priority over optional caveats.
            var sparse = draft.Layout == "cover" || draft.Layout == "divider" || draft.Layout == "closing";
            var visibleSource = source;
            if (source.Length > (sparse ? 90 : 120))
            {
                var shortReference = string.IsNullOrWhiteSpace(draft.Sources) ?
                    audienceFootnote : VisibleSourceCaption(draft.Sources);
                shortReference = shortReference.Split(';')[0].Trim();
                var limit = sparse ? 90 : 120;
                if (shortReference.Length > limit)
                    shortReference = shortReference.Substring(0, limit - 1).TrimEnd() + "…";
                visibleSource = shortReference;
            }
            if (source.Length > 0) elements.Add(TextElement(visibleSource,
                draft.Layout == "cover" ? SamsungSlideDesign.Percent(3.8f, 90.8f, 87f, 3f) : SamsungSlideDesign.Footer, 10, 9, "Arial Narrow"));
            var pageNumber = TextElement("- " + index + " -", SamsungSlideDesign.Page, 10.5f, 8, "Calibri"); pageNumber.Alignment = 3; elements.Add(pageNumber); page.PageNumber = pageNumber;
            elements.Add(TextElement(DraftMarker, SamsungSlideDesign.Percent(3.8f, 97, 32, 2.8f), 7, 7, "Arial", false, null, "#7F7F7F"));
            if (draft.Layout == "closing")
                foreach (var element in elements) if (element.Fill == null) element.Color = "#FFFFFF";
            return page;
        }

        private static string VisibleSourceCaption(string exactSources)
        {
            if (string.IsNullOrWhiteSpace(exactSources)) return string.Empty;
            var references = exactSources.Split(';').Select(value =>
                value.Trim()).Where(value => value.Length > 0).ToArray();
            var parsed = references.Select(value => Regex.Match(value,
                @"^(?<id>excel:[0-9a-f]{32}) / (?<location>.+)$",
                RegexOptions.IgnoreCase)).ToArray();
            if (parsed.Length == 0 || parsed.Any(match => !match.Success))
                return exactSources;
            var workbooks = parsed.Select(match => match.Groups["id"].Value)
                .Distinct(StringComparer.OrdinalIgnoreCase).Count();
            if (workbooks > 1)
                return "Sources: " + workbooks +
                    " workbooks (exact references in notes)";
            if (parsed.Length > 3)
                return "Source: " + parsed[0].Groups["location"].Value
                    .Split('!')[0] + " (" + parsed.Length +
                    " references; details in notes)";
            return "Source: " + string.Join("; ", parsed.Select(match =>
                match.Groups["location"].Value));
        }

        private static void AddStructuredCards(List<SamsungElement> elements, DraftSlide draft, RectangleF region)
        {
            var count = draft.Cards.Count;
            if (draft.Layout == "scorecard" && count == 2 &&
                draft.Cards.All(card => card.Points.Count == 2 &&
                    card.Heading.Length <= 32 && card.Points[1].Length <= 46 &&
                    Regex.IsMatch(card.Points[0],
                        @"^(?:(?:EUR|USD|GBP|AED)\s+)?[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?%?$")))
            {
                // Keep both sourced measures and their contexts in one
                // editable composition. The source order conveys no rank.
                var panel = new RectangleF(region.X, region.Y + 10f,
                    region.Width, region.Height - 20f);
                elements.Add(TextElement("", panel,
                    fill: SamsungSlideDesign.Navy));
                elements.Add(TextElement("", new RectangleF(panel.X,
                    panel.Y, panel.Width, 5f),
                    fill: SamsungSlideDesign.SoftBlue));
                var columnWidth = panel.Width / 2f;
                elements.Add(TextElement("", new RectangleF(
                    panel.X + columnWidth - .5f, panel.Y + 30f, 1f,
                    panel.Height - 60f), fill: "#596674"));
                for (var index = 0; index < 2; index++)
                {
                    var card = draft.Cards[index];
                    var left = panel.X + index * columnWidth + 28f;
                    var width = columnWidth - 56f;
                    elements.Add(TextElement(card.Heading.ToUpperInvariant(),
                        new RectangleF(left, panel.Y + 28f, width, 38f),
                        18, 16, "Arial", true, null, "#B8D8FF"));
                    elements.Add(TextElement(card.Points[0],
                        new RectangleF(left, panel.Y + 76f, width, 81f),
                        56, 38, MetoTheme.TitleFont, true, null,
                        "#FFFFFF"));
                    elements.Add(TextElement(card.Points[1],
                        new RectangleF(left, panel.Y + 180f, width, 32f),
                        17, 15, "Arial", false, null, "#B8D8FF"));
                }
                return;
            }
            if (draft.Layout == "cards" && count == 2 &&
                draft.Cards.All(card => card.Points.Count == 1 &&
                    Regex.IsMatch(card.Points[0],
                        @"^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?%?$") &&
                    Regex.IsMatch(card.Heading,
                        @"\b(?:revenue|cost|margin|sales|budget|profit|headcount|rate|total)\b",
                        RegexOptions.IgnoreCase)))
            {
                // Two verified numeric facts need a single visual system.
                // Their plan labels and values remain exact native text; no
                // inferred delta, rank, or symbol is added.
                var panel = new RectangleF(region.X, region.Y + 10f,
                    region.Width, region.Height - 20f);
                elements.Add(TextElement("", panel,
                    fill: SamsungSlideDesign.Navy));
                elements.Add(TextElement("", new RectangleF(panel.X,
                    panel.Y, panel.Width, 5f),
                    fill: SamsungSlideDesign.SoftBlue));
                var columnWidth = panel.Width / 2f;
                elements.Add(TextElement("", new RectangleF(
                    panel.X + columnWidth - .5f, panel.Y + 37f, 1f,
                    panel.Height - 74f), fill: "#596674"));
                for (var index = 0; index < 2; index++)
                {
                    var card = draft.Cards[index];
                    var left = panel.X + index * columnWidth + 28f;
                    var width = columnWidth - 56f;
                    elements.Add(TextElement(
                        card.Heading.ToUpperInvariant(),
                        new RectangleF(left, panel.Y + 38f, width, 40f),
                        19, 16, "Arial", true, null, "#B8D8FF"));
                    elements.Add(TextElement(card.Points[0],
                        new RectangleF(left, panel.Y + 111f, width, 96f),
                        58, 38, MetoTheme.TitleFont, true, null,
                        "#FFFFFF"));
                    elements.Add(TextElement("", new RectangleF(left,
                        panel.Bottom - 41f, Math.Min(150f, width / 2f),
                        4f), fill: SamsungSlideDesign.SoftBlue));
                }
                return;
            }
            if (draft.Layout == "cards" && count == 1 &&
                draft.Cards[0].Points.Count == 1 &&
                draft.Cards[0].Points[0].Length <= 100 &&
                !Regex.IsMatch(draft.Cards[0].Points[0], @"\d"))
            {
                // Give a short, sourced qualitative statement a deliberate
                // focal panel rather than leaving it stranded in empty space.
                var card = draft.Cards[0];
                var panel = new RectangleF(region.X + 14f,
                    region.Y + 8f, region.Width - 28f,
                    region.Height - 16f);
                elements.Add(TextElement("", panel,
                    fill: SamsungSlideDesign.Navy));
                elements.Add(TextElement("", new RectangleF(panel.X,
                    panel.Y, panel.Width, 6f),
                    fill: SamsungSlideDesign.SoftBlue));
                elements.Add(TextElement(card.Heading.ToUpperInvariant(),
                    new RectangleF(panel.X + 34f, panel.Y + 31f,
                        panel.Width - 68f, 30f), 16, 16, "Arial", true,
                    null, "#B8D8FF"));
                elements.Add(TextElement(card.Points[0],
                    new RectangleF(panel.X + 34f, panel.Y + 84f,
                        panel.Width - 68f, panel.Height - 123f), 36, 26,
                    MetoTheme.TitleFont, true, null, "#FFFFFF"));
                elements.Add(TextElement("", new RectangleF(panel.X + 34f,
                    panel.Bottom - 34f, Math.Min(170f,
                        panel.Width / 4f), 4f),
                    fill: SamsungSlideDesign.SoftBlue));
                return;
            }
            // When a factual card grid carries several numbers, reserve the
            // lower part of its panels for source-backed metric anchors. This
            // creates hierarchy without asking the model to invent artwork or
            // weakening the exact evidence/citation checks.
            var evidenceHeroes = new string[count];
            var secondaryHero = new string[count];
            var heroLabels = new string[count, 2];
            if (draft.Layout == "cards" && count > 1 && count <= 4)
            {
                var used = new HashSet<string>(StringComparer.Ordinal);
                var columns = count == 4 ? 2 : count;
                var width = (region.Width - 26f * (columns - 1)) / columns - 28f;
                for (var cardIndex = 0; cardIndex < count; cardIndex++)
                {
                    var body = string.Join("\n", draft.Cards[cardIndex].Points);
                    if (MeasureEvidenceBody(body, width) > region.Height - 160f) continue;
                    var tokens = draft.Cards[cardIndex].Points.SelectMany(point =>
                            Regex.Matches(point,
                                @"(?<![A-Za-z0-9])(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?%?(?![A-Za-z0-9,])")
                                .Cast<Match>().Select(match => new { point, match }))
                        .Where(candidate =>
                        {
                            int year;
                            var value = candidate.match.Value;
                            return !Regex.IsMatch(value, @"^0\d{2,}$") &&
                                (!int.TryParse(value, out year) || year < 1900 || year > 2100) &&
                                IsExplicitCardMetric(candidate.point, candidate.match);
                        })
                        .Where(candidate => used.Add(candidate.match.Value)).Take(2).ToArray();
                    for (var tokenIndex = 0; tokenIndex < tokens.Length; tokenIndex++)
                    {
                        if (tokenIndex == 0) evidenceHeroes[cardIndex] = tokens[tokenIndex].match.Value;
                        else secondaryHero[cardIndex] = tokens[tokenIndex].match.Value;
                        var following = tokens[tokenIndex].point.Substring(
                            tokens[tokenIndex].match.Index + tokens[tokenIndex].match.Length);
                        var label = Regex.Match(following, @"^\s+(?<word>[A-Za-z]+)");
                        var preceding = tokens[tokenIndex].point.Substring(
                            tokens[tokenIndex].point.LastIndexOf(';', tokens[tokenIndex].match.Index) + 1,
                            tokens[tokenIndex].match.Index - tokens[tokenIndex].point.LastIndexOf(';', tokens[tokenIndex].match.Index) - 1)
                            .Trim().TrimEnd(':', '=');
                        var standalone = count == 4 ? StandaloneHeroLabel(
                            tokens[tokenIndex].point, tokens[tokenIndex]
                                .match.Value) : null;
                        heroLabels[cardIndex, tokenIndex] = standalone ??
                            (preceding.Length > 0 && preceding.Length <= 24
                                ? preceding : label.Success ?
                                    label.Groups["word"].Value : "");
                    }
                }
            }
            var primaryHeroCount = evidenceHeroes.Count(value => !string.IsNullOrEmpty(value));
            // A qualitative card set can put every quantified fact in one
            // panel. Keep two distinct, explicitly written metrics together
            // there instead of allowing a text-only grid to masquerade as a
            // visual scorecard. Never promote a year or cell-range fragment.
            var dualHeroIndex = primaryHeroCount < 2
                ? Array.FindIndex(secondaryHero, value => !string.IsNullOrEmpty(value)) : -1;
            var heroMode = primaryHeroCount >= 2 || dualHeroIndex >= 0;
            for (var i = 0; i < count; i++)
            {
                var card = draft.Cards[i];
                if (draft.Layout == "scorecard")
                {
                    var metricGap = 18f;
                    var metricWidth = (region.Width - metricGap * (count - 1)) / count;
                    var detailLines = draft.Cards.Max(c => Math.Max(0, c.Points.Count - 1));
                    var metricHeight = Math.Min(region.Height,
                        Math.Max(196f, 164f + 24f * Math.Min(3, detailLines)));
                    var metricBox = new RectangleF(
                        region.X + i * (metricWidth + metricGap),
                        region.Y + (region.Height - metricHeight) / 2f,
                        metricWidth, metricHeight);
                    var value = card.Points.FirstOrDefault() ?? "";
                    var detail = string.Join("\n", card.Points.Skip(1));
                    var focal = i == 0;
                    elements.Add(TextElement("", metricBox, fill: focal ?
                        SamsungSlideDesign.Blue : SamsungSlideDesign.Gray));
                    elements.Add(TextElement("", new RectangleF(metricBox.X, metricBox.Y, metricBox.Width, 6f),
                        fill: focal ? SamsungSlideDesign.SoftBlue : SamsungSlideDesign.Blue));
                    elements.Add(TextElement(card.Heading.ToUpperInvariant(),
                        new RectangleF(metricBox.X + 14f, metricBox.Y + 16f, metricBox.Width - 28f, 42f), 14, 14, "Arial", true, null,
                        focal ? "#FFFFFF" : "#596674"));
                    elements.Add(TextElement(value,
                        new RectangleF(metricBox.X + 14f, metricBox.Y + 61f, metricBox.Width - 28f, 65f), 34, 24, MetoTheme.TitleFont, true, null,
                        focal ? "#FFFFFF" : SamsungSlideDesign.Blue));
                    if (detail.Length > 0)
                        elements.Add(TextElement(detail,
                            new RectangleF(metricBox.X + 14f, metricBox.Y + 134f, metricBox.Width - 28f, Math.Max(32f, metricBox.Height - 150f)), 15, 14, "Arial", false, null,
                            focal ? "#FFFFFF" : "#202A35"));
                    continue;
                }
                if (draft.Layout == "action_list")
                {
                    var y = 26.4f + i * 14f;
                    elements.Add(TextElement(card.Heading, SamsungSlideDesign.Percent(4.7f, y, 16.9f, 12.8f), 18, 14, "Arial", true));
                    elements.Add(TextElement(string.Join("\n", card.Points.Take(Math.Max(0, card.Points.Count - 1))), SamsungSlideDesign.Percent(23.1f, y, 54.8f, 12.8f)));
                    elements.Add(TextElement(card.Points.LastOrDefault() ?? "", SamsungSlideDesign.Percent(79.3f, y, 12.3f, 12.8f), 18, 14));
                    continue;
                }
                if (draft.Layout == "cards")
                {
                    var columns = count == 4 ? 2 : count;
                    var rows = count == 4 ? 2 : 1;
                    var compact = rows == 2;
                    var compactHero = compact && heroMode && !string.IsNullOrEmpty(evidenceHeroes[i]);
                    const float columnGap = 26f, rowGap = 18f;
                    var evidenceWidth = (region.Width - columnGap * (columns - 1)) / columns;
                    var evidenceHeight = (region.Height - rowGap * (rows - 1)) / rows;
                    // A one-row evidence grid should frame its actual content,
                    // not leave three tall, half-empty panels behind short facts.
                    // Preserve equal panel heights and center the whole row in
                    // its assigned region so the whitespace is intentional.
                    if (rows == 1 && (!heroMode || dualHeroIndex >= 0))
                    {
                        var bodyWidth = evidenceWidth - 28f;
                        var longestBody = draft.Cards.Max(c => EvidenceCardContentHeight(c, bodyWidth));
                        // A dual metric needs a lower badge rail, but short
                        // evidence still should not become a full-height gray
                        // panel merely because two facts share one card.
                        evidenceHeight = Math.Min(region.Height, Math.Max(
                            heroMode ? 218f : 178f,
                            (heroMode ? 160f : 72f) + longestBody +
                            (heroMode ? 14f : 20f)));
                    }
                    var rowInset = rows == 1 ? (region.Height - evidenceHeight) / 2f : 0f;
                    var evidenceBox = new RectangleF(
                        region.X + (i % columns) * (evidenceWidth + columnGap),
                        region.Y + rowInset + (i / columns) * (evidenceHeight + rowGap),
                        evidenceWidth, evidenceHeight);
                    elements.Add(TextElement("", evidenceBox, fill: SamsungSlideDesign.Gray));
                    elements.Add(TextElement("", new RectangleF(evidenceBox.X, evidenceBox.Y, evidenceBox.Width, 4f),
                        fill: SamsungSlideDesign.Blue));
                    elements.Add(TextElement(card.Heading,
                        new RectangleF(evidenceBox.X + 14f, evidenceBox.Y + (compact ? 10f : 16f),
                            evidenceBox.Width - (compactHero ? 166f : 28f), compact ? 29f : 38f),
                        compact ? 18 : 20, 16, MetoTheme.TitleFont, true, null, SamsungSlideDesign.Blue));
                    if (compactHero)
                    {
                        var hero = TextElement(evidenceHeroes[i],
                            new RectangleF(evidenceBox.Right - 146f, evidenceBox.Y + 7f, 132f, 38f),
                            29, 26, MetoTheme.TitleFont, true, null, SamsungSlideDesign.Blue);
                        hero.Alignment = 3;
                        elements.Add(hero);
                    }
                    // The dual metric rail repeats two numbers as large callouts.
                    // Omit simple metric-only source lines from the body and put
                    // their original labels beneath the callouts instead.
                    var bodyPoints = i == dualHeroIndex || compactHero
                        ? card.Points.Select(point => StripHeroClauses(point,
                            evidenceHeroes[i], compactHero ? null :
                                secondaryHero[i]))
                            .Where(point => point.Length > 0).ToArray()
                        : card.Points.ToArray();
                    if (compactHero && !string.IsNullOrWhiteSpace(
                        heroLabels[i, 0]))
                        bodyPoints = new[] { heroLabels[i, 0] }
                            .Concat(bodyPoints).ToArray();
                    var body = string.Join("\n", bodyPoints);
                    if (body.Length > 0 || (!compact && heroMode && !string.IsNullOrEmpty(evidenceHeroes[i])))
                    {
                        var bodyStart = compact ? 45f : 64f;
                        var bodySpace = evidenceBox.Height - (compact ? 52f : 78f);
                        var bodyWidth = evidenceBox.Width - 28f;
                        if (!compact && heroMode && !string.IsNullOrEmpty(evidenceHeroes[i]))
                        {
                            if (body.Length > 0)
                                elements.Add(TextElement(body,
                                    new RectangleF(evidenceBox.X + 14f, evidenceBox.Y + bodyStart,
                                        bodyWidth, evidenceBox.Height - 160f),
                                    16, 14, "Arial", false, null, "#202A35"));
                            if (i == dualHeroIndex)
                            {
                                var metricWidth = (bodyWidth - 8f) / 2f;
                                for (var metric = 0; metric < 2; metric++)
                                {
                                    var metricX = evidenceBox.X + 14f + metric * (metricWidth + 8f);
                                    elements.Add(TextElement(metric == 0 ? evidenceHeroes[i] : secondaryHero[i],
                                        new RectangleF(metricX, evidenceBox.Bottom - 82f, metricWidth, 43f),
                                        30, 26, MetoTheme.TitleFont, true, null, SamsungSlideDesign.Blue));
                                    if (!string.IsNullOrEmpty(heroLabels[i, metric]))
                                        elements.Add(TextElement(heroLabels[i, metric],
                                            new RectangleF(metricX, evidenceBox.Bottom - 39f, metricWidth, 24f),
                                            14, 14, "Arial", false, null, "#596674"));
                                }
                            }
                            else elements.Add(TextElement(evidenceHeroes[i],
                                new RectangleF(evidenceBox.X + 14f, evidenceBox.Bottom - 78f,
                                    bodyWidth, 60f), 32, 26, MetoTheme.TitleFont, true,
                                null, SamsungSlideDesign.Blue));
                            continue;
                        }
                        var lead = bodyPoints.FirstOrDefault() ?? "";
                        var supporting = string.Join("\n", bodyPoints.Skip(1));
                        var leadHeight = MeasureEvidenceBody(lead, bodyWidth, 18f, true) + 5f;
                        var supportingHeight = supporting.Length == 0 ? 0f :
                            MeasureEvidenceBody(supporting, bodyWidth) + 5f;
                        var contentHeight = leadHeight + (supporting.Length == 0 ? 0f : 10f + supportingHeight);
                        if (contentHeight <= bodySpace)
                        {
                            var inset = (bodySpace - contentHeight) / 2f;
                            elements.Add(TextElement(lead,
                                new RectangleF(evidenceBox.X + 14f, evidenceBox.Y + bodyStart + inset,
                                    bodyWidth, leadHeight), 18, 16, "Arial", true, null, "#202A35"));
                            if (supporting.Length > 0)
                                elements.Add(TextElement(supporting,
                                    new RectangleF(evidenceBox.X + 14f,
                                        evidenceBox.Y + bodyStart + inset + leadHeight + 10f,
                                        bodyWidth, supportingHeight), 16, 14, "Arial", false, null, "#202A35"));
                        }
                        else
                            elements.Add(TextElement(body,
                                new RectangleF(evidenceBox.X + 14f, evidenceBox.Y + bodyStart,
                                    bodyWidth, bodySpace), 16, 14, "Arial", false, null, "#202A35"));
                    }
                    continue;
                }
                var vertical = draft.Layout == "stack";
                var gap = 14f;
                var width = vertical ? region.Width : (region.Width - gap * (count - 1)) / count;
                var height = vertical ? (region.Height - gap * (count - 1)) / count : region.Height;
                var box = new RectangleF(region.X + (vertical ? 0 : i * (width + gap)), region.Y + (vertical ? i * (height + gap) : 0), width, height);
                var text = card.Heading + "\n" + string.Join("\n", card.Points);
                elements.Add(TextElement(text, box, 18, 14, "Arial", false, SamsungSlideDesign.Gray));
                if (draft.Layout == "roadmap" && i < count - 1)
                    elements.Add(new SamsungElement { Box = new RectangleF(box.Right, box.Top + box.Height / 2, gap, 1), Connector = true });
            }
        }

        private static bool IsExplicitCardMetric(string point, Match token)
        {
            var before = point.Substring(0, token.Index).Trim();
            var after = point.Substring(token.Index + token.Length);
            // A number in a row identifier or in parenthetical row detail is
            // evidence, not the KPI of the card. Require a metric noun next to
            // the number or an explicit metric label before it.
            return Regex.IsMatch(after,
                       @"^\s+(?:(?:unique|distinct|complete|valid)\s+)?(?:rows?|records?|rowids?|observations?|obs|groups?|units?|blanks?|duplicates?|points?|percent|EUR|USD)\b",
                       RegexOptions.IgnoreCase) ||
                   Regex.IsMatch(before,
                       @"\b(?:revenue|cost|margin|sales|budget|profit|headcount|rate|total)\s*(?:EUR|USD|%)?\s*[:=]?\s*$",
                       RegexOptions.IgnoreCase);
        }

        private static bool IsStandaloneHeroPoint(string point, string hero)
        { return StandaloneHeroLabel(point, hero) != null; }

        private static string StandaloneHeroLabel(string point, string hero)
        {
            if (string.IsNullOrWhiteSpace(point) ||
                string.IsNullOrWhiteSpace(hero)) return null;
            var index = point.IndexOf(hero, StringComparison.Ordinal);
            if (index < 0 || point.IndexOf(hero, index + hero.Length,
                StringComparison.Ordinal) >= 0) return null;
            var label = point.Remove(index, hero.Length).Trim()
                .TrimEnd('.', ':', '=').Trim().TrimEnd(':', '=')
                .Trim();
            return label.Length > 0 && label.Length <= 32 &&
                Regex.IsMatch(label, @"^[\p{L}\s:/%=-]+$") ?
                    label : null;
        }

        private static string StripHeroClauses(string point, string primary, string secondary)
        {
            return string.Join("; ", (point ?? "").Split(';').Select(clause =>
                StripHeroMetric(clause.Trim(), primary, secondary))
                .Where(clause => clause.Length > 0));
        }

        private static string StripHeroMetric(string clause, params string[] heroes)
        {
            foreach (var hero in heroes.Where(value => !string.IsNullOrWhiteSpace(value)))
            {
                if (IsStandaloneHeroPoint(clause, hero)) return "";
                clause = Regex.Replace(clause,
                    @"\b(?:revenue|cost|margin|sales|budget|profit|headcount|rate|total)\s*(?:EUR|USD|%)?\s*[:=]?\s*" +
                    Regex.Escape(hero) + @"(?![A-Za-z0-9])", "", RegexOptions.IgnoreCase).Trim().TrimStart(',', ':', '.', '-', '–').TrimStart();
            }
            return Regex.Replace(clause, @"^(?:across|at|is|was)\s+", "", RegexOptions.IgnoreCase).Trim();
        }

        private static float EvidenceCardContentHeight(DraftCard card, float width)
        {
            var lead = card.Points.FirstOrDefault() ?? "";
            var supporting = string.Join("\n", card.Points.Skip(1));
            return MeasureEvidenceBody(lead, width, 18f, true) + 5f +
                (supporting.Length == 0 ? 0f : 10f + MeasureEvidenceBody(supporting, width) + 5f);
        }

        private static float MeasureEvidenceBody(string body, float width, float size = 16f, bool bold = false)
        {
            if (string.IsNullOrEmpty(body)) return 0f;
            using (var bitmap = new Bitmap(1, 1))
            using (var graphics = Graphics.FromImage(bitmap))
            using (var font = new Font(SamsungSlideDesign.FontFor(body, "Arial"), size,
                bold ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Point))
            using (var format = new StringFormat(StringFormat.GenericTypographic))
            {
                graphics.PageUnit = GraphicsUnit.Point;
                return graphics.MeasureString(body, font,
                    new SizeF(Math.Max(20f, width - 4f), 2000f), format).Height;
            }
        }

        private static void FitTable(SamsungElement element)
        {
            var table = element.Table;
            var columns = Math.Max(table.Headers.Count, table.Rows.Select(r => r.Count).DefaultIfEmpty(0).Max());
            var rows = table.Rows.Count + (table.Headers.Count > 0 ? 1 : 0);
            if (columns == 0 || rows == 0) throw new InvalidOperationException("SLIDE_TABLE_EMPTY");
            var weights = Enumerable.Range(0, columns).Select(c => (float)Math.Sqrt(Math.Max(4,
                new[] { table.Headers }.Concat(table.Rows).Where(r => r.Count > c).Select(r => r[c].Length).DefaultIfEmpty(4).Max()))).ToArray();
            var sum = weights.Sum();
            element.ColumnWidths = weights.Select(w => element.Box.Width * w / sum).ToArray();
            var size = SamsungSlideDesign.TableSize;
            foreach (var row in new[] { table.Headers }.Concat(table.Rows))
                for (var col = 0; col < row.Count; col++)
                    size = Math.Min(size, SamsungSlideDesign.Fit(row[col], "Arial Narrow", new RectangleF(0, 0, element.ColumnWidths[col] - 4, element.Box.Height / rows), SamsungSlideDesign.TableSize, SamsungSlideDesign.TableMinimum));
            element.Size = size; element.Minimum = SamsungSlideDesign.TableMinimum;
        }

        internal static void ScaleSamsungPage(SamsungPage page, float scale)
        {
            foreach (var element in page.Elements)
            {
                var box = element.Box;
                element.Box = new RectangleF(box.X * scale, box.Y * scale, box.Width * scale, box.Height * scale);
                element.Size *= scale; element.Minimum *= scale;
                if (element.ColumnWidths != null) element.ColumnWidths = element.ColumnWidths.Select(w => w * scale).ToArray();
            }
        }

        private static void AddSamsungAnnotations(List<SamsungElement> elements, DraftSlide draft, DraftTable primary, DraftTable secondary, int part, int perPage)
        {
            if (draft.Annotations.Length > 24) throw new InvalidOperationException("Too many evidence annotations.");
            foreach (var raw in draft.Annotations)
            {
                var annotation = SamsungAuthoringPolicy.ReadMap(raw);
                var target = SamsungAuthoringPolicy.Text(annotation, "target");
                var original = target == "table" ? draft.Table : target == "secondary_table" ? draft.SecondaryTable : null;
                var chart = target == "chart" ? draft.Chart : target == "secondary_chart" ? draft.SecondaryChart : null;
                var displayed = target == "table" ? primary : secondary;
                var row = Convert.ToInt32(annotation["row"]);
                var count = original != null ? original.Rows.Count : chart?.Categories.Count ?? 0;
                if (row < 1 || row > count) throw new InvalidOperationException("SLIDE_ANNOTATION_INVALID: Missing evidence target.");
                var element = elements.FirstOrDefault(e => original != null ? e.Table == displayed : e.Chart == chart);
                if (element == null) throw new InvalidOperationException("SLIDE_ANNOTATION_INVALID: Target is not displayed.");
                if (original != null)
                {
                    var local = row - part * perPage;
                    if (local < 1 || local > displayed.Rows.Count) continue;
                    var columns = displayed.Headers.Count > 0 ? displayed.Headers.Count : displayed.Rows[0].Count;
                    var offset = displayed.Headers.Count > 0 ? 1 : 0;
                    var h = element.Box.Height / (displayed.Rows.Count + offset);
                    var box = new RectangleF(element.Box.X, element.Box.Y + (local - 1 + offset) * h, element.Box.Width, h);
                    if (annotation.ContainsKey("column"))
                    {
                        var col = Convert.ToInt32(annotation["column"]);
                        if (col < 1 || col > columns) throw new InvalidOperationException("SLIDE_ANNOTATION_INVALID: Missing table column.");
                        FitTable(element);
                        box.X += element.ColumnWidths.Take(col - 1).Sum(); box.Width = element.ColumnWidths[col - 1];
                    }
                    elements.Add(new SamsungElement { Box = box, Hollow = true });
                }
                else
                {
                    if (annotation.ContainsKey("series"))
                    {
                        var series = Convert.ToInt32(annotation["series"]);
                        if (series < 1 || series > chart.Series.Count) throw new InvalidOperationException("SLIDE_ANNOTATION_INVALID: Missing chart series.");
                        elements.Add(new SamsungElement { Box = element.Box, Hollow = true, HighlightChart = chart, HighlightCategory = row, HighlightSeries = series });
                    }
                    else elements.Add(new SamsungElement { Box = element.Box, Hollow = true,
                        HighlightChart = chart, HighlightCategory = row, HighlightSeries = 1 });
                }
            }
        }

        // Adds one blank slide and draws the page on it. If drawing fails, the
        // slide this call just created (never receipted, reviewed or shown as
        // complete) is removed again, so the deck returns to its last receipted
        // state and the same payload can be retried. Without this a single
        // native failure left an unreceipted slide that made every retry end
        // in SLIDE_RECOVERY_UNCERTAIN.
        internal static SamsungOutput DrawNewSamsungSlide(object nativeSlidesObject, int index, SamsungPage page, string owner,
            bool skipChartPreview = false)
        {
            dynamic nativeSlides = nativeSlidesObject;
            dynamic created = nativeSlides.Add(index, PpLayoutBlank);
            try
            {
                var output = DrawSamsungPage((object)created, page, owner);
                output.Image = skipChartPreview &&
                    PresentationInspection.ContainsNativeChart((object)created)
                        ? null : ExportSamsung(output);
                return output;
            }
            catch
            {
                try { created.Delete(); } catch (Exception) { }
                throw;
            }
        }

        internal static SamsungOutput DrawSamsungPage(object slideObject, SamsungPage page, string owner, int? displaySlideNumber = null)
        {
            dynamic slide = slideObject;
            page.PageNumber.Text = "- " + (displaySlideNumber ?? (int)slide.SlideIndex) + " -";
            slide.Tags.Add("ScribbleTask", owner);
            PaintBackground(slide, MetoTheme.Rgb(page.Background));
            var output = new SamsungOutput { Slide = slideObject, Page = page, Owner = owner };
            var chartIndices = new Dictionary<DraftChart, int>();
            foreach (var element in page.Elements)
            {
                var box = element.Box;
                if (element.HighlightChart != null)
                {
                    dynamic chartShape = slide.Shapes[chartIndices[element.HighlightChart]];
                    dynamic plot = chartShape.Chart.PlotArea;
                    var area = new RectangleF((float)chartShape.Left + (float)plot.InsideLeft, (float)chartShape.Top + (float)plot.InsideTop, (float)plot.InsideWidth, (float)plot.InsideHeight);
                    var count = element.HighlightChart.Categories.Count;
                    if (element.HighlightChart.TypeCode == DraftChartTypes.BarClustered || element.HighlightChart.TypeCode == DraftChartTypes.BarStacked)
                    {
                        var reverse = (bool)chartShape.Chart.Axes(1).ReversePlotOrder;
                        var slot = reverse ? element.HighlightCategory - 1 : count - element.HighlightCategory;
                        box = new RectangleF(area.X, area.Y + slot * area.Height / count, area.Width, area.Height / count);
                    }
                    else box = new RectangleF(area.X + (element.HighlightCategory - 1) * area.Width / count, area.Y, area.Width / count, area.Height);
                    if (element.HighlightSeries > 0)
                    {
                        dynamic point = chartShape.Chart.SeriesCollection(element.HighlightSeries).Points(element.HighlightCategory);
                        box = new RectangleF((float)chartShape.Left + (float)point.Left, (float)chartShape.Top + (float)point.Top, Math.Max(4, (float)point.Width), Math.Max(4, (float)point.Height));
                    }
                    dynamic canvas = slide.Parent;
                    if (box.Left < 0 || box.Top < 0 || box.Right > (float)canvas.PageSetup.SlideWidth || box.Bottom > (float)canvas.PageSetup.SlideHeight) throw new InvalidOperationException("SLIDE_CHART_GEOMETRY_INVALID");
                    element.Box = box;
                    ReleaseSamsungCom((object)chartShape);
                }
                dynamic shape;
                if (element.ImageData != null)
                {
                    var path = Path.Combine(Path.GetTempPath(), "scribble-source-" + Guid.NewGuid().ToString("N") + ".png");
                    try
                    {
                        var bytes = Convert.FromBase64String(element.ImageData.Substring(element.ImageData.IndexOf(',') + 1));
                        using (var stream = new MemoryStream(bytes))
                        using (var image = Image.FromStream(stream))
                        {
                            image.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                            var ratio = Math.Min(box.Width / image.Width, box.Height / image.Height);
                            var width = image.Width * ratio; var height = image.Height * ratio;
                            shape = slide.Shapes.AddPicture(path, 0, -1, box.X + (box.Width - width) / 2, box.Y + (box.Height - height) / 2, width, height);
                        }
                    }
                    finally { if (File.Exists(path)) File.Delete(path); }
                }
                else if (element.Connector)
                {
                    shape = slide.Shapes.AddConnector(2, box.Left, box.Top, box.Right, box.Bottom);
                    shape.Line.Weight = .5f; shape.Line.ForeColor.RGB = MetoTheme.Rgb("#7F7F7F");
                    shape.Line.EndArrowheadStyle = 3;
                }
                else if (element.Chart != null)
                {
                    var beforeChartShapes = (int)slide.Shapes.Count;
                    var chartCreated = false;
                    for (var attempt = 0; attempt < 3; attempt++)
                    {
                        if (AddChartToSlide(slide, element.Chart, box.X, box.Y, box.Width, box.Height))
                        { chartCreated = true; break; }
                        if (attempt == 2 || !RetryableSamsungChartFailure(LastChartFailure)) break;
                        // AddChart2 may have left a partial native shape. Only
                        // remove shapes added by this failed attempt, never a
                        // pre-existing slide shape or another user's content.
                        while ((int)slide.Shapes.Count > beforeChartShapes)
                            slide.Shapes[(int)slide.Shapes.Count].Delete();
                        System.Threading.Thread.Sleep(350 * (attempt + 1));
                    }
                    if (!chartCreated)
                        throw new InvalidOperationException("SLIDE_CHART_FAILED: Native chart could not be created; the draft remains incomplete. Host step " +
                            (LastChartFailure ?? "unknown") + ".");
                    shape = slide.Shapes[slide.Shapes.Count];
                    chartIndices[element.Chart] = (int)slide.Shapes.Count;
                }
                else if (element.Table != null)
                {
                    var rows = new[] { element.Table.Headers }.Where(r => r.Count > 0).Concat(element.Table.Rows).ToArray();
                    var columns = rows.Max(r => r.Count);
                    shape = slide.Shapes.AddTable(rows.Length, columns, box.X, box.Y, box.Width, box.Height);
                    dynamic table = shape.Table;
                    if (element.ColumnWidths != null)
                        for (var col = 0; col < columns; col++) table.Columns[col + 1].Width = element.ColumnWidths[col];
                    for (var row = 0; row < rows.Length; row++)
                    for (var col = 0; col < columns; col++)
                    {
                        dynamic cell = table.Cell(row + 1, col + 1);
                        dynamic cellShape = cell.Shape;
                        var aggregateRow = row > 0 && rows[row].Count > 0 &&
                            Regex.IsMatch(rows[row][0] ?? "",
                                @"(?i)^\s*(?:all groups|grand total|total)\s*$");
                        cellShape.Fill.Solid(); cellShape.Fill.ForeColor.RGB = MetoTheme.Rgb(
                            row == 0 ? SamsungSlideDesign.Blue : aggregateRow ? "#D7DDE3" :
                            row % 2 == 0 ? SamsungSlideDesign.Gray : "#FFFFFF");
                        for (var edge = 1; edge <= 4; edge++) { cell.Borders(edge).Weight = .5f; cell.Borders(edge).ForeColor.RGB = MetoTheme.Rgb("#A6A6A6"); }
                        var cellText = TextElement(col < rows[row].Count ? rows[row][col] : "",
                            new RectangleF(0, 0, element.ColumnWidths == null ? box.Width / columns : element.ColumnWidths[col], box.Height / rows.Length),
                            element.Size, element.Minimum, columns <= 4 ? "Arial" : "Arial Narrow",
                            row == 0 || aggregateRow, null, row == 0 ? "#FFFFFF" : "#202A35");
                        if (row > 0 && col > 0 && Regex.IsMatch(cellText.Text,
                            @"^\s*(?:[-+\u2212]?\d|(?:EUR|€)\s*[-+\u2212]?\d)"))
                            cellText.Alignment = 3;
                        ApplySamsungText(cellShape, cellText);
                        try { cellShape.TextFrame.VerticalAnchor = 3; }
                        catch (Exception exception) when (IsUnsupportedFrameSetting(exception)) { }
                    }
                    // New rows start at PowerPoint's default height for 18pt
                    // text, nearly twice the planned box. With the table font
                    // applied, return each row to its planned share so the
                    // table cannot run into the block below it; PowerPoint
                    // still enforces the minimum its text needs.
                    for (var row = 0; row < rows.Length; row++)
                    {
                        try { table.Rows[row + 1].Height = box.Height / rows.Length; }
                        catch (Exception exception) when (IsUnsupportedFrameSetting(exception)) { }
                    }
                }
                else
                {
                    shape = element.Fill != null || element.Hollow
                        ? slide.Shapes.AddShape(element.Circle ? 9 : 1, box.X, box.Y, box.Width, box.Height)
                        : slide.Shapes.AddTextbox(1, box.X, box.Y, box.Width, box.Height);
                    shape.Line.Visible = element.Hollow ? -1 : 0;
                    if (element.Hollow) { shape.Fill.Visible = 0; shape.Line.ForeColor.RGB = MetoTheme.Rgb(SamsungSlideDesign.Blue); shape.Line.Weight = 1f; }
                    else
                    {
                        if (element.Fill != null) { shape.Fill.Solid(); shape.Fill.ForeColor.RGB = MetoTheme.Rgb(element.Fill); }
                        ApplySamsungText(shape, element);
                    }
                }
                shape.Tags.Add("ScribbleTask", owner);
                output.ShapeIds.Add((int)shape.Id);
                ReleaseSamsungCom((object)shape);
            }
            // The complete source references stay in notes even when the visible
            // footer is short. No presentation is saved by this writer.
            dynamic notes = slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange;
            var existingNotes = Convert.ToString(notes.Text) ?? "";
            notes.InsertAfter((existingNotes.Length > 0 ? "\n\n" : "") + page.Source.Sources + "\n" + page.Source.Footnote + "\nEvidence:\n" + page.Source.Evidence);
            return output;
        }
        private static bool IsUnsupportedFrameSetting(Exception exception)
        {
            return exception is System.Runtime.InteropServices.COMException ||
                   exception is ArgumentException ||
                   exception is System.Reflection.TargetInvocationException;
        }

        private static void ApplySamsungText(dynamic shape, SamsungElement element)
        {
            dynamic frame = shape.TextFrame;
            // A native table cell owns its wrapping and sizing: PowerPoint
            // rejects AutoSize and WordWrap there with "The specified value is
            // out of range", which used to abort every slide that carried a
            // table. They only matter for free text boxes, so each is applied
            // where the shape accepts it.
            try { frame.AutoSize = 0; } catch (Exception exception) when (IsUnsupportedFrameSetting(exception)) { }
            try { frame.WordWrap = -1; } catch (Exception exception) when (IsUnsupportedFrameSetting(exception)) { }
            frame.MarginLeft = 2f; frame.MarginRight = 2f; frame.MarginTop = 1f; frame.MarginBottom = 1f;
            try { shape.TextFrame2.AutoSize = 0; } catch (Exception exception) when (IsUnsupportedFrameSetting(exception)) { }
            dynamic range = frame.TextRange;
            range.Text = element.Text;
            if (element.Text.Length == 0) return;
            range.Font.Name = SamsungSlideDesign.FontFor(element.Text, element.Font);
            range.Font.Size = element.Size; range.Font.Bold = element.Bold ? -1 : 0;
            range.Font.Color.RGB = MetoTheme.Rgb(element.Color);
            range.ParagraphFormat.SpaceAfter = 0; range.ParagraphFormat.SpaceBefore = 0;
            range.ParagraphFormat.Alignment = element.Alignment;
            range.ParagraphFormat.LineRuleWithin = -1; range.ParagraphFormat.SpaceWithin = 1f;
            // Bounded native repair. Never let AutoFit shrink below design minima.
            while (((double)range.BoundHeight > element.Box.Height - 2 || (double)range.BoundWidth > element.Box.Width - 3) && (float)range.Font.Size > element.Minimum)
                range.Font.Size = Math.Max(element.Minimum, (float)range.Font.Size - .5f);
            if ((double)range.BoundHeight > element.Box.Height || (double)range.BoundWidth > element.Box.Width)
                throw new InvalidOperationException("SLIDE_OVERFLOW: PowerPoint text metrics require splitting this content: " +
                    TextBoundary.SingleLine(element.Text, 80) + " (box " +
                    Math.Round(element.Box.Width) + " x " + Math.Round(element.Box.Height) +
                    ", minimum " + element.Minimum + " pt).");
        }
        internal static void SetSamsungPageNumber(SamsungOutput output, int index)
        {
            output.Page.PageNumber.Text = "- " + index + " -";
            var element = output.Page.Elements.IndexOf(output.Page.PageNumber);
            ApplySamsungText(((dynamic)output.Slide).Shapes[element + 1], output.Page.PageNumber);
        }
        internal static string ExportSamsung(SamsungOutput output)
        {
            dynamic slide = output.Slide;
            if (!SamsungSlideDesign.SameOwner((string)slide.Tags["ScribbleTask"], output.Owner)) throw new InvalidOperationException("SLIDE_OWNERSHIP_CHANGED");
            var path = Path.Combine(Path.GetTempPath(), "scribble-slide-" + Guid.NewGuid().ToString("N") + ".png");
            try { slide.Export(path, "PNG", 1600, 900); return "data:image/png;base64," + Convert.ToBase64String(File.ReadAllBytes(path)); }
            finally { if (File.Exists(path)) File.Delete(path); }
        }
        internal static void ReplaceOwnedSamsung(SamsungOutput output,
            SamsungPage replacement, Action afterNativeDeletion = null)
        {
            dynamic slide = output.Slide;
            if (ExportSamsung(output) != output.Image || (int)slide.Shapes.Count != output.ShapeIds.Count)
                throw new InvalidOperationException("SLIDE_CHANGED_DURING_REPAIR");
            for (var i = 0; i < output.ShapeIds.Count; i++)
                if ((int)slide.Shapes[i + 1].Id != output.ShapeIds[i] || !SamsungSlideDesign.SameOwner((string)slide.Shapes[i + 1].Tags["ScribbleTask"], output.Owner))
                    throw new InvalidOperationException("SLIDE_OWNERSHIP_CHANGED");
            dynamic originalDeck = slide.Parent;
            dynamic temporary = slide.Application.Presentations.Add(0);
            temporary.PageSetup.SlideWidth = originalDeck.PageSetup.SlideWidth;
            temporary.PageSetup.SlideHeight = originalDeck.PageSetup.SlideHeight;
            var scale = (float)originalDeck.PageSetup.SlideWidth / SamsungSlideDesign.Width;
            if (Math.Abs(scale - 1) > .001) ScaleSamsungPage(replacement, scale);
            dynamic staged = temporary.Slides.Add(1, PpLayoutBlank);
            DrawSamsungPage((object)staged, replacement, output.Owner);
            // Preserve live numbering when the staging deck contains one slide.
            replacement.PageNumber.Text = "- " + (int)slide.SlideIndex + " -";
            var numberIndex = replacement.Elements.IndexOf(replacement.PageNumber);
            ApplySamsungText(staged.Shapes[numberIndex + 1], replacement.PageNumber);
            PresentationRevision.ValidateNativeGeometry((object)staged);
            // Preserve a native original until the replacement succeeds. This is an
            // unsaved Office object and never a saved/exported presentation.
            slide.Copy();
            dynamic backupRange = temporary.Slides.Paste(2);
            dynamic backup = backupRange[1];
            try
            {
                if (ExportSamsung(output) != output.Image) throw new InvalidOperationException("SLIDE_CHANGED_DURING_REPAIR");
                for (var i = (int)slide.Shapes.Count; i >= 1; i--) slide.Shapes[i].Delete();
                // Acceptance harness fault point: the original is now only
                // available in the unsaved recovery presentation.
                afterNativeDeletion?.Invoke();
                PaintBackground(slide, MetoTheme.Rgb(replacement.Background));
                staged.Shapes.Range().Copy(); slide.Shapes.Paste();
                output.Page = replacement;
                output.ShapeIds.Clear();
                for (var i = 1; i <= (int)slide.Shapes.Count; i++) output.ShapeIds.Add((int)slide.Shapes[i].Id);
                output.Image = ExportSamsung(output);
                temporary.Close();
            }
            catch
            {
                // The native original remains available rather than falsely
                // claiming successful recovery after an uncertain COM failure.
                temporary.NewWindow();
                throw new InvalidOperationException("SLIDE_REPAIR_RECOVERY_REQUIRED: The unsaved recovery presentation contains the original slide.");
            }
        }
        private static void ReleaseSamsungCom(object value)
        { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    }
}
