using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;

namespace Scribble.Office
{
    internal static partial class PresentationDraftWriter
    {
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
                    { double number; if (!double.TryParse(Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out number) || double.IsNaN(number) || double.IsInfinity(number)) throw new InvalidOperationException("Chart values must be finite numbers from the source."); }
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
            internal float Size = 18, Minimum = 14;
            internal bool Bold, Hollow;
            internal bool Circle, Connector;
            internal int Alignment = 1;
            internal DraftTable Table;
            internal DraftChart Chart;
            internal string ImageData;
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

        internal static List<SamsungPage> ComposeSamsung(IReadOnlyList<DraftSlide> drafts)
        {
            var pages = new List<SamsungPage>();
            foreach (var draft in drafts)
            {
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
                        perPage = Math.Min(perPage, Math.Max(1, (int)(area.Height / (lines * 10 + 6)) - 1));
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
                elements.Add(TextElement(draft.Title, regions[0], draft.Layout == "cover" ? 66 : 40, 28,
                    MetoTheme.TitleFont, true, null, draft.Layout == "closing" ? "#FFFFFF" : "#000000"));
                if (draft.Subtitle.Length > 0) elements.Add(TextElement(draft.Subtitle, SamsungSlideDesign.Percent(4.6f, 78, 84.4f, 8), 22, 18));
                if (draft.Layout == "cover") elements.Add(TextElement("", SamsungSlideDesign.Percent(0, 94.7f, 100, 2.1f), fill: SamsungSlideDesign.Blue));
                if (draft.Bullets.Count > 0 || draft.Cards.Count > 0 || table != null || draft.Chart != null)
                    throw new InvalidOperationException("Cover/divider/closing accepts title and subtitle only; put supporting content on a content slide.");
            }
            else
            {
                elements.Add(TextElement(draft.Title + (parts > 1 ? " (" + (part + 1) + "/" + parts + ")" : ""), SamsungSlideDesign.Title, 24, 18, MetoTheme.TitleFont, true));
                elements.Add(TextElement(draft.Subtitle, SamsungSlideDesign.Action, 14, 14));
                var queue = new Queue<SamsungElement>();
                if (table != null) queue.Enqueue(new SamsungElement { Table = table });
                if (draft.Chart != null) queue.Enqueue(new SamsungElement { Chart = draft.Chart });
                if (secondaryTable != null) queue.Enqueue(new SamsungElement { Table = secondaryTable });
                if (draft.SecondaryChart != null) queue.Enqueue(new SamsungElement { Chart = draft.SecondaryChart });
                foreach (var image in draft.ImageData) queue.Enqueue(new SamsungElement { ImageData = image });
                if (draft.ImageNames.Count != draft.ImageData.Count) throw new InvalidOperationException("SLIDE_IMAGE_UNRESOLVED: Attach the named source images before drafting.");
                if (draft.Cards.Count > 0 && (draft.Layout == "roadmap" || draft.Layout == "stack" || draft.Layout == "cards" || draft.Layout == "action_list"))
                {
                    if (queue.Count > 0) throw new InvalidOperationException("Use two_pane or visual_grid to combine card commentary with data.");
                    AddStructuredCards(elements, draft, regions[0]);
                    if (draft.Bullets.Count > 0) throw new InvalidOperationException("Move supporting bullets into the structured card points so nothing is omitted.");
                }
                else
                {
                    var commentary = string.Join("\n", draft.Bullets.Select(b => b.Text).Concat(draft.Cards.Select(c => c.Heading + "\n" + string.Join("\n", c.Points))));
                    if (commentary.Length > 0)
                    {
                        var comment = TextElement(commentary, RectangleF.Empty);
                        if (draft.Layout == "two_pane") queue = new Queue<SamsungElement>(new[] { comment }.Concat(queue));
                        else queue.Enqueue(comment);
                    }
                    if (queue.Count == 0) throw new InvalidOperationException("A content slide needs source-backed content.");
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
                        var element = queue.Dequeue(); element.Box = regions[i]; elements.Add(element);
                    }
                }
                if (draft.Caption.Length > 0) elements.Add(TextElement(draft.Caption, SamsungSlideDesign.Percent(15.6f, 21f, 57, 3.5f), 14, 11, "Arial Narrow", true));
                if (draft.Unit.Length > 0) { var unit = TextElement(draft.Unit, SamsungSlideDesign.Percent(80, 21.5f, 16.2f, 3.1f), 8, 8, "Calibri"); unit.Alignment = 3; elements.Add(unit); }
                if (draft.Takeaway.Length > 0) elements.Add(TextElement(draft.Takeaway, SamsungSlideDesign.Takeaway, 14, 11, "Arial Narrow", true, SamsungSlideDesign.Blue, "#FFFFFF"));
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
                    elements.Add(new SamsungElement { Box = box, Hollow = true });
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
            var source = string.Join("; ", new[] { draft.Footnote, draft.Sources }.Where(s => !string.IsNullOrWhiteSpace(s)));
            if (source.Length > 0) elements.Add(TextElement(source.Length > 240 ? "Source references and evidence: see speaker notes." : source, SamsungSlideDesign.Footer, 7, 7, "Arial Narrow"));
            var pageNumber = TextElement("- " + index + " -", SamsungSlideDesign.Page, 10.5f, 8, "Calibri"); pageNumber.Alignment = 3; elements.Add(pageNumber); page.PageNumber = pageNumber;
            elements.Add(TextElement(DraftMarker, SamsungSlideDesign.Percent(3.8f, 97, 32, 2.8f), 7, 7, "Arial", false, null, "#7F7F7F"));
            return page;
        }

        private static void AddStructuredCards(List<SamsungElement> elements, DraftSlide draft, RectangleF region)
        {
            var count = draft.Cards.Count;
            for (var i = 0; i < count; i++)
            {
                var card = draft.Cards[i];
                if (draft.Layout == "action_list")
                {
                    var y = 26.4f + i * 14f;
                    elements.Add(TextElement(card.Heading, SamsungSlideDesign.Percent(4.7f, y, 16.9f, 12.8f), 18, 14, "Arial", true));
                    elements.Add(TextElement(string.Join("\n", card.Points.Take(Math.Max(0, card.Points.Count - 1))), SamsungSlideDesign.Percent(23.1f, y, 54.8f, 12.8f)));
                    elements.Add(TextElement(card.Points.LastOrDefault() ?? "", SamsungSlideDesign.Percent(79.3f, y, 12.3f, 12.8f), 18, 14));
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

        private static void FitTable(SamsungElement element)
        {
            var table = element.Table;
            var columns = Math.Max(table.Headers.Count, table.Rows.Select(r => r.Count).DefaultIfEmpty(0).Max());
            var rows = table.Rows.Count + (table.Headers.Count > 0 ? 1 : 0);
            var box = new RectangleF(0, 0, element.Box.Width / columns - 4, element.Box.Height / rows - 2);
            var size = 10f;
            foreach (var row in new[] { table.Headers }.Concat(table.Rows))
                foreach (var text in row) size = Math.Min(size, SamsungSlideDesign.Fit(text, "Arial Narrow", box, 10, 7.5f));
            element.Size = size; element.Minimum = 7.5f;
        }

        internal static SamsungOutput DrawSamsungPage(object slideObject, SamsungPage page, string owner)
        {
            dynamic slide = slideObject;
            page.PageNumber.Text = "- " + (int)slide.SlideIndex + " -";
            slide.Tags.Add("ScribbleTask", owner);
            PaintBackground(slide, MetoTheme.Rgb(page.Background));
            var output = new SamsungOutput { Slide = slideObject, Page = page, Owner = owner };
            foreach (var element in page.Elements)
            {
                var box = element.Box;
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
                    if (!AddChartToSlide(slide, element.Chart, box.X, box.Y, box.Width, box.Height))
                        throw new InvalidOperationException("SLIDE_CHART_FAILED: Native chart could not be created; the draft remains incomplete.");
                    shape = slide.Shapes[slide.Shapes.Count];
                }
                else if (element.Table != null)
                {
                    var rows = new[] { element.Table.Headers }.Where(r => r.Count > 0).Concat(element.Table.Rows).ToArray();
                    var columns = rows.Max(r => r.Count);
                    shape = slide.Shapes.AddTable(rows.Length, columns, box.X, box.Y, box.Width, box.Height);
                    dynamic table = shape.Table;
                    for (var row = 0; row < rows.Length; row++)
                    for (var col = 0; col < columns; col++)
                    {
                        dynamic cell = table.Cell(row + 1, col + 1);
                        dynamic cellShape = cell.Shape;
                        cellShape.Fill.Solid(); cellShape.Fill.ForeColor.RGB = MetoTheme.Rgb(row == 0 ? SamsungSlideDesign.Gray : "#FFFFFF");
                        for (var edge = 1; edge <= 4; edge++) { cell.Borders(edge).Weight = .5f; cell.Borders(edge).ForeColor.RGB = MetoTheme.Rgb("#A6A6A6"); }
                        ApplySamsungText(cellShape, TextElement(col < rows[row].Count ? rows[row][col] : "", new RectangleF(0, 0, box.Width / columns, box.Height / rows.Length), element.Size, 7.5f, "Arial Narrow", row == 0));
                    }
                }
                else
                {
                    shape = element.Fill != null || element.Hollow
                        ? slide.Shapes.AddShape(element.Circle ? 9 : 1, box.X, box.Y, box.Width, box.Height)
                        : slide.Shapes.AddTextbox(1, box.X, box.Y, box.Width, box.Height);
                    shape.Line.Visible = element.Hollow ? -1 : 0;
                    if (element.Hollow) { shape.Fill.Visible = 0; shape.Line.ForeColor.RGB = MetoTheme.Rgb(SamsungSlideDesign.Red); shape.Line.Weight = 1f; }
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
            slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = DraftMarker + "\n" + SamsungSlideDesign.Version + "\n" + page.Source.Sources + "\n" + page.Source.Footnote + "\nEvidence:\n" + page.Source.Evidence;
            return output;
        }
        private static void ApplySamsungText(dynamic shape, SamsungElement element)
        {
            dynamic frame = shape.TextFrame;
            frame.AutoSize = 0; frame.WordWrap = -1;
            frame.MarginLeft = 2f; frame.MarginRight = 2f; frame.MarginTop = 1f; frame.MarginBottom = 1f;
            shape.TextFrame2.AutoSize = 0;
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
                throw new InvalidOperationException("SLIDE_OVERFLOW: PowerPoint text metrics require splitting this content.");
        }
        internal static string ExportSamsung(SamsungOutput output)
        {
            dynamic slide = output.Slide;
            if (!SamsungSlideDesign.SameOwner((string)slide.Tags["ScribbleTask"], output.Owner)) throw new InvalidOperationException("SLIDE_OWNERSHIP_CHANGED");
            var path = Path.Combine(Path.GetTempPath(), "scribble-slide-" + Guid.NewGuid().ToString("N") + ".png");
            try { slide.Export(path, "PNG", 1600, 900); return "data:image/png;base64," + Convert.ToBase64String(File.ReadAllBytes(path)); }
            finally { if (File.Exists(path)) File.Delete(path); }
        }
        internal static void RepairSamsung(SamsungOutput output)
        {
            dynamic slide = output.Slide;
            if (!SamsungSlideDesign.SameOwner((string)slide.Tags["ScribbleTask"], output.Owner) || (int)slide.Shapes.Count != output.ShapeIds.Count)
                throw new InvalidOperationException("SLIDE_OWNERSHIP_CHANGED: The draft was edited during review.");
            for (var i = 0; i < output.ShapeIds.Count; i++)
            {
                dynamic shape = slide.Shapes[i + 1];
                if ((int)shape.Id != output.ShapeIds[i] || !SamsungSlideDesign.SameOwner((string)shape.Tags["ScribbleTask"], output.Owner))
                    throw new InvalidOperationException("SLIDE_OWNERSHIP_CHANGED");
                var element = output.Page.Elements[i];
                if (element.Chart == null && element.Table == null && element.ImageData == null && !element.Hollow && !element.Connector && !element.Circle)
                {
                    if (((string)shape.TextFrame.TextRange.Text).Replace("\r\n", "\n").Replace("\r", "\n") != element.Text.Replace("\r\n", "\n").Replace("\r", "\n")) throw new InvalidOperationException("SLIDE_CONTENT_CHANGED: User edits are preserved.");
                    element.Size = Math.Max(element.Minimum, element.Size - 1f);
                    ApplySamsungText(shape, element);
                }
                ReleaseSamsungCom((object)shape);
            }
        }
        private static void ReleaseSamsungCom(object value)
        { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    }
}
