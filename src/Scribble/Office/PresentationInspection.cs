using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Web.Script.Serialization;
using Scribble.Chat;

namespace Scribble.Office
{
    public static class PresentationInspection
    {
        private sealed class Identity { internal readonly string Id = Guid.NewGuid().ToString("N"); }
        private static readonly ConditionalWeakTable<object, Identity> Identities = new ConditionalWeakTable<object, Identity>();
        public static string IdentityFor(object presentation)
        {
            try { dynamic deck = presentation; string saved = deck.Tags["ScribblePresentationId"]; if (!string.IsNullOrEmpty(saved)) return saved; }
            catch (Exception e) when (e is System.Runtime.InteropServices.COMException || e is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
            return Identities.GetValue(presentation, p => new Identity()).Id;
        }
        public static object FindSlide(object presentation, int id)
        {
            dynamic deck = presentation;
            for (var i = 1; i <= (int)deck.Slides.Count; i++)
            { dynamic slide = deck.Slides[i]; if ((int)slide.SlideID == id) return slide; }
            throw new InvalidOperationException("SLIDE_TARGET_MISSING: The original slide is no longer in the bound presentation.");
        }
        public static object FindShape(object slide, int id)
        {
            dynamic page = slide;
            return FindInShapes((object)page.Shapes, id) ?? throw new InvalidOperationException("SLIDE_SHAPE_MISSING");
        }
        private static object FindInShapes(object shapes, int id)
        {
            dynamic collection = shapes;
            for (var i = 1; i <= (int)collection.Count; i++)
            {
                dynamic shape = collection[i];
                if ((int)shape.Id == id) return shape;
                if ((int)shape.Type == 6)
                { var match = FindInShapes((object)shape.GroupItems, id); if (match != null) return match; }
            }
            return null;
        }
        public static string Preview(object slide)
        {
            var path = Path.Combine(Path.GetTempPath(), "scribble-review-" + Guid.NewGuid().ToString("N") + ".png");
            try
            {
                dynamic page = slide; dynamic deck = page.Parent;
                var height = Math.Max(1, (int)Math.Round(1600.0 * (double)deck.PageSetup.SlideHeight / (double)deck.PageSetup.SlideWidth));
                page.Export(path, "PNG", 1600, height);
                return "data:image/png;base64," + Convert.ToBase64String(File.ReadAllBytes(path));
            }
            finally { if (File.Exists(path)) File.Delete(path); }
        }
        public static Dictionary<string, object> Capture(object slide)
        {
            dynamic page = slide;
            var unsupported = new List<string>();
            return new Dictionary<string, object> {
                { "slide_id", (int)page.SlideID }, { "index", (int)page.SlideIndex },
                { "shapes", Shapes((object)page.Shapes, unsupported, 0) },
                { "notes", Notes(slide) }, { "unsupported", unsupported },
                { "background", Background(slide) },
                { "hidden", (int)page.SlideShowTransition.Hidden },
                { "animation_count", (int)page.TimeLine.MainSequence.Count },
                { "hyperlinks", Hyperlinks(slide) }
            };
        }
        private static object Background(object slide)
        {
            dynamic page = slide; dynamic fill = page.Background.Fill;
            return new { inherited = (int)page.FollowMasterBackground, type = (int)fill.Type,
                color = (int)fill.ForeColor.RGB, transparency = (float)fill.Transparency };
        }
        internal static object[] Hyperlinks(object slide)
        {
            dynamic page = slide; var links = new List<object>();
            for (var i = 1; i <= (int)page.Hyperlinks.Count; i++)
            {
                dynamic link = page.Hyperlinks[i];
                links.Add(new { address = Convert.ToString(link.Address), subaddress = Convert.ToString(link.SubAddress), screen_tip = Convert.ToString(link.ScreenTip) });
            }
            return links.ToArray();
        }
        public static string Notes(object slide)
        {
            dynamic page = slide;
            var result = new List<string>();
            for (var i = 1; i <= (int)page.NotesPage.Shapes.Count; i++)
            { dynamic shape = page.NotesPage.Shapes[i]; if ((int)shape.HasTextFrame != 0) result.Add(Convert.ToString(shape.TextFrame.TextRange.Text)); }
            return string.Join("\n", result);
        }
        private static List<object> Shapes(object value, List<string> unsupported, int depth)
        {
            if (depth > 16) throw new InvalidOperationException("Slide groups exceed the inspection depth limit.");
            dynamic shapes = value;
            var result = new List<object>();
            for (var i = 1; i <= (int)shapes.Count; i++)
            {
                dynamic shape = shapes[i]; var id = (int)shape.Id;
                var data = new Dictionary<string, object> {
                    { "id", id }, { "name", Convert.ToString(shape.Name) }, { "type", (int)shape.Type },
                    { "x", (float)shape.Left }, { "y", (float)shape.Top }, { "width", (float)shape.Width }, { "height", (float)shape.Height },
                    { "rotation", (float)shape.Rotation }, { "z_order", (int)shape.ZOrderPosition }
                };
                if ((int)shape.Type == 6) data["children"] = Shapes((object)shape.GroupItems, unsupported, depth + 1);
                if ((int)shape.HasTextFrame != 0)
                {
                    dynamic range = shape.TextFrame.TextRange;
                    data["text"] = Convert.ToString(range.Text);
                    data["font"] = Convert.ToString(range.Font.Name); data["font_size"] = (float)range.Font.Size;
                    data["bold"] = (int)range.Font.Bold;
                    var runs = new List<object>();
                    for (var n = 1; n <= (int)range.Runs().Count; n++)
                    {
                        dynamic run = range.Runs(n, 1);
                        runs.Add(new { text = Convert.ToString(run.Text), font = Convert.ToString(run.Font.Name), far_east_font = Convert.ToString(run.Font.NameFarEast),
                            size = (float)run.Font.Size, bold = (int)run.Font.Bold, italic = (int)run.Font.Italic, underline = (int)run.Font.Underline,
                            color = (int)run.Font.Color.RGB, alignment = (int)run.ParagraphFormat.Alignment });
                    }
                    data["text_runs"] = runs;
                    data["text_bounds"] = new[] { (float)range.BoundLeft, (float)range.BoundTop, (float)range.BoundWidth, (float)range.BoundHeight };
                }
                if ((int)shape.HasTable != 0)
                {
                    dynamic table = shape.Table;
                    var rows = new List<object>();
                    for (var r = 1; r <= (int)table.Rows.Count; r++)
                    {
                        var cells = new List<object>();
                        for (var c = 1; c <= (int)table.Columns.Count; c++)
                        { dynamic cell = table.Cell(r, c).Shape; cells.Add(new { text = Convert.ToString(cell.TextFrame.TextRange.Text), width = (float)cell.Width, height = (float)cell.Height, font_size = (float)cell.TextFrame.TextRange.Font.Size }); }
                        rows.Add(cells);
                    }
                    data["table"] = rows;
                }
                if ((int)shape.HasChart != 0)
                {
                    dynamic chart = shape.Chart;
                    var series = new List<object>();
                    try
                    {
                        for (var n = 1; n <= (int)chart.SeriesCollection().Count; n++)
                        {
                            dynamic item = chart.SeriesCollection(n);
                            series.Add(new { name = Convert.ToString(item.Name), formula = Convert.ToString(item.Formula), values = Values((object)item.Values), categories = Values((object)item.XValues) });
                        }
                        data["chart"] = new { type = (int)chart.ChartType, series };
                    }
                    catch (Exception ex) when (ex is System.Runtime.InteropServices.COMException || ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    { unsupported.Add("chart-data:" + id); }
                }
                if (new[] { 7, 10, 12, 16, 21, 24 }.Contains((int)shape.Type)) unsupported.Add("preserve-object:" + id + ":type:" + (int)shape.Type);
                result.Add(data);
            }
            return result;
        }
        private static object[] Values(object value)
        { var sequence = value as IEnumerable; return sequence == null || value is string ? new[] { value } : sequence.Cast<object>().ToArray(); }
        internal static string ContentFingerprint(object slide)
        {
            return Scribble.Chat.TaskCheckpointStore.Fingerprint(new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.Serialize(Normalized(Capture(slide))));
        }
        private static object Normalized(object value)
        {
            var map = value as IDictionary<string, object>;
            if (map != null) return map.Where(p => p.Key != "slide_id" && p.Key != "index" && p.Key != "id" && p.Key != "unsupported")
                .ToDictionary(p => p.Key, p => Normalized(p.Value));
            if (value is string || value == null) return value;
            var list = value as IEnumerable;
            return list == null ? value : list.Cast<object>().Select(Normalized).ToArray();
        }
        public static string Fingerprint(object slide)
        {
            // Native render includes geometry, formatting and artwork not exposed as text.
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            return TaskCheckpointStore.Fingerprint(json.Serialize(Capture(slide)) + Preview(slide));
        }
        public static object ReadPage(object presentation, object slide, int offset, bool preview)
        {
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var content = json.Serialize(Capture(slide));
            if (offset < 0 || offset > content.Length) throw new InvalidOperationException("Invalid inspection page offset.");
            var count = Math.Min(12000, content.Length - offset);
            var render = Preview(slide);
            return new { presentation_id = IdentityFor(presentation), slide_id = (int)((dynamic)slide).SlideID,
                fingerprint = TaskCheckpointStore.Fingerprint(content + render), content = content.Substring(offset, count), offset, total_characters = content.Length,
                next_offset = offset + count < content.Length ? (int?)(offset + count) : null,
                image = preview ? render : null, untrusted_document_data = true };
        }
    }
}
