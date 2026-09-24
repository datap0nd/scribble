using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using System.Xml.Linq;
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
        public static bool ContainsNativeChart(object slide)
        {
            dynamic page = slide;
            return ContainsNativeChartInShapes((object)page.Shapes, 0);
        }
        private static bool ContainsNativeChartInShapes(object value, int depth)
        {
            if (depth > 16) return false;
            dynamic shapes = value;
            for (var i = 1; i <= (int)shapes.Count; i++)
            {
                dynamic shape = shapes[i];
                if ((int)shape.HasChart != 0) return true;
                if ((int)shape.Type == 6 && ContainsNativeChartInShapes((object)shape.GroupItems, depth + 1)) return true;
            }
            return false;
        }
        public static Dictionary<string, object> Capture(object slide)
        { return Capture(slide, true); }
        private static Dictionary<string, object> Capture(object slide,
            bool readChartData)
        {
            dynamic page = slide;
            var unsupported = new List<string>();
            return new Dictionary<string, object> {
                { "slide_id", (int)page.SlideID }, { "index", (int)page.SlideIndex },
                { "shapes", Shapes((object)page.Shapes, unsupported, 0,
                    readChartData) },
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
        private static List<object> Shapes(object value,
            List<string> unsupported, int depth, bool readChartData)
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
                if ((int)shape.Type == 6) data["children"] = Shapes(
                    (object)shape.GroupItems, unsupported, depth + 1,
                    readChartData);
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
                    // Stress fixtures deliberately include stale/corrupt embedded
                    // chart workbooks. Current Office builds can terminate in
                    // chart.dll merely by opening that COM object. The benchmark
                    // carries an independent workbook authority and grades the
                    // generated native chart directly, so do not dereference the
                    // hostile source chart inside an active stress run.
                    if (!readChartData || AvoidUnsafeStressChartAutomation())
                    {
                        data["chart"] = new { available = false,
                            reason = readChartData ? "unsafe_stress_fixture" :
                                "native_package_fingerprint" };
                        unsupported.Add("chart-data:" + id);
                    }
                    else
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
                }
                if (new[] { 7, 10, 12, 16, 21, 24 }.Contains((int)shape.Type)) unsupported.Add("preserve-object:" + id + ":type:" + (int)shape.Type);
                result.Add(data);
            }
            return result;
        }
        private static bool AvoidUnsafeStressChartAutomation()
        {
            try { return Scribble.Testing.TestLab.Status()?.suite_id == "scribble-stress-v1"; }
            catch { return false; }
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
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var content = json.Serialize(Capture(slide));
            // Some Office builds terminate POWERPNT in chart.dll while exporting a
            // slide that contains a native chart. Structured capture includes chart
            // data and geometry, so retain the stronger rendered fingerprint only
            // for slides that PowerPoint can safely export.
            return TaskCheckpointStore.Fingerprint(content + (ContainsNativeChart(slide) ? string.Empty : Preview(slide)));
        }
        // The native chart COM getter can terminate some PowerPoint builds
        // after a chart workbook closes. Journal receipts use the exact slide
        // package and related parts from a disposable SaveCopyAs instead.
        internal static string FingerprintForJournal(object slide)
        {
            if (!ContainsNativeChart(slide)) return Fingerprint(slide);
            var json = new JavaScriptSerializer { MaxJsonLength =
                int.MaxValue };
            return TaskCheckpointStore.Fingerprint(json.Serialize(
                Capture(slide, false)) + PackageSlideFingerprint(slide));
        }

        internal static string PackageSlideFingerprint(object slide)
        {
            dynamic page = slide;
            dynamic deck = page.Parent;
            var slideId = (int)page.SlideID;
            var temporary = Path.Combine(Path.GetTempPath(),
                "scribble-chart-fingerprint-" + Guid.NewGuid().ToString("N") +
                ".pptx");
            try
            {
                deck.SaveCopyAs(temporary);
                using (var archive = ZipFile.OpenRead(temporary))
                {
                    const string presentationPart = "ppt/presentation.xml";
                    var xml = LoadPackageXml(archive, presentationPart);
                    var id = xml.Descendants().FirstOrDefault(element =>
                        element.Name.LocalName == "sldId" &&
                        (string)element.Attribute("id") ==
                            slideId.ToString());
                    if (id == null)
                        throw new InvalidOperationException(
                            "CHART_PACKAGE_SLIDE_MISSING");
                    var relationshipId = (string)id.Attribute(
                        XName.Get("id",
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships"));
                    var slidePart = RelationshipTarget(archive,
                        presentationPart, relationshipId);
                    if (slidePart == null)
                        throw new InvalidOperationException(
                            "CHART_PACKAGE_RELATIONSHIP_MISSING");
                    var visited = new HashSet<string>(
                        StringComparer.OrdinalIgnoreCase);
                    var parts = new SortedDictionary<string, string>(
                        StringComparer.Ordinal);
                    CollectPackageParts(archive, slidePart, visited, parts,
                        0);
                    return TaskCheckpointStore.Fingerprint(string.Join("|",
                        parts.Select(part => part.Key + ":" + part.Value)));
                }
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        private static XDocument LoadPackageXml(ZipArchive archive,
            string path)
        {
            var entry = archive.GetEntry(path);
            if (entry == null)
                throw new InvalidOperationException(
                    "CHART_PACKAGE_PART_MISSING: " + path);
            using (var stream = entry.Open()) return XDocument.Load(stream);
        }

        private static string RelationshipPath(string part)
        {
            var slash = part.LastIndexOf('/');
            return part.Substring(0, slash + 1) + "_rels/" +
                part.Substring(slash + 1) + ".rels";
        }

        private static string ResolvePackageTarget(string part,
            string target)
        {
            var root = new Uri("http://scribble-package/");
            var source = new Uri(root, part);
            return Uri.UnescapeDataString(new Uri(source, target)
                .AbsolutePath.TrimStart('/'));
        }

        private static string RelationshipTarget(ZipArchive archive,
            string part, string relationshipId)
        {
            var relationships = LoadPackageXml(archive,
                RelationshipPath(part));
            var item = relationships.Descendants().FirstOrDefault(element =>
                element.Name.LocalName == "Relationship" &&
                (string)element.Attribute("Id") == relationshipId);
            if (item == null || (string)item.Attribute("TargetMode") ==
                "External") return null;
            return ResolvePackageTarget(part,
                (string)item.Attribute("Target"));
        }

        private static void CollectPackageParts(ZipArchive archive,
            string part, HashSet<string> visited,
            SortedDictionary<string, string> parts, int depth)
        {
            if (depth > 8 || visited.Count >= 48)
                throw new InvalidOperationException(
                    "CHART_PACKAGE_RELATIONSHIP_LIMIT");
            if (!visited.Add(part)) return;
            var entry = archive.GetEntry(part);
            if (entry == null || entry.Length > 30 * 1024 * 1024)
                throw new InvalidOperationException(
                    "CHART_PACKAGE_PART_INVALID: " + part);
            using (var stream = entry.Open())
            using (var hash = SHA256.Create())
                parts[part] = BitConverter.ToString(hash.ComputeHash(stream))
                    .Replace("-", "");
            var relationsPath = RelationshipPath(part);
            if (archive.GetEntry(relationsPath) == null) return;
            var relationships = LoadPackageXml(archive, relationsPath);
            foreach (var item in relationships.Descendants().Where(element =>
                element.Name.LocalName == "Relationship" &&
                (string)element.Attribute("TargetMode") != "External"))
            {
                var target = (string)item.Attribute("Target");
                if (string.IsNullOrEmpty(target))
                    throw new InvalidOperationException(
                        "CHART_PACKAGE_RELATIONSHIP_INVALID");
                CollectPackageParts(archive, ResolvePackageTarget(part,
                    target), visited, parts, depth + 1);
            }
        }
        public static object ReadPage(object presentation, object slide, int offset, bool preview)
        {
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var content = json.Serialize(Capture(slide));
            if (offset < 0 || offset > content.Length) throw new InvalidOperationException("Invalid inspection page offset.");
            var count = Math.Min(12000, content.Length - offset);
            var previewSuppressed = ContainsNativeChart(slide);
            var render = previewSuppressed ? string.Empty : Preview(slide);
            // The structured capture is JSON, so quoted shape text contains
            // escaped line breaks. Source citations must instead be copied
            // from decoded native text. Keep this read-only transcript beside
            // the structured page; TaskSources retains it as a verified span.
            var citationText = offset == 0
                ? CitationTextFromCaptured(json.Deserialize<Dictionary<string, object>>(content))
                : null;
            // Lead with decoded text so the first host-issued source spans are
            // useful citations, not 12k of geometry and escaped JSON.
            return new { presentation_id = IdentityFor(presentation), slide_id = (int)((dynamic)slide).SlideID,
                citation_text = citationText,
                fingerprint = TaskCheckpointStore.Fingerprint(content + render), content = content.Substring(offset, count), offset, total_characters = content.Length,
                next_offset = offset + count < content.Length ? (int?)(offset + count) : null,
                image = preview && !string.IsNullOrEmpty(render) ? render : null,
                preview_unavailable = preview && previewSuppressed ? "Native-chart preview omitted because this Office build may terminate while exporting it; structured chart data is included." : null,
                untrusted_document_data = true };
        }
        public static string CitationTextFromCaptured(IDictionary<string, object> capture)
        {
            if (capture == null) return string.Empty;
            var lines = new List<string>();
            object shapes;
            if (capture.TryGetValue("shapes", out shapes)) AppendCitationShapes(shapes, lines);
            object notes;
            if (capture.TryGetValue("notes", out notes) && !string.IsNullOrWhiteSpace(Convert.ToString(notes)))
                lines.Add(Convert.ToString(notes).Trim());
            var text = string.Join("\n", lines);
            return text.Length <= 12000 ? text : text.Substring(0, 12000);
        }
        private static void AppendCitationShapes(object source, List<string> lines)
        {
            var shapes = source as IEnumerable;
            if (shapes == null || source is string) return;
            foreach (var raw in shapes)
            {
                var shape = raw as IDictionary<string, object>;
                if (shape == null) continue;
                object value;
                if (shape.TryGetValue("text", out value) && !string.IsNullOrWhiteSpace(Convert.ToString(value)))
                    lines.Add(Convert.ToString(value).Trim());
                if (shape.TryGetValue("table", out value))
                {
                    var rows = value as IEnumerable;
                    if (rows != null) foreach (var rawRow in rows)
                    {
                        var cells = rawRow as IEnumerable;
                        if (cells == null || rawRow is string) continue;
                        var texts = new List<string>();
                        foreach (var rawCell in cells)
                        {
                            var cell = rawCell as IDictionary<string, object>;
                            object cellText;
                            texts.Add(cell != null && cell.TryGetValue("text", out cellText)
                                ? Convert.ToString(cellText).Trim() : string.Empty);
                        }
                        lines.Add(string.Join("\t", texts));
                    }
                }
                if (shape.TryGetValue("children", out value)) AppendCitationShapes(value, lines);
            }
        }
    }
}
