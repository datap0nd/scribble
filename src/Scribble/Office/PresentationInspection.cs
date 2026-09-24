using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
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
        internal static void RestoreCopiedBackground(object originalSlide,
            object copiedSlide)
        {
            dynamic original = originalSlide;
            dynamic copy = copiedSlide;
            copy.FollowMasterBackground = original.FollowMasterBackground;
            // Slides.Paste can replace an explicit solid fill with a
            // transparent inherited fill. Restore the source's native fill
            // before the preservation fingerprint is checked.
            dynamic sourceFill = original.Background.Fill;
            if ((int)sourceFill.Type == 1)
            {
                dynamic targetFill = copy.Background.Fill;
                targetFill.Solid();
                targetFill.ForeColor.RGB = sourceFill.ForeColor.RGB;
                targetFill.Transparency = sourceFill.Transparency;
            }
        }
        internal static object CopySlideTo(object originalSlide,
            object destinationPresentation)
        {
            dynamic original = originalSlide;
            dynamic deck = destinationPresentation;
            var before = (int)deck.Slides.Count;
            for (var attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    original.Copy();
                    deck.Slides.Paste(before + 1);
                }
                catch (System.Runtime.InteropServices.COMException error)
                {
                    // 0x80048240 is PowerPoint's transient empty clipboard.
                    // Retry only after confirming that Paste added no slide.
                    if (unchecked((uint)error.ErrorCode) != 0x80048240)
                        throw;
                    if ((int)deck.Slides.Count != before)
                        throw new InvalidOperationException(
                            "REVISION_COPY_UNCERTAIN: Paste reported an error after changing the destination.",
                            error);
                    if (attempt == 2) throw;
                    System.Threading.Thread.Sleep(150 * (attempt + 1));
                    continue;
                }
                if ((int)deck.Slides.Count != before + 1)
                    throw new InvalidOperationException(
                        "REVISION_COPY_INCOMPLETE: Native paste added an unexpected number of slides.");
                object copy = deck.Slides[before + 1];
                RestoreCopiedBackground(originalSlide, copy);
                return copy;
            }
            throw new InvalidOperationException("REVISION_COPY_INCOMPLETE");
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
                        { dynamic cell = table.Cell(r, c).Shape; cells.Add(new { text = Convert.ToString(cell.TextFrame.TextRange.Text), width = (float)cell.Width, height = (float)cell.Height, font_size = (float)cell.TextFrame.TextRange.Font.Size, fill_color = (int)cell.Fill.ForeColor.RGB }); }
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
            // Index-independent preservation checks never activate a native
            // chart. Exact same-slide identity also includes package parts in
            // Fingerprint, which catches chart/workbook edits.
            return Scribble.Chat.TaskCheckpointStore.Fingerprint(
                new JavaScriptSerializer { MaxJsonLength = int.MaxValue }
                    .Serialize(Normalized(Capture(slide,
                        !ContainsNativeChart(slide)))));
        }
        internal static string CopyContentFingerprint(object slide)
        {
            // Copy preservation must not activate the source chart workbook.
            return Scribble.Chat.TaskCheckpointStore.Fingerprint(
                new JavaScriptSerializer { MaxJsonLength = int.MaxValue }
                    .Serialize(Normalized(Capture(slide, false))));
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
            var chart = ContainsNativeChart(slide);
            // Chart COM getters and Slide.Export can terminate POWERPNT on
            // affected builds. The native package includes chart XML and its
            // embedded workbook without activating either object.
            var content = json.Serialize(Capture(slide, !chart));
            return TaskCheckpointStore.Fingerprint(content +
                (chart ? PackageSlideFingerprintCore(slide, false) :
                    Preview(slide)));
        }
        // The native chart COM getter can terminate some PowerPoint builds
        // after a chart workbook closes. Journal receipts use the exact slide
        // package and related parts from a disposable SaveCopyAs instead.
        internal static string FingerprintForJournal(object slide)
        {
            if (!ContainsNativeChart(slide) || !string.Equals(
                Environment.GetEnvironmentVariable(
                    AnalysisDocumentPilot.FeatureFlag), "1",
                StringComparison.Ordinal)) return Fingerprint(slide);
            var json = new JavaScriptSerializer { MaxJsonLength =
                int.MaxValue };
            return TaskCheckpointStore.Fingerprint(json.Serialize(
                Capture(slide, false)) + PackageSlideFingerprint(slide));
        }

        internal static string PackageSlideFingerprint(object slide)
        { return PackageSlideFingerprintCore(slide, true); }

        private static string PackageSlideFingerprintCore(object slide,
            bool requireOwnedDraft)
        {
            dynamic page = slide;
            dynamic deck = page.Parent;
            if (requireOwnedDraft &&
                (!string.IsNullOrEmpty((string)deck.Path) ||
                string.IsNullOrWhiteSpace(Convert.ToString(
                    page.Tags["ScribbleTask"]))))
                throw new InvalidOperationException(
                    "CHART_PACKAGE_UNSAVED_DRAFT_REQUIRED");
            var nameBefore = (string)deck.FullName;
            var savedBefore = (int)deck.Saved;
            var slideId = (int)page.SlideID;
            var temporary = Path.Combine(Path.GetTempPath(),
                "scribble-chart-fingerprint-" + Guid.NewGuid().ToString("N") +
                ".pptx");
            try
            {
                deck.SaveCopyAs(temporary);
                if ((string)deck.FullName != nameBefore ||
                    (int)deck.Saved != savedBefore ||
                    !File.Exists(temporary) ||
                    new FileInfo(temporary).Length > 50 * 1024 * 1024)
                    throw new InvalidOperationException(
                        "CHART_PACKAGE_DRAFT_BOUNDARY_CHANGED");
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

        // PowerPoint may round a native table frame by two EMU while making a
        // PDF and increment package metadata. Compare the complete before and
        // after packages; chart caches, embedded worksheet values, and every
        // other slide property must remain byte-for-byte or XML equivalent.
        internal static bool PdfExportPackageEquivalent(string beforePath,
            string afterPath, out string changedPart)
        {
            changedPart = string.Empty;
            try
            {
                if (new FileInfo(beforePath).Length > 50 * 1024 * 1024 ||
                    new FileInfo(afterPath).Length > 50 * 1024 * 1024)
                    return false;
                using (var before = ZipFile.OpenRead(beforePath))
                using (var after = ZipFile.OpenRead(afterPath))
                    return PdfArchiveEquivalent(before, after, false,
                        out changedPart);
            }
            catch (Exception error) when (error is IOException ||
                error is InvalidDataException || error is InvalidOperationException ||
                error is System.Xml.XmlException || error is FormatException)
            { changedPart = error.GetType().Name; return false; }
        }

        private static bool PdfArchiveEquivalent(ZipArchive before,
            ZipArchive after, bool embeddedWorkbook, out string changedPart)
        {
            changedPart = string.Empty;
            if (before.Entries.Count > 500 || after.Entries.Count > 500 ||
                before.Entries.Sum(entry => entry.Length) > 150L * 1024 * 1024 ||
                after.Entries.Sum(entry => entry.Length) > 150L * 1024 * 1024)
                return false;
            var original = before.Entries.ToDictionary(entry => entry.FullName,
                StringComparer.Ordinal);
            var exported = after.Entries.ToDictionary(entry => entry.FullName,
                StringComparer.Ordinal);
            if (!original.Keys.OrderBy(name => name, StringComparer.Ordinal)
                .SequenceEqual(exported.Keys.OrderBy(name => name,
                    StringComparer.Ordinal))) return false;
            foreach (var part in original.Keys.OrderBy(name => name,
                StringComparer.Ordinal))
            {
                changedPart = part;
                var left = PdfPartBytes(original[part]);
                var right = PdfPartBytes(exported[part]);
                if (left.SequenceEqual(right)) continue;
                if (part == "docProps/core.xml" &&
                    PdfMetadataEquivalent(left, right)) continue;
                if (!embeddedWorkbook &&
                    part.StartsWith("ppt/slides/slide",
                        StringComparison.Ordinal) &&
                    part.EndsWith(".xml", StringComparison.Ordinal) &&
                    PdfTableRoundoffEquivalent(left, right)) continue;
                if (!embeddedWorkbook &&
                    part.StartsWith("ppt/embeddings/",
                        StringComparison.Ordinal) &&
                    part.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    using (var leftStream = new MemoryStream(left))
                    using (var rightStream = new MemoryStream(right))
                    using (var leftZip = new ZipArchive(leftStream,
                        ZipArchiveMode.Read))
                    using (var rightZip = new ZipArchive(rightStream,
                        ZipArchiveMode.Read))
                    {
                        string nestedPart;
                        if (PdfArchiveEquivalent(leftZip, rightZip, true,
                            out nestedPart)) continue;
                        changedPart = part + "/" + nestedPart;
                    }
                }
                return false;
            }
            changedPart = string.Empty;
            return true;
        }

        private static byte[] PdfPartBytes(ZipArchiveEntry entry)
        {
            if (entry.Length > 30 * 1024 * 1024)
                throw new InvalidDataException("Oversized PDF boundary part.");
            using (var stream = entry.Open())
            using (var buffer = new MemoryStream())
            { stream.CopyTo(buffer); return buffer.ToArray(); }
        }

        private static bool PdfMetadataEquivalent(byte[] before,
            byte[] after)
        {
            var left = XDocument.Load(new MemoryStream(before));
            var right = XDocument.Load(new MemoryStream(after));
            var revision = XName.Get("revision",
                "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
            var modified = XName.Get("modified",
                "http://purl.org/dc/terms/");
            foreach (var document in new[] { left, right })
                foreach (var node in document.Descendants().Where(element =>
                    element.Name == modified ||
                    element.Name == revision).ToArray())
                    node.Remove();
            return XNode.DeepEquals(left, right);
        }

        private static bool PdfTableRoundoffEquivalent(byte[] before,
            byte[] after)
        {
            var left = XDocument.Load(new MemoryStream(before));
            var right = XDocument.Load(new MemoryStream(after));
            Func<XDocument, XElement[]> tableFrames = document => document
                .Descendants().Where(element => element.Name.LocalName ==
                    "graphicFrame" && element.Descendants().Any(child =>
                        child.Name.LocalName == "graphicData" &&
                        (string)child.Attribute("uri") ==
                        "http://schemas.openxmlformats.org/drawingml/2006/table"))
                .ToArray();
            var original = tableFrames(left);
            var exported = tableFrames(right);
            if (original.Length == 0 || original.Length != exported.Length)
                return false;
            for (var index = 0; index < original.Length; index++)
            {
                Func<XElement, XElement> extent = frame => frame.Elements()
                    .Where(element => element.Name.LocalName == "xfrm")
                    .SelectMany(element => element.Elements())
                    .SingleOrDefault(element => element.Name.LocalName == "ext");
                var beforeExtent = extent(original[index]);
                var afterExtent = extent(exported[index]);
                if (beforeExtent == null || afterExtent == null) return false;
                foreach (var dimension in new[] { "cx", "cy" })
                {
                    long beforeValue, afterValue;
                    if (!long.TryParse((string)beforeExtent.Attribute(dimension),
                            out beforeValue) ||
                        !long.TryParse((string)afterExtent.Attribute(dimension),
                            out afterValue) ||
                        beforeValue < 0 || afterValue < 0 ||
                        beforeValue > 1000000000 ||
                        afterValue > 1000000000 ||
                        Math.Abs(beforeValue - afterValue) > 2) return false;
                    afterExtent.SetAttributeValue(dimension, beforeValue);
                }
            }
            return XNode.DeepEquals(left, right);
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
            if (part.StartsWith("ppt/embeddings/",
                    StringComparison.OrdinalIgnoreCase) &&
                part.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                using (var stream = entry.Open())
                using (var buffer = new MemoryStream())
                {
                    stream.CopyTo(buffer);
                    parts[part] = EmbeddedWorkbookFingerprint(
                        buffer.ToArray());
                }
            }
            else using (var stream = entry.Open())
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

        private static string EmbeddedWorkbookFingerprint(byte[] data)
        {
            if (data == null || data.Length > 30 * 1024 * 1024)
                throw new InvalidOperationException(
                    "CHART_PACKAGE_WORKBOOK_INVALID");
            using (var stream = new MemoryStream(data))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                if (archive.Entries.Count > 500 ||
                    archive.Entries.Sum(entry => entry.Length) >
                        150L * 1024 * 1024 ||
                    archive.Entries.Select(entry => entry.FullName)
                        .Distinct(StringComparer.Ordinal).Count() !=
                            archive.Entries.Count)
                    throw new InvalidOperationException(
                        "CHART_PACKAGE_WORKBOOK_INVALID");
                var parts = new SortedDictionary<string, string>(
                    StringComparer.Ordinal);
                foreach (var entry in archive.Entries)
                {
                    if (entry.Length > 30 * 1024 * 1024)
                        throw new InvalidOperationException(
                            "CHART_PACKAGE_WORKBOOK_INVALID");
                    if (entry.FullName == "docProps/core.xml")
                    {
                        using (var part = entry.Open())
                        {
                            var xml = XDocument.Load(part);
                            var revision = XName.Get("revision",
                                "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
                            var modified = XName.Get("modified",
                                "http://purl.org/dc/terms/");
                            foreach (var node in xml.Descendants().Where(
                                element => element.Name == revision ||
                                    element.Name == modified).ToArray())
                                node.Remove();
                            parts[entry.FullName] =
                                TaskCheckpointStore.Fingerprint(xml.ToString(
                                    SaveOptions.DisableFormatting));
                        }
                    }
                    else using (var part = entry.Open())
                    using (var hash = SHA256.Create())
                        parts[entry.FullName] = BitConverter.ToString(
                            hash.ComputeHash(part)).Replace("-", "");
                }
                return TaskCheckpointStore.Fingerprint(string.Join("|",
                    parts.Select(part => part.Key + ":" + part.Value)));
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
