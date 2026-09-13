using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;

namespace Scribble.Testing
{
    // Invoked only by the operator window. Only documents tagged at creation by this run qualify.
    public static class BenchmarkArtifactCollector
    {
        public static string Capture(string runId, string phase = "final")
        {
            if (!new[] { "source", "intermediate", "final" }.Contains(phase)) throw new ArgumentException("Invalid capture phase.");
            var run = TestLab.GetRun(runId);
            var directory = Path.Combine(TestLab.RunDirectory(runId), "capture", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            var nativeDirectory = TestLab.NativeDirectory(runId, "capture-" + Guid.NewGuid().ToString("N").Substring(0, 8));
            var report = new List<string>();
            var destinations = new Dictionary<string, string> { { "xlsx", "Excel" }, { "pptx", "PowerPoint" }, { "docx", "Word" }, { "msg", "Outlook" } };
            var hosts = new HashSet<string>(new[] { run.host }.Concat((run.required_artifacts ?? new string[0]).Where(destinations.ContainsKey).Select(a => destinations[a])));
            // An unrelated busy Outlook window must not stall an Excel capture.
            foreach (var kind in new[] { "Excel", "PowerPoint", "Word" }.Where(hosts.Contains))
            {
                object instance = null;
                bool releaseInstance = true;
                try
                {
                    instance = TestLabOfficeEnvironment.Borrow(kind, out releaseInstance); dynamic app = instance;
                    dynamic documents = kind == "Excel" ? app.Workbooks : kind == "PowerPoint" ? app.Presentations : app.Documents;
                    for (int i = 1; i <= (int)documents.Count; i++)
                    {
                        object value = documents.Item(i); dynamic document = value;
                        try
                        {
                            var runOutput = TestLab.IsRunOutput(value, run.run_id);
                            var originalPath = Convert.ToString(document.FullName);
                            if (runOutput) {
                                try { originalPath = Convert.ToString(document.CustomDocumentProperties["ScribbleTestSourcePath"].Value); }
                                catch { originalPath = null; }
                            }
                            var source = TestLabSuite.OwnsNativeSource(
                                run.run_id,
                                originalPath,
                                kind);
                            if (!runOutput && !source) continue;
                            SourceBoundary boundary = null;
                            if (source)
                            {
                                var boundaryPath = Path.Combine(TestLab.RunDirectory(runId), "source-boundary-" +
                                    Scribble.Chat.TaskCheckpointStore.Fingerprint(originalPath.ToUpperInvariant()) + ".json");
                                if (phase == "source")
                                {
                                    boundary = SourceBoundary.Read(value, kind);
                                    File.WriteAllText(boundaryPath, TestLab.Serialize(boundary), new UTF8Encoding(false));
                                }
                                else if (File.Exists(boundaryPath)) boundary = TestLabSuite.Read<SourceBoundary>(boundaryPath);
                                else throw new InvalidDataException("The source boundary was not captured before this case. Its existing draft labels cannot establish a new output.");
                            }

                            // Excel drafts normally live in a new worksheet of the
                            // read-only fixture workbook.  The workbook itself is
                            // intentionally not tagged as a run-created document,
                            // so retain a separate final copy once a Scribble Draft
                            // sheet exists.  Earlier code captured only its readback
                            // and then reported the requested XLSX as missing.
                            var sourceDerivedOutput =
                                source &&
                                phase != "source" &&
                                HasNewDraft(value, kind, boundary);
                            var effectiveOutput = runOutput || sourceDerivedOutput;
                            var role = effectiveOutput ? "output" : "source";
                            var stem = Path.Combine(
                                directory,
                                kind + "-" + phase + "-" + role + "-" + i);
                            try {
                                var readback = ReadNativeWithinBoundary(value, kind, !effectiveOutput, boundary);
                                var readbackExtension = kind == "Excel" ? ".xlsx" : kind == "PowerPoint" ? ".pptx" : ".docx";
                                File.WriteAllText(stem + "-readback.json", TestLab.Serialize(new { schema = 1, run_id = runId,
                                    host = kind, captured_utc = DateTime.UtcNow.ToString("O"), native_readback = true,
                                    phase = phase, run_created_output = effectiveOutput,
                                    artifact_extension = effectiveOutput ? readbackExtension.TrimStart('.') : null,
                                    text = readback }), Encoding.UTF8);
                                TestLab.Collect(runId, stem + "-readback.json");
                                report.Add("Captured " + kind + " cell/text/structure readback.");
                            } catch (Exception ex) { report.Add(kind + " readback failed: " + ex.Message); }
                            if (sourceDerivedOutput)
                            {
                                // Preserve a source-only readback as well.  Source
                                // preservation compares the original worksheets,
                                // while the output readback and XLSX include the new
                                // draft sheet.
                                var sourceStem = Path.Combine(
                                    directory,
                                    kind + "-" + phase + "-source-" + i);
                                var sourceReadback = ReadNativeWithinBoundary(value, kind, true, boundary);
                                File.WriteAllText(
                                    sourceStem + "-readback.json",
                                    TestLab.Serialize(new {
                                        schema = 1,
                                        run_id = runId,
                                        host = kind,
                                        captured_utc = DateTime.UtcNow.ToString("O"),
                                        native_readback = true,
                                        phase = phase,
                                        run_created_output = false,
                                        text = sourceReadback
                                    }),
                                    Encoding.UTF8);
                                TestLab.Collect(runId, sourceStem + "-readback.json");
                            }
                            if (!effectiveOutput) continue;
                            stem = Path.Combine(nativeDirectory, kind + "-" + phase + "-output-" + i);
                            var extension = kind == "Excel" ? ".xlsx" : kind == "PowerPoint" ? ".pptx" : ".docx";
                            try {
                            if (kind == "Excel") document.SaveCopyAs(stem + extension);
                            else if (kind == "PowerPoint") document.SaveCopyAs(stem + extension, 24);
                            else SaveFlatOpc(Convert.ToString(document.WordOpenXML), stem + extension);
                            if (!IsValidOfficePackage(stem + extension, kind))
                            {
                                // NASCA/MarkAny may transparently replace an
                                // Office SaveCopyAs payload with DRM ciphertext.
                                // The live COM readback above remains authoritative;
                                // never collect encrypted bytes as if they were OOXML.
                                File.Delete(stem + extension);
                                report.Add(
                                    kind + " disk copy was DRM-wrapped; retained " +
                                    "the in-memory native readback instead.");
                                // A protected native package must not suppress
                                // independently exportable visual evidence.
                            }
                            if (File.Exists(stem + extension)) {
                                var collected = TestLab.Collect(runId, stem + extension);
                                report.Add("Captured " + Path.GetFileName(collected));
                            }
                            } catch (Exception ex) { report.Add(kind + " native file export failed; retained live readback: " + ex.Message); }
                            try
                            {
                                if (kind == "Excel") document.ExportAsFixedFormat(0, stem + ".pdf");
                                else if (kind == "PowerPoint") document.SaveCopyAs(stem + ".pdf", 32);
                                else document.ExportAsFixedFormat(stem + ".pdf", 17);
                                TestLab.Collect(runId, stem + ".pdf");
                            }
                            catch (Exception ex) { report.Add(kind + " PDF export failed: " + ex.Message); }
                            if (kind == "PowerPoint") for (int n = 1; n <= (int)document.Slides.Count; n++)
                            {
                                try
                                {
                                    var png = stem + "-slide-" + n + ".png";
                                    document.Slides.Item(n).Export(png, "PNG", 1280, 720); TestLab.Collect(runId, png);
                                }
                                catch (Exception ex) { report.Add("Slide " + n + " preview export failed: " + ex.Message); }
                            }
                        }
                        catch (Exception ex) { report.Add(kind + " capture: " + ex.Message); }
                        finally { if (Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
                    }
                }
                catch (COMException) { report.Add(kind + ": no accessible running app. Save and collect manually if needed."); }
                finally { if (releaseInstance && instance != null && Marshal.IsComObject(instance)) Marshal.ReleaseComObject(instance); }
            }
            object outlookInstance = null;
            bool releaseOutlook = true;
            if (hosts.Contains("Outlook"))
            {
            try
            {
                outlookInstance = TestLabOfficeEnvironment.Borrow("Outlook", out releaseOutlook); dynamic outlook = outlookInstance;
                for (int i = 1; i <= (int)outlook.Inspectors.Count; i++)
                {
                    object value = outlook.Inspectors.Item(i).CurrentItem; dynamic mail = value;
                    try
                    {
                        dynamic tag = mail.UserProperties.Find("ScribbleTestRunId");
                        if (tag == null || Convert.ToString(tag.Value) != runId) continue;
                        if (Convert.ToBoolean(mail.Sent)) throw new InvalidOperationException("A run-owned email has been sent. It cannot be certified as an unsent draft.");
                        var stem = Path.Combine(directory, "Outlook-" + phase + "-output-" + i);
                        var nativeText = "To: " + Convert.ToString(mail.To) + "\nCC: " + Convert.ToString(mail.CC) +
                            "\nSubject: " + Convert.ToString(mail.Subject) + "\nUnsent: true\n" + Convert.ToString(mail.Body);
                        File.WriteAllText(stem + "-readback.json", TestLab.Serialize(new { schema = 1, run_id = runId,
                            native_readback = true, phase, run_created_output = true, artifact_extension = "msg", unsent = true, text = nativeText }), Encoding.UTF8);
                        TestLab.Collect(runId, stem + "-readback.json");
                        var nativeMail = Path.Combine(nativeDirectory, Path.GetFileName(stem) + ".msg");
                        try { mail.SaveAs(nativeMail, 9); TestLab.Collect(runId, nativeMail); }
                        catch (Exception error) { report.Add("Outlook native save failed; retained live draft readback: " + error.Message); }
                        File.WriteAllText(stem + ".html", Convert.ToString(mail.HTMLBody), Encoding.UTF8);
                        TestLab.Collect(runId, stem + ".html");
                        var attachments = new List<object>();
                        for (int n = 1; n <= (int)mail.Attachments.Count; n++)
                        {
                            dynamic attachment = mail.Attachments.Item(n);
                            var saved = Path.Combine(nativeDirectory, "mail-" + i + "-attachment-" + n + ".bin");
                            attachment.SaveAsFile(saved);
                            attachments.Add(new { name = Convert.ToString(attachment.FileName), sha256 = TestLab.FileHash(saved), size = new FileInfo(saved).Length });
                        }
                        File.WriteAllText(stem + ".json", TestLab.Serialize(new { schema = 1, run_id = runId, native_readback = true, unsent = true,
                            to = Convert.ToString(mail.To), cc = Convert.ToString(mail.CC), subject = Convert.ToString(mail.Subject), body = Convert.ToString(mail.Body), attachments }), Encoding.UTF8);
                        TestLab.Collect(runId, stem + ".json");
                        report.Add("Captured unsent Outlook draft and attachment hashes.");
                    }
                    finally { if (Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
                }
            }
            catch (COMException) { report.Add("Outlook: no accessible running app. Save the test draft as MSG and collect manually if needed."); }
            finally { if (releaseOutlook && outlookInstance != null && Marshal.IsComObject(outlookInstance)) Marshal.ReleaseComObject(outlookInstance); }
            }
            if (report.Count == 0) report.Add("No new run-owned document found. For drafts inside a source workbook/deck, save a separate copy and use Collect saved outputs. Save Outlook drafts as MSG.");
            return string.Join(Environment.NewLine, report);
        }
        private static object At(object value, int row, int column)
        {
            var array = value as Array;
            return array == null ? value : array.GetValue(row + array.GetLowerBound(0), column + array.GetLowerBound(1));
        }
        private static bool HasScribbleDraft(object value, string kind)
        {
            if (kind == "PowerPoint")
            {
                dynamic presentation = value;
                for (int i = 1; i <= (int)presentation.Slides.Count; i++)
                    if (IsDraftSlide((object)presentation.Slides.Item(i))) return true;
                return false;
            }
            if (kind != "Excel") return false;
            dynamic workbook = value;
            for (int n = 1; n <= (int)workbook.Worksheets.Count; n++)
            {
                var name = Convert.ToString(workbook.Worksheets.Item(n).Name) ?? "";
                if (name == "Scribble Draft" ||
                    name.StartsWith("Scribble Draft ", StringComparison.Ordinal))
                    return true;
            }
            return false;
        }

        private sealed class SourceBoundary
        {
            public SourceBoundary() { }
            public string[] sheets { get; set; } = new string[0];
            public int[] slides { get; set; } = new int[0];
            internal static SourceBoundary Read(object value, string kind)
            {
                dynamic document = value;
                var boundary = new SourceBoundary();
                if (kind == "Excel") boundary.sheets = Enumerable.Range(1, (int)document.Worksheets.Count)
                    .Select(i => Convert.ToString(document.Worksheets.Item(i).Name)).Cast<string>().ToArray();
                if (kind == "PowerPoint") boundary.slides = Enumerable.Range(1, (int)document.Slides.Count)
                    .Select(i => Convert.ToInt32(document.Slides.Item(i).SlideID)).Cast<int>().ToArray();
                return boundary;
            }
        }

        private static bool HasNewDraft(object value, string kind, SourceBoundary boundary)
        {
            if (boundary == null) return false;
            dynamic document = value;
            if (kind == "PowerPoint")
            {
                for (int i = 1; i <= (int)document.Slides.Count; i++)
                    if (!boundary.slides.Contains((int)Convert.ToInt32(document.Slides.Item(i).SlideID)) &&
                        IsDraftSlide((object)document.Slides.Item(i))) return true;
            }
            if (kind == "Excel")
            {
                for (int i = 1; i <= (int)document.Worksheets.Count; i++)
                {
                    string name = Convert.ToString(document.Worksheets.Item(i).Name);
                    if (!boundary.sheets.Contains(name) && (name == "Scribble Draft" || name.StartsWith("Scribble Draft ", StringComparison.Ordinal))) return true;
                }
            }
            return false;
        }
        private static bool IsDraftSlide(object value)
        {
            dynamic slide = value;
            for (int i = 1; i <= (int)slide.Shapes.Count; i++) {
                dynamic shape = slide.Shapes.Item(i);
                if ((int)shape.HasTextFrame != 0 && Convert.ToString(shape.TextFrame.TextRange.Text).Contains("[Scribble draft]")) return true;
            }
            return false;
        }

        private static bool IsValidOfficePackage(string path, string kind)
        {
            try
            {
                if (!File.Exists(path)) return false;
                using (var archive = ZipFile.OpenRead(path))
                {
                    var required =
                        kind == "Excel" ? "xl/workbook.xml" :
                        kind == "PowerPoint" ? "ppt/presentation.xml" :
                        "word/document.xml";
                    return archive.GetEntry("[Content_Types].xml") != null &&
                           archive.GetEntry(required) != null;
                }
            }
            catch (InvalidDataException)
            {
                return false;
            }
            catch (IOException)
            {
                return false;
            }
        }

        private static string ReadNative(object value, string kind)
        {
            return ReadNative(value, kind, false);
        }

        private static string ReadNative(object value, string kind, bool excludeScribbleDrafts)
        { return ReadNativeWithinBoundary(value, kind, excludeScribbleDrafts, null); }

        private static string ReadNativeWithinBoundary(object value, string kind, bool excludeScribbleDrafts, SourceBoundary boundary)
        {
            dynamic document = value;
            var text = new StringBuilder();
            if (kind == "Word") return Convert.ToString(document.Content.Text);
            if (kind == "Excel") {
                for (int n = 1; n <= (int)document.Worksheets.Count; n++) {
                    dynamic sheet = document.Worksheets.Item(n); dynamic used = sheet.UsedRange;
                    string sheetName = Convert.ToString(sheet.Name) ?? "";
                    if (excludeScribbleDrafts && (boundary != null ? !boundary.sheets.Contains(sheetName) :
                        (sheetName == "Scribble Draft" || sheetName.StartsWith("Scribble Draft ", StringComparison.Ordinal))))
                        continue;
                    int rows = (int)used.Rows.Count, columns = (int)used.Columns.Count;
                    text.AppendLine("Worksheet: " + sheetName + " | used range: " + Convert.ToString(used.Address));
                    if ((long)rows * columns > 100000) { text.AppendLine("READBACK LIMIT: range exceeds 100000 cells. Screenshot the relevant range in Excel."); continue; }
                    object values = used.Value2, formulas = used.Formula;
                    for (int r = 0; r < rows; r++) for (int c = 0; c < columns; c++) {
                        var cell = At(values, r, c); var formula = Convert.ToString(At(formulas, r, c));
                        if (cell is ErrorWrapper) cell = Convert.ToString(used.Cells.Item(r + 1, c + 1).Text);
                        if (cell == null && string.IsNullOrEmpty(formula)) continue;
                        text.Append("R").Append((int)used.Row + r).Append("C").Append((int)used.Column + c)
                            .Append(": ").Append(Convert.ToString(cell, System.Globalization.CultureInfo.InvariantCulture));
                        if (formula.StartsWith("=", StringComparison.Ordinal)) text.Append(" | formula: ").Append(formula);
                        text.AppendLine();
                    }
                    dynamic charts = sheet.ChartObjects();
                    text.AppendLine("Native charts: " + (int)charts.Count);
                    for (int c = 1; c <= (int)charts.Count; c++) {
                        dynamic chart = charts.Item(c).Chart;
                        text.AppendLine("Chart " + c + " | type: " + Convert.ToString(chart.ChartType) + " | title: " + ((bool)chart.HasTitle ? Convert.ToString(chart.ChartTitle.Text) : ""));
                        dynamic series = chart.SeriesCollection();
                        for (int j = 1; j <= (int)series.Count; j++) text.AppendLine("Series " + j + ": " + Convert.ToString(series.Item(j).Formula));
                    }
                }
                text.AppendLine("Values are native cached readback; no recalculation was forced. Screenshot charts and formatting for visual review.");
            } else {
                var selectedSlides = new List<int>();
                for (int n = 1; n <= (int)document.Slides.Count; n++)
                    if (!excludeScribbleDrafts || (boundary != null ? boundary.slides.Contains((int)Convert.ToInt32(document.Slides.Item(n).SlideID)) :
                        !IsDraftSlide((object)document.Slides.Item(n)))) selectedSlides.Add(n);
                text.AppendLine("Slide count: " + selectedSlides.Count);
                for (int n = 1; n <= (int)document.Slides.Count; n++) {
                    if (!selectedSlides.Contains(n)) continue;
                    dynamic slide = document.Slides.Item(n); text.AppendLine("Slide " + (excludeScribbleDrafts && boundary != null ? selectedSlides.IndexOf(n) + 1 : n));
                    for (int j = 1; j <= (int)slide.Shapes.Count; j++) {
                        dynamic shape = slide.Shapes.Item(j);
                        text.AppendLine("Shape " + j + " | type " + Convert.ToString(shape.Type) + " | x,y,w,h: " +
                            Convert.ToString(shape.Left) + "," + Convert.ToString(shape.Top) + "," + Convert.ToString(shape.Width) + "," + Convert.ToString(shape.Height));
                        if ((int)shape.HasTextFrame != 0) text.AppendLine(Convert.ToString(shape.TextFrame.TextRange.Text));
                        if ((int)shape.HasTable != 0) {
                            dynamic table = shape.Table;
                            for (int r = 1; r <= (int)table.Rows.Count; r++) for (int c = 1; c <= (int)table.Columns.Count; c++)
                                text.AppendLine("Table R" + r + "C" + c + ": " + Convert.ToString(table.Cell(r, c).Shape.TextFrame.TextRange.Text));
                        }
                        if ((int)shape.HasChart != 0) text.AppendLine("Native chart type: " + Convert.ToString(shape.Chart.ChartType));
                    }
                    try
                    {
                        dynamic notes = slide.NotesPage;
                        var noteText = new StringBuilder();
                        for (int j = 1; j <= (int)notes.Shapes.Count; j++)
                        {
                            dynamic noteShape = notes.Shapes.Item(j);
                            if ((int)noteShape.HasTextFrame != 0)
                            {
                                var noteValue = Convert.ToString(
                                    noteShape.TextFrame.TextRange.Text);
                                if (!string.IsNullOrWhiteSpace(noteValue))
                                    noteText.Append(noteValue).Append(" ");
                            }
                        }
                        if (noteText.Length > 0)
                            text.AppendLine("Notes: " + noteText.ToString().Trim());
                    }
                    catch
                    {
                    }
                }
            }
            return text.ToString();
        }
        public static void SaveFlatOpc(string xml, string path)
        {
            if (xml == null || xml.Length > 100 * 1024 * 1024) throw new InvalidDataException("Word output exceeds capture budget.");
            XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
            XNamespace content = "http://schemas.openxmlformats.org/package/2006/content-types";
            var doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            var types = new XElement(content + "Types");
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write))
            using (var zip = new ZipArchive(file, ZipArchiveMode.Create))
            {
                foreach (var part in doc.Root.Elements(pkg + "part"))
                {
                    var name = ((string)part.Attribute(pkg + "name") ?? "").TrimStart('/');
                    if (name.Length == 0 || name.Contains("..") || name.Contains('\\') || !names.Add(name)) throw new InvalidDataException("Invalid Word package part.");
                    var entry = zip.CreateEntry(name);
                    using (var stream = entry.Open())
                    {
                        var binary = part.Element(pkg + "binaryData");
                        if (binary != null) { var bytes = Convert.FromBase64String(binary.Value); stream.Write(bytes, 0, bytes.Length); }
                        else { var element = part.Element(pkg + "xmlData")?.Elements().Single(); if (element == null) throw new InvalidDataException("Missing Word XML part."); using (var writer = new StreamWriter(stream, new UTF8Encoding(false))) writer.Write(element.ToString(SaveOptions.DisableFormatting)); }
                    }
                    types.Add(new XElement(content + "Override", new XAttribute("PartName", "/" + name), new XAttribute("ContentType", (string)part.Attribute(pkg + "contentType") ?? "application/xml")));
                }
                using (var writer = new StreamWriter(zip.CreateEntry("[Content_Types].xml").Open(), new UTF8Encoding(false))) writer.Write(types.ToString());
            }
        }
    }
}
