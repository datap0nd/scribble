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
        public static string Capture(string runId)
        {
            var run = TestLab.GetRun(runId);
            var directory = Path.Combine(TestLab.RunDirectory(runId), "capture", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            var report = new List<string>();
            foreach (var kind in new[] { "Excel", "PowerPoint", "Word" })
            {
                object instance = null;
                try
                {
                    instance = Marshal.GetActiveObject(kind + ".Application"); dynamic app = instance;
                    dynamic documents = kind == "Excel" ? app.Workbooks : kind == "PowerPoint" ? app.Presentations : app.Documents;
                    for (int i = 1; i <= (int)documents.Count; i++)
                    {
                        object value = documents.Item(i); dynamic document = value;
                        try
                        {
                            if (!TestLab.IsRunOutput(value, run.run_id) && !TestLabSuite.OwnsNativeSource(run.run_id, Convert.ToString(document.FullName), kind)) continue;
                            var stem = Path.Combine(directory, kind + "-" + i);
                            try {
                                var readback = ReadNative(value, kind);
                                File.WriteAllText(stem + "-readback.json", TestLab.Serialize(new { schema = 1, run_id = runId,
                                    host = kind, captured_utc = DateTime.UtcNow.ToString("O"), native_readback = true,
                                    run_created_output = TestLab.IsRunOutput(value, runId), text = readback }), Encoding.UTF8);
                                TestLab.Collect(runId, stem + "-readback.json");
                                report.Add("Captured " + kind + " cell/text/structure readback.");
                            } catch (Exception ex) { report.Add(kind + " readback failed: " + ex.Message); }
                            var extension = kind == "Excel" ? ".xlsx" : kind == "PowerPoint" ? ".pptx" : ".docx";
                            if (kind == "Excel") document.SaveCopyAs(stem + extension);
                            else if (kind == "PowerPoint") document.SaveCopyAs(stem + extension, 24);
                            else SaveFlatOpc(Convert.ToString(document.WordOpenXML), stem + extension);
                            var collected = TestLab.Collect(runId, stem + extension);
                            report.Add("Captured " + Path.GetFileName(collected));
                            try
                            {
                                if (kind == "Excel") document.ExportAsFixedFormat(0, stem + ".pdf");
                                else if (kind == "PowerPoint") document.SaveCopyAs(stem + ".pdf", 32);
                                else document.ExportAsFixedFormat(stem + ".pdf", 17);
                                TestLab.Collect(runId, stem + ".pdf");
                                if (kind == "PowerPoint") for (int n = 1; n <= (int)document.Slides.Count; n++)
                                {
                                    var png = stem + "-slide-" + n + ".png";
                                    document.Slides.Item(n).Export(png, "PNG", 1280, 720); TestLab.Collect(runId, png);
                                }
                            }
                            catch (Exception ex) { report.Add(kind + " native file captured; preview export failed: " + ex.GetType().Name); }
                        }
                        catch (Exception ex) { report.Add(kind + " capture: " + ex.Message); }
                        finally { if (Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
                    }
                }
                catch (COMException) { report.Add(kind + ": no accessible running app. Save and collect manually if needed."); }
                finally { if (instance != null && Marshal.IsComObject(instance)) Marshal.ReleaseComObject(instance); }
            }
            object outlookInstance = null;
            try
            {
                outlookInstance = Marshal.GetActiveObject("Outlook.Application"); dynamic outlook = outlookInstance;
                for (int i = 1; i <= (int)outlook.Inspectors.Count; i++)
                {
                    object value = outlook.Inspectors.Item(i).CurrentItem; dynamic mail = value;
                    try
                    {
                        dynamic tag = mail.UserProperties.Find("ScribbleTestRunId");
                        if (tag == null || Convert.ToString(tag.Value) != runId) continue;
                        if (Convert.ToBoolean(mail.Sent)) throw new InvalidOperationException("A run-owned email has been sent. It cannot be certified as an unsent draft.");
                        var stem = Path.Combine(directory, "Outlook-" + i);
                        mail.SaveAs(stem + ".msg", 9);
                        TestLab.Collect(runId, stem + ".msg");
                        File.WriteAllText(stem + ".html", Convert.ToString(mail.HTMLBody), Encoding.UTF8);
                        TestLab.Collect(runId, stem + ".html");
                        var attachments = new List<object>();
                        for (int n = 1; n <= (int)mail.Attachments.Count; n++)
                        {
                            dynamic attachment = mail.Attachments.Item(n);
                            var saved = Path.Combine(directory, "mail-" + i + "-attachment-" + n + ".bin");
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
            finally { if (outlookInstance != null && Marshal.IsComObject(outlookInstance)) Marshal.ReleaseComObject(outlookInstance); }
            if (report.Count == 0) report.Add("No new run-owned document found. For drafts inside a source workbook/deck, save a separate copy and use Collect saved outputs. Save Outlook drafts as MSG.");
            return string.Join(Environment.NewLine, report);
        }
        private static object At(object value, int row, int column)
        {
            var array = value as Array;
            return array == null ? value : array.GetValue(row + array.GetLowerBound(0), column + array.GetLowerBound(1));
        }
        private static string ReadNative(object value, string kind)
        {
            dynamic document = value;
            var text = new StringBuilder();
            if (kind == "Word") return Convert.ToString(document.Content.Text);
            if (kind == "Excel") {
                for (int n = 1; n <= (int)document.Worksheets.Count; n++) {
                    dynamic sheet = document.Worksheets.Item(n); dynamic used = sheet.UsedRange;
                    int rows = (int)used.Rows.Count, columns = (int)used.Columns.Count;
                    text.AppendLine("Worksheet: " + Convert.ToString(sheet.Name) + " | used range: " + Convert.ToString(used.Address));
                    if ((long)rows * columns > 100000) { text.AppendLine("READBACK LIMIT: range exceeds 100000 cells. Screenshot the relevant range in Excel."); continue; }
                    object values = used.Value2, formulas = used.Formula;
                    for (int r = 0; r < rows; r++) for (int c = 0; c < columns; c++) {
                        var cell = At(values, r, c); var formula = Convert.ToString(At(formulas, r, c));
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
                text.AppendLine("Slide count: " + (int)document.Slides.Count);
                for (int n = 1; n <= (int)document.Slides.Count; n++) {
                    dynamic slide = document.Slides.Item(n); text.AppendLine("Slide " + n);
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
