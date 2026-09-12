using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace Scribble.Testing
{
    // Pure reporting boundary: this class never starts Office, Chrome, or a model.
    // It consumes only the sealed suite HTML, log, and exported evidence archives.
    public static class TestLabPdfWriter
    {
        public static void CreateText(string html, string pdf)
        {
            var temporary = pdf + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                using (var document = new PdfDocument())
                using (var writer = new PageWriter(document))
                {
                    document.Info.Title = "Scribble Test Lab evidence";
                    writer.Heading("Scribble Test Lab evidence", 20);
                    writer.Text(TestLabSuiteReport.ToText(File.ReadAllText(html)));
                    writer.ClosePage(); document.Save(temporary);
                }
                Validate(temporary);
                if (File.Exists(pdf)) File.Replace(temporary, pdf, null); else new FileInfo(temporary).MoveTo(pdf);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        public static string CreateRecovered(string evidence, string destination)
        {
            LabRun run;
            using (var archive = ZipFile.OpenRead(evidence))
            using (var reader = new StreamReader(archive.GetEntry("run.json").Open()))
                run = new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.Deserialize<LabRun>(reader.ReadToEnd());
            Directory.CreateDirectory(destination);
            File.WriteAllText(Path.Combine(destination, "report.html"),
                TestLabReport.BuildHtml(evidence, Path.Combine(destination, "summary.txt")), new UTF8Encoding(false));
            var state = new SuiteState { id = run.run_id, folder = destination, commit = "recovered capture",
                kitHash = run.manifest_sha256, caseId = run.case_id, host = run.host, runId = run.run_id };
            var result = new SuiteCaseResult { id = run.case_id, host = run.host, status = "incomplete",
                started = run.started_utc, finished = run.finished_utc, evidence = evidence,
                error = "Recovered after an interrupted operator run; correctness was not evaluated." };
            return CreateSuite(state, new[] { result });
        }

        public static string CreateSuite(SuiteState state, SuiteCaseResult[] results)
        {
            if (state == null) throw new ArgumentNullException("state");
            if (results == null) results = new SuiteCaseResult[0];
            Directory.CreateDirectory(state.folder);
            var html = Path.Combine(state.folder, "report.html");
            if (!File.Exists(html))
                File.WriteAllText(html, TestLabSuiteReport.BuildHtml(state, results), new UTF8Encoding(false));
            var pdf = Path.Combine(state.folder, "report.pdf");
            var temporary = pdf + "." + Guid.NewGuid().ToString("N") + ".tmp";
            var staging = Path.Combine(state.folder, ".report-staging-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(staging);
            try
            {
                using (var document = new PdfDocument())
                using (var writer = new PageWriter(document))
                {
                    document.Info.Title = "Scribble Test Lab " + state.id;
                    document.Info.Subject = "Synthetic suite evidence and retained diagnostics";
                    writer.Heading("Scribble Test Lab", 22);
                    writer.Badge(TerminalLabel(results));
                    writer.Text("Suite: " + state.id + "\nRunner build: " +
                        System.Diagnostics.FileVersionInfo.GetVersionInfo(typeof(TestLab).Assembly.Location).FileVersion +
                        "\nMain: " + (state.commit ?? "not recorded") +
                        "\nKit SHA256: " + (state.kitHash ?? "not recorded") +
                        "\nGenerated UTC: " + DateTime.UtcNow.ToString("O") +
                        "\nResults folder: " + state.folder);
                    writer.Text("This report contains the actual retained model responses, native readback, complete captured diagnostics, and every usable native visual derivative. File presence is not a correctness pass.");
                    writer.PageBreak();
                    writer.Heading("Complete suite report and diagnostic trace", 16);
                    writer.Text(TestLabSuiteReport.ToText(File.ReadAllText(html)));

                    foreach (var result in results)
                    {
                        if (string.IsNullOrEmpty(result.evidence) || !File.Exists(result.evidence)) continue;
                        AppendVisualEvidence(document, writer, result, staging);
                    }
                    writer.ClosePage();
                    document.Save(temporary);
                }
                Validate(temporary);
                if (File.Exists(pdf)) File.Replace(temporary, pdf, null); else new FileInfo(temporary).MoveTo(pdf);
                File.WriteAllText(Path.Combine(state.folder, "last-report.txt"), pdf, new UTF8Encoding(false));
                return pdf;
            }
            finally
            {
                if (File.Exists(temporary)) File.Delete(temporary);
                try { Directory.Delete(staging, true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }

        public static bool IsValid(string path)
        {
            try { Validate(path); return true; } catch { return false; }
        }

        private static void Validate(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path) || new FileInfo(path).Length < 1000)
                throw new InvalidDataException("The final PDF is missing or empty.");
            using (var input = PdfReader.Open(path, PdfDocumentOpenMode.Import))
                if (input.PageCount == 0) throw new InvalidDataException("The final PDF has no readable pages.");
        }

        private static string TerminalLabel(SuiteCaseResult[] results)
        {
            if (results.Length == 0) return "BLOCKED — NO CASES RAN";
            if (results.Any(r => r.status == "stopped")) return "STOPPED — PARTIAL EVIDENCE PRESERVED";
            if (results.Any(r => r.status == "failed")) return "DETERMINISTIC CHECKS FAILED — REVIEW FINDINGS";
            if (results.Any(r => r.status == "blocked" || r.status == "incomplete")) return "COMPLETED WITH EVIDENCE GAPS";
            return "READY FOR REVIEW — CORRECTNESS NOT IMPLIED";
        }

        private static void AppendVisualEvidence(PdfDocument document, PageWriter writer,
            SuiteCaseResult result, string staging)
        {
            writer.PageBreak();
            writer.Heading(result.id + " / " + result.host + " — native visual evidence", 16);
            using (var archive = ZipFile.OpenRead(result.evidence))
            {
                var visuals = archive.Entries.Where(e => e.FullName.StartsWith("artifacts/", StringComparison.Ordinal) &&
                    (e.FullName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ||
                     e.FullName.EndsWith(".png", StringComparison.OrdinalIgnoreCase))).ToArray();
                if (visuals.Length == 0)
                {
                    writer.Text("No native PDF or PNG derivative was retained for this case. See the preceding artifact inventory and capture error for the precise evidence gap.");
                    return;
                }
                foreach (var entry in visuals)
                {
                    var target = Path.Combine(staging, Guid.NewGuid().ToString("N") + Path.GetExtension(entry.Name));
                    using (var source = entry.Open()) using (var file = new FileStream(target, FileMode.CreateNew)) source.CopyTo(file);
                    writer.Text("Artifact: " + entry.FullName + " (" + entry.Length + " bytes)");
                    try
                    {
                        if (entry.FullName.EndsWith(".png", StringComparison.OrdinalIgnoreCase))
                            writer.Image(target);
                        else
                        {
                            writer.ClosePage();
                            using (var imported = PdfReader.Open(target, PdfDocumentOpenMode.Import))
                                for (int i = 0; i < imported.PageCount; i++) document.AddPage(imported.Pages[i]);
                            writer.ResetPage();
                        }
                    }
                    catch (Exception error)
                    {
                        writer.Text("VISUAL EVIDENCE GAP: " + error.GetType().Name + ": " + error.Message);
                    }
                }
            }
        }

        private sealed class PageWriter : IDisposable
        {
            private readonly PdfDocument document;
            private readonly XFont body = new XFont("Segoe UI", 8.5, XFontStyleEx.Regular,
                new XPdfFontOptions(PdfFontEncoding.Unicode));
            private readonly XFont heading = new XFont("Segoe UI", 16, XFontStyleEx.Bold,
                new XPdfFontOptions(PdfFontEncoding.Unicode));
            private readonly XFont badge = new XFont("Segoe UI", 11, XFontStyleEx.Bold,
                new XPdfFontOptions(PdfFontEncoding.Unicode));
            private PdfPage page;
            private XGraphics graphics;
            private double y;
            private const double Margin = 40;

            internal PageWriter(PdfDocument document) { this.document = document; }

            internal void Heading(string value, double size)
            {
                EnsurePage(36);
                var font = Math.Abs(size - 16) < .1 ? heading : new XFont("Segoe UI", size,
                    XFontStyleEx.Bold, new XPdfFontOptions(PdfFontEncoding.Unicode));
                graphics.DrawString(Clean(value), font, XBrushes.DarkSlateGray,
                    new XRect(Margin, y, page.Width.Point - Margin * 2, size + 8), XStringFormats.TopLeft);
                y += size + 12;
            }

            internal void Badge(string value)
            {
                EnsurePage(34);
                graphics.DrawRectangle(new XSolidBrush(XColor.FromArgb(229, 241, 242)),
                    Margin, y, page.Width.Point - Margin * 2, 28);
                graphics.DrawString(Clean(value), badge, XBrushes.DarkSlateGray,
                    new XRect(Margin + 8, y + 6, page.Width.Point - Margin * 2 - 16, 18), XStringFormats.TopLeft);
                y += 38;
            }

            internal void Text(string value)
            {
                foreach (var logical in Clean(value).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
                {
                    foreach (var line in Wrap(logical))
                    {
                        EnsurePage(13);
                        graphics.DrawString(line, body, XBrushes.Black,
                            new XRect(Margin, y, page.Width.Point - Margin * 2, 12), XStringFormats.TopLeft);
                        y += 11.5;
                    }
                    y += 2;
                }
                y += 4;
            }

            internal void Image(string path)
            {
                using (var source = System.Drawing.Image.FromFile(path))
                using (var image = XImage.FromFile(path))
                {
                    var width = page == null ? 515 : page.Width.Point - Margin * 2;
                    var height = width * source.Height / source.Width;
                    if (height > 690) { height = 690; width = height * source.Width / source.Height; }
                    EnsurePage(height + 12);
                    graphics.DrawImage(image, Margin, y, width, height);
                    y += height + 12;
                }
            }

            internal void PageBreak() { ClosePage(); }
            internal void ResetPage() { page = null; graphics = null; y = 0; }

            internal void ClosePage()
            {
                if (graphics != null) { graphics.Dispose(); graphics = null; }
                page = null; y = 0;
            }

            private void EnsurePage(double needed)
            {
                if (page == null || y + needed > page.Height.Point - Margin)
                {
                    ClosePage();
                    page = document.AddPage();
                    page.Size = PdfSharp.PageSize.A4;
                    graphics = XGraphics.FromPdfPage(page);
                    y = Margin;
                }
            }

            private IEnumerable<string> Wrap(string value)
            {
                if (value.Length == 0) return new[] { "" };
                var width = page == null ? 515 : page.Width.Point - Margin * 2;
                var lines = new List<string>(); var current = new StringBuilder();
                foreach (var word in value.Split(new[] { ' ', '\t' }, StringSplitOptions.None))
                {
                    var candidate = current.Length == 0 ? word : current + " " + word;
                    if (graphics == null) EnsurePage(13);
                    if (graphics.MeasureString(candidate, body).Width <= width) { current.Clear(); current.Append(candidate); continue; }
                    if (current.Length > 0) { lines.Add(current.ToString()); current.Clear(); }
                    var fragment = new StringBuilder();
                    foreach (var character in word)
                    {
                        if (graphics.MeasureString(fragment.ToString() + character, body).Width > width && fragment.Length > 0)
                        { lines.Add(fragment.ToString()); fragment.Clear(); }
                        fragment.Append(character);
                    }
                    current.Append(fragment);
                }
                if (current.Length > 0) lines.Add(current.ToString());
                return lines;
            }

            private static string Clean(string value)
            {
                if (string.IsNullOrEmpty(value)) return "";
                return new string(value.Select(c => c == '\n' || c == '\r' || c == '\t' || !char.IsControl(c) ? c : ' ').ToArray());
            }

            public void Dispose()
            {
                ClosePage();
            }
        }
    }
}
