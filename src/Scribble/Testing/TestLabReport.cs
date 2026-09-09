using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;
using System.Xml.Linq;

namespace Scribble.Testing
{
    public static class TestLabReport
    {
        private static string E(string value) { return WebUtility.HtmlEncode(value ?? ""); }
        private static string Value(IDictionary<string, object> d, string key) { object v; return d != null && d.TryGetValue(key, out v) ? Convert.ToString(v, CultureInfo.InvariantCulture) : ""; }
        private static bool ErrorLike(object value)
        {
            var d = value as IDictionary<string, object>;
            if (d != null) return d.Any(p => ((p.Key == "error" || p.Key == "exception") && p.Value != null && Convert.ToString(p.Value) != "False" && Convert.ToString(p.Value).Length > 0) ||
                ((p.Key == "status" || p.Key == "outcome") && new[] { "failed", "error", "rejected", "cancelled" }.Contains(Convert.ToString(p.Value))) || ErrorLike(p.Value));
            var array = value as object[];
            return array != null && array.Any(ErrorLike);
        }
        private static string Pretty(object value)
        {
            var d = value as IDictionary<string, object>;
            if (d != null) return string.Join("\n", d.Select(p => p.Key + ": " + Pretty(p.Value)));
            var list = value as object[];
            return list == null ? Convert.ToString(value, CultureInfo.InvariantCulture) : string.Join("\n", list.Select(Pretty));
        }
        private static string Read(ZipArchive zip, string name)
        {
            var entry = zip.GetEntry(name);
            if (entry == null) return "";
            using (var reader = new StreamReader(entry.Open())) return reader.ReadToEnd();
        }
        public static string BuildHtml(string zipPath, string summaryPath)
        {
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            using (var zip = ZipFile.OpenRead(zipPath))
            {
                var run = json.Deserialize<LabRun>(Read(zip, "run.json"));
                var events = Read(zip, "timeline.jsonl").Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(line => json.Deserialize<Dictionary<string, object>>(line)).ToArray();
                var errors = events.Where(v => new[] { "error", "fail", "reject", "timeout", "cancel" }.Any(s => Value(v, "stage").IndexOf(s, StringComparison.OrdinalIgnoreCase) >= 0) || ErrorLike(v)).ToArray();
                var outputs = zip.Entries.Where(e => e.FullName.StartsWith("artifacts/", StringComparison.Ordinal) && !e.FullName.EndsWith(".receipt.json", StringComparison.Ordinal)).ToArray();
                var missing = string.Join(", ", run.missing_artifacts ?? new string[0]);
                var title = !run.trace_complete ? "INCOMPLETE CAPTURE" : missing.Length > 0 ? "MISSING OUTPUTS" : "READY FOR REVIEW";
                var summary = "SCRIBBLE TEST REPORT\n" + title + " - correctness not yet reviewed\nCase: " + run.case_id + " | App: " + run.host +
                    "\nRun: " + run.run_id + "\nBuild: " + run.assembly_version + " | Model: " + run.selected_model +
                    "\nStarted UTC: " + run.started_utc + "\nFinished UTC: " + (run.finished_utc ?? "Not finished") +
                    "\nTrace: " + (run.trace_complete ? "complete" : "INCOMPLETE") + " | Assisted: " + run.assisted +
                    "\nEvents: " + events.Length + " | Error-like events: " + errors.Length + " | Collected outputs: " + outputs.Length +
                    "\nMissing types: " + (missing.Length == 0 ? "none" : missing) + "\n";
                foreach (var error in errors.Take(3)) {
                    var text = Pretty(error.ContainsKey("detail") ? error["detail"] : error).Replace('\n', ' ').Replace('\r', ' ');
                    summary += "\n" + Value(error, "utc") + " | " + Value(error, "stage") + " | " + (text.Length > 220 ? text.Substring(0, 220) + "... [full event in PDF]" : text);
                }
                summary += "\n\nOutputs: " + string.Join(", ", outputs.Select(e => Path.GetFileName(e.FullName))) +
                    "\n\nPlease diagnose this run using the attached PDF. The ZIP contains the full machine trace and original outputs.\n";
                File.WriteAllText(summaryPath, summary, new UTF8Encoding(false));
                var h = new StringBuilder("<!doctype html><html><head><meta charset='utf-8'><meta http-equiv='Content-Security-Policy' content=\"default-src 'none'; style-src 'unsafe-inline'; img-src data:\"><title>Scribble test report</title><style>");
                h.Append("@page{size:A4;margin:16mm}*{box-sizing:border-box}body{font:11px/1.45 Arial,sans-serif;color:#173047;margin:0}h1{font-size:27px;margin:0 0 8px}h2{font-size:18px;color:#16717a;border-bottom:1px solid #bed5d9;padding-bottom:6px;margin-top:24px}h3{font-size:12px;margin-bottom:6px}pre{white-space:pre-wrap;overflow-wrap:anywhere;word-break:break-word;font:10px/1.45 Consolas,monospace;margin:8px 0}table{width:100%;border-collapse:collapse;table-layout:fixed}td,th{padding:7px;border:1px solid #ccdbe0;text-align:left;overflow-wrap:anywhere;vertical-align:top}th{background:#edf4f5}img{max-width:100%;max-height:230mm;object-fit:contain}section{break-before:page}article{margin:14px 0}.badge{padding:12px;background:#edf4f5;font-size:16px;font-weight:bold}.muted{color:#587080}.cover pre{font:12px/1.5 Arial,sans-serif}.preview{break-inside:avoid}thead{display:table-header-group}</style></head><body>");
                h.Append("<div class='cover'><p class='muted'>SCRIBBLE / SYNTHETIC TEST EVIDENCE</p><h1>" + E(run.case_id) + " test report</h1><div class='badge'>" + E(title) + "</div><pre>" + E(summary) + "</pre><p>Screenshot this page or paste the adjacent summary.txt. No pass is inferred from file presence or model claims. Error-like events are diagnostic flags; the complete recorded timeline is included below.</p></div>");
                h.Append("<section><h2>Prompt and expected result</h2><pre>" + E(run.prompt) + "</pre>");
                try {
                    var cases = json.Deserialize<Dictionary<string, object>[]>(File.ReadAllText(TestLab.SafeChild(run.fixture_root, "operator/cases.json")));
                    var c = cases.FirstOrDefault(v => Value(v, "id") == run.case_id);
                    if (c != null) h.Append("<h3>Expected</h3><pre>" + E(Value(c, "expected")) + "</pre><h3>Setup</h3><pre>" + E(Value(c, "setup")) + "</pre>");
                } catch { h.Append("<p>Original kit is unavailable. Use the case rubric from the downloaded kit.</p>"); }
                h.Append("<h3>Capture limitations</h3><pre>" + E("Native output correctness and visual review: pending\nAssembly SHA-256: " + run.assembly_sha256 + "\nKit SHA-256: " + run.manifest_sha256) + "</pre>");
                foreach (var entry in zip.Entries.Where(e => e.FullName.StartsWith("incomplete-", StringComparison.Ordinal))) h.Append("<pre>" + E(Read(zip, entry.FullName)) + "</pre>");
                h.Append("<h3>Video markers (UTC)</h3><pre>" + E(Read(zip, "video-markers.csv")) + "</pre></section><section><h2>Errors and interruptions</h2>");
                if (errors.Length == 0) h.Append("<p>No error-like events were detected. This does not establish a successful result.</p>");
                foreach (var v in errors) Event(h, v);
                h.Append("</section><section><h2>Results and collected outputs</h2><p>These are actual collected files and model responses, not reference answers. Text previews do not verify formulas or native rendering. Original editable files and PDFs are retained in the evidence ZIP.</p>");
                foreach (var v in events.Where(v => { object detail; return v.TryGetValue("detail", out detail) && Value(detail as IDictionary<string, object>, "type") == "assistant"; })) Event(h, v);
                if (outputs.Length == 0) h.Append("<p>No output files were collected.</p>");
                foreach (var entry in outputs)
                {
                    h.Append("<article><h3>" + E(entry.FullName) + "</h3><p>" + entry.Length + " bytes</p><pre>" + E(Read(zip, entry.FullName + ".receipt.json")) + "</pre>");
                    if (entry.FullName.EndsWith(".png", StringComparison.OrdinalIgnoreCase)) {
                        using (var memory = new MemoryStream()) { using (var stream = entry.Open()) stream.CopyTo(memory); h.Append("<div class='preview'><img src='data:image/png;base64," + Convert.ToBase64String(memory.ToArray()) + "'></div>"); }
                    } else if (new[] { ".xlsx", ".pptx", ".docx" }.Contains(Path.GetExtension(entry.FullName).ToLowerInvariant())) {
                        try { h.Append("<pre>" + E(NativePreview(entry)) + "</pre>"); } catch (Exception e) { h.Append("<p>Preview unavailable: " + E(e.Message) + ". Inspect the original file in the ZIP.</p>"); }
                    } else if (entry.FullName.EndsWith(".json", StringComparison.OrdinalIgnoreCase) || entry.FullName.EndsWith(".eml", StringComparison.OrdinalIgnoreCase)) {
                        h.Append("<pre>" + E(Read(zip, entry.FullName)) + "</pre>");
                    } else { h.Append("<p>Open the original file in the ZIP for its full native or rendered result.</p>"); }
                    h.Append("</article>");
                }
                h.Append("</section><section><h2>Complete recorded timeline</h2><p>All retained events are included, with UTC timestamps, task IDs and per-process sequence numbers. No event body is shortened here. Any capture gaps are listed above.</p>");
                foreach (var v in events) Event(h, v);
                h.Append("</section></body></html>");
                return h.ToString();
            }
        }
        private static void Event(StringBuilder h, IDictionary<string, object> v)
        {
            h.Append("<article><h3>" + E(Value(v, "utc") + " | " + Value(v, "stage")) + "</h3><p class='muted'>" + E("Task " + Value(v, "task_id") + " | Instance " + Value(v, "instance_id") + " | Sequence " + Value(v, "sequence")) + "</p><pre>" + E(Pretty(v)) + "</pre></article>");
        }
        private static string NativePreview(ZipArchiveEntry entry)
        {
            using (var memory = new MemoryStream())
            {
                using (var input = entry.Open()) input.CopyTo(memory); memory.Position = 0;
                using (var package = new ZipArchive(memory, ZipArchiveMode.Read))
                {
                    if (package.Entries.Sum(e => e.Length) > 100L * 1024 * 1024) return "Preview omitted: expanded Office package exceeds 100 MB. Original retained in ZIP.";
                    var text = new StringBuilder();
                    foreach (var part in package.Entries.Where(e => e.FullName == "word/document.xml" ||
                        System.Text.RegularExpressions.Regex.IsMatch(e.FullName, @"^ppt/(slides/slide\d+|notesSlides/notesSlide\d+)\.xml$")))
                    {
                        using (var stream = part.Open()) { var doc = XDocument.Load(stream); text.AppendLine(part.FullName); text.AppendLine(string.Join(" ", doc.Descendants().Where(n => n.Name.LocalName == "t").Select(n => n.Value))); }
                    }
                    var shared = package.GetEntry("xl/sharedStrings.xml"); string[] strings = new string[0];
                    if (shared != null) using (var stream = shared.Open()) strings = XDocument.Load(stream).Descendants().Where(n => n.Name.LocalName == "si").Select(n => string.Concat(n.Descendants().Where(x => x.Name.LocalName == "t").Select(x => x.Value))).ToArray();
                    foreach (var part in package.Entries.Where(e => System.Text.RegularExpressions.Regex.IsMatch(e.FullName, @"^xl/worksheets/sheet\d+\.xml$")))
                    {
                        text.AppendLine(part.FullName + " - first 200 rows; formulas are shown, not recalculated");
                        using (var stream = part.Open()) foreach (var row in XDocument.Load(stream).Descendants().Where(n => n.Name.LocalName == "row").Take(200))
                            foreach (var cell in row.Elements().Where(n => n.Name.LocalName == "c")) {
                                var value = string.Join(" ", cell.Descendants().Where(n => n.Name.LocalName == "v" || n.Name.LocalName == "t").Select(n => n.Value));
                                int index; if ((string)cell.Attribute("t") == "s" && int.TryParse(value, out index) && index >= 0 && index < strings.Length) value = strings[index];
                                var formula = cell.Elements().FirstOrDefault(n => n.Name.LocalName == "f");
                                text.AppendLine((string)cell.Attribute("r") + ": " + value + (formula == null ? "" : " [=" + formula.Value + "]"));
                            }
                    }
                    return text.ToString();
                }
            }
        }
        public static string Create(string zipPath)
        {
            zipPath = Path.GetFullPath(zipPath);
            var prefix = Path.Combine(Path.GetDirectoryName(zipPath), Path.GetFileNameWithoutExtension(zipPath));
            var summary = prefix + "-summary.txt"; var html = prefix + "-report.html"; var pdf = prefix + "-report.pdf";
            File.WriteAllText(html, BuildHtml(zipPath, summary), new UTF8Encoding(false));
            var browsers = new[] {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"Google\Chrome\Application\chrome.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Google\Chrome\Application\chrome.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Microsoft\Edge\Application\msedge.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"Microsoft\Edge\Application\msedge.exe") }.Where(File.Exists).Distinct().ToArray();
            if (browsers.Length == 0) throw new InvalidOperationException("PDF export needs Microsoft Edge or Chrome. The evidence ZIP, HTML report and pasteable summary were saved at " + prefix + ".");
            bool rendered = false; var rendererLog = new StringBuilder();
            foreach (var browser in browsers) {
                if (File.Exists(pdf)) File.Delete(pdf);
                var profile = Path.Combine(TestLab.Root, "report-browser", Guid.NewGuid().ToString("N"));
                var args = "--headless --disable-gpu --disable-extensions --disable-background-networking --no-first-run --no-default-browser-check --no-pdf-header-footer --user-data-dir=\"" + profile + "\" --print-to-pdf=\"" + pdf + "\" \"" + new Uri(html).AbsoluteUri + "\"";
                try {
                    using (var process = Process.Start(new ProcessStartInfo(browser, args) { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden, RedirectStandardError = true })) {
                        var stderr = process.StandardError.ReadToEndAsync();
                        if (!process.WaitForExit(60000)) { process.Kill(); rendererLog.AppendLine(browser + ": exceeded 60 seconds"); continue; }
                        rendererLog.AppendLine(browser + ": exit " + process.ExitCode + "\n" + stderr.GetAwaiter().GetResult());
                        rendered = process.ExitCode == 0 && File.Exists(pdf) && new FileInfo(pdf).Length > 100;
                    }
                } catch (Exception e) { rendererLog.AppendLine(browser + ": " + e.Message); }
                if (rendered) break;
            }
            File.WriteAllText(prefix + "-report-render.log", rendererLog.ToString());
            if (!rendered) throw new IOException("PDF rendering failed. The HTML report, summary, renderer log and evidence ZIP are preserved at " + prefix + ".");
            using (var zip = ZipFile.Open(zipPath, ZipArchiveMode.Update))
            {
                foreach (var pair in new[] { new[] { html, "report.html" }, new[] { summary, "summary.txt" }, new[] { pdf, "report.pdf" } }) {
                    var entry = zip.GetEntry(pair[1]) ?? zip.CreateEntry(pair[1]);
                    using (var stream = entry.Open()) { stream.SetLength(0); using (var file = File.OpenRead(pair[0])) file.CopyTo(stream); }
                }
                var files = zip.Entries.Where(e => e.FullName != "export-manifest.json").Select(e => {
                    using (var stream = e.Open()) using (var sha = System.Security.Cryptography.SHA256.Create()) return new {
                        path = e.FullName, size = stream.Length, sha256 = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant() };
                }).ToArray();
                using (var stream = zip.GetEntry("export-manifest.json").Open()) { stream.SetLength(0); using (var writer = new StreamWriter(stream, new UTF8Encoding(false))) writer.Write(TestLab.Serialize(new { schema = 1, files })); }
            }
            return pdf;
        }
    }
}
