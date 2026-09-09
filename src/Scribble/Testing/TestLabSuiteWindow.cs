using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Scribble.Testing
{
    public sealed class TestLabSuiteWindow : Form
    {
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr window, int command);
        private readonly TextBox log = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false, Dock = DockStyle.Fill };
        private readonly CancellationTokenSource cancellation = new CancellationTokenSource();
        private readonly string folder;
        private bool finished;
        public static void Open()
        {
            var active = TestLabSuite.Active();
            if (active != null) {
                File.WriteAllText(Path.Combine(TestLab.Root, "activate-" + active.id), "activate");
                using (var process = Process.GetProcessById(active.pid)) { var handle = process.MainWindowHandle; if (handle != IntPtr.Zero) { ShowWindow(handle, 9); SetForegroundWindow(handle); } }
                return;
            }
            var host = Path.Combine(Path.GetDirectoryName(typeof(TestLab).Assembly.Location), "ScribbleBrowserHost.exe");
            if (!File.Exists(host)) throw new FileNotFoundException("Install the current Scribble build to use the standalone Test Lab runner.", host);
            using (var process = Process.Start(new ProcessStartInfo(host, "--test-lab-suite") { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Normal })) { }
        }
        public TestLabSuiteWindow()
        {
            Text = "Scribble Test Lab — running all cases"; Size = new Size(960, 700); MinimumSize = new Size(720, 500);
            Font = new Font("Segoe UI", 10); StartPosition = FormStartPosition.CenterScreen;
            folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Scribble Testcases", "suite-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(folder);
            var controls = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 80, Padding = new Padding(8) };
            Add(controls, "Stop suite", () => { cancellation.Cancel(); Append("Stopping after the current operation; evidence will be saved."); });
            Add(controls, "Recover incomplete run", () => {
                if (!finished) throw new InvalidOperationException("Stop the current suite and wait for its report before recovering.");
                var report = TestLabSuite.RecoverIncomplete(folder);
                Append("Preserved unfinished capture: " + report);
                TestLabSuiteWindow.Open(); Close();
            });
            Add(controls, "Copy summary", () => Clipboard.SetText(File.Exists(Path.Combine(folder, "summary.txt")) ? File.ReadAllText(Path.Combine(folder, "summary.txt")) : log.Text));
            Add(controls, "Open results folder", () => Process.Start(new ProcessStartInfo(folder) { UseShellExecute = true }));
            Add(controls, "Open HTML report", () => { var path = Path.Combine(folder, "report.html"); if (!File.Exists(path)) throw new InvalidOperationException("The HTML report is created when the suite finishes or stops."); Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); });
            Controls.Add(log); Controls.Add(new Label { Dock = DockStyle.Top, Height = 58, Text = "Runs the synthetic suite visibly using your configured model. Prompts and outputs are recorded locally; emails remain unsent drafts.\r\nResults: " + folder }); Controls.Add(controls);
            var activation = new System.Windows.Forms.Timer { Interval = 500 };
            activation.Tick += (s, e) => {
                var active = TestLabSuite.Active();
                if (active == null || active.pid != Process.GetCurrentProcess().Id) return;
                var signal = Path.Combine(TestLab.Root, "activate-" + active.id);
                if (!File.Exists(signal)) return;
                try { File.Delete(signal); Show(); WindowState = FormWindowState.Normal; Activate(); } catch (IOException) { }
            };
            activation.Start(); FormClosed += (s, e) => activation.Dispose();
            Shown += async (s, e) => { Activate(); await Run(); };
            FormClosing += (s, e) => { if (!finished) { e.Cancel = true; cancellation.Cancel(); Append("Stopping and saving the report before closing. Keep this window open until export finishes."); } };
        }
        private void Add(FlowLayoutPanel panel, string text, Action action)
        {
            var b = new Button { Text = text, AutoSize = true }; b.Click += (s, e) => { try { action(); } catch (Exception ex) { Append(ex.Message); } }; panel.Controls.Add(b);
        }
        private void Append(string text)
        {
            if (InvokeRequired) { BeginInvoke((Action)(() => Append(text))); return; }
            log.AppendText(text + Environment.NewLine);
        }
        // Office COM stays on one STA with a message pump; the operator window remains responsive.
        private static Task RunOnSta(Func<Task> action)
        {
            var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            var thread = new Thread(() => {
                try {
                    using (var context = new ApplicationContext()) using (var dispatcher = new Control()) {
                        var handle = dispatcher.Handle;
                        SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                        dispatcher.BeginInvoke((Action)(async () => {
                            try { await action(); completion.TrySetResult(true); }
                            catch (Exception e) { completion.TrySetException(e); }
                            finally { context.ExitThread(); }
                        }));
                        Application.Run(context);
                    }
                } catch (Exception e) { completion.TrySetException(e); }
            }) { IsBackground = true, Name = "Scribble test suite" };
            thread.SetApartmentState(ApartmentState.STA); thread.Start(); return completion.Task;
        }
        private async Task Run()
        {
            Append("Test Lab runner " + FileVersionInfo.GetVersionInfo(typeof(TestLab).Assembly.Location).FileVersion);
            if (TestLab.ActiveRunId() != null) {
                Append("An unfinished capture is still active. Save your work, close Excel, PowerPoint, Word, Outlook and Chrome, then click Recover incomplete run here. The old evidence will be preserved before a fresh suite opens.");
            }
            var runner = new TestLabSuiteRunner(folder, Append, cancellation.Token);
            try { await RunOnSta(runner.Run); }
            catch (Exception e) { runner.Log("Cannot run suite: " + e); }
            try {
                Append("Creating the HTML report with copyable diagnostics and available output previews...");
                await Task.Run(() => TestLabSuiteReport.Create(runner.State, runner.Results.ToArray()));
                Append("Saved " + Path.Combine(folder, "report.html") + ". Open HTML report to copy diagnostic parts or take screenshots.");
            } catch (Exception e) {
                runner.Log("HTML export failed: " + e);
                File.AppendAllText(Path.Combine(folder, "summary.txt"), "\nHTML EXPORT FAILED: " + e.Message + "\nRecorded logs are preserved.\n");
            }
            finally { finished = true; Text = "Scribble Test Lab — finished"; }
        }
    }
    public static class TestLabSuiteReport
    {
        private static string E(string text) => WebUtility.HtmlEncode(text ?? "");
        public static string BuildHtml(SuiteState s, SuiteCaseResult[] results)
        {
            var summary = new StringBuilder("SCRIBBLE TEST LAB\nRunner build: " + FileVersionInfo.GetVersionInfo(typeof(TestLab).Assembly.Location).FileVersion + "\nSuite: " + s.id + "\nMain: " + s.commit + "\nKit SHA256: " + s.kitHash + "\nResults: " + s.folder + "\n");
            summary.AppendLine("Needs review: " + results.Count(r => r.status == "needs_review") + " | Blocked/stopped: " + results.Count(r => r.status == "blocked" || r.status == "stopped" || r.status == "incomplete") + " | Not run: " + results.Count(r => r.status == "not_run"));
            summary.AppendLine("Completion is not a correctness pass. See native output and visual review in the evidence.");
            foreach (var r in results) summary.AppendLine(r.id + " " + r.host + " — " + r.status + (string.IsNullOrEmpty(r.error) ? "" : " — " + (r.error.Split('\n')[0].Length > 140 ? r.error.Split('\n')[0].Substring(0, 140) + "…" : r.error.Split('\n')[0])));
            if (results.Length == 0) summary.AppendLine("No cases ran. See the startup/download error below.");
            var logPath = Path.Combine(s.folder, "suite.log"); var log = File.Exists(logPath) ? File.ReadAllText(logPath) : "No suite log was produced.";
            if (results.Length == 0) summary.AppendLine(log.Length > 1600 ? log.Substring(0, 1600) + "\n[Full error in the HTML log]" : log);
            var pasteable = summary.ToString();
            var firstFailure = results.FirstOrDefault(r => !string.IsNullOrEmpty(r.error));
            if (firstFailure != null) {
                var diagnostic = firstFailure.error;
                pasteable += "\nFIRST FAILURE DETAILS (" + firstFailure.id + ")\n" +
                    (diagnostic.Length > 20000 ? diagnostic.Substring(0, 20000) + "\n[Continued in report.html and suite.log]" : diagnostic);
            }
            File.WriteAllText(Path.Combine(s.folder, "summary.txt"), pasteable, new UTF8Encoding(false));
            var html = new StringBuilder("<!doctype html><html><head><meta charset='utf-8'><meta http-equiv='Content-Security-Policy' content=\"default-src 'none'; style-src 'unsafe-inline'; img-src data:\"><title>Scribble suite report</title><style>@page{size:A4;margin:14mm}body{font:11pt/1.5 Arial;color:#172033;max-width:1100px;margin:24px auto;padding:0 24px}button,summary{cursor:pointer;padding:10px}textarea{font:11pt Consolas}details{margin:10px 0;border:1px solid #ccd5e0;padding:8px}h1{font-size:22pt}h2{font-size:16pt}pre{white-space:pre-wrap;overflow-wrap:anywhere;font:11pt/1.5 Consolas}table{border-collapse:collapse;width:100%}td,th{border:1px solid #ccd5e0;padding:5px;text-align:left}img{max-width:100%}.case{break-before:page}.cover{break-after:page}</style></head><body><section class='cover'><h1>Scribble Test Lab</h1><pre>");
            html.Append(E(summary.ToString())).Append("</pre><p>Start by pasting the summary. Then use the numbered diagnostic parts below to relay logs and results in chat. Take screenshots of the relevant output previews or visible apps for layout review. No file upload is required.</p></section><h2>Timestamped suite log</h2><pre>").Append(E(log)).Append("</pre>");
            foreach (var r in results) {
                html.Append("<section class='case'><h1>").Append(E(r.id + " / " + r.host)).Append("</h1><p>").Append(E(r.status + " | " + r.started + " → " + r.finished)).Append("</p><pre>").Append(E(r.error)).Append("</pre>");
                var caseHtml = TestLab.SafeChild(s.folder, "cases/" + r.id + "/report.html");
                if (File.Exists(caseHtml)) {
                    var text = File.ReadAllText(caseHtml); var body = Regex.Match(text, @"<body[^>]*>([\s\S]*)</body>", RegexOptions.IgnoreCase);
                    html.Append(body.Success ? body.Groups[1].Value : "<pre>Case report could not be embedded. See its HTML and evidence ZIP.</pre>");
                }
                html.Append("</section>");
            }
            var report = html.Append("</body></html>").ToString();
            var diagnostics = ToText(report);
            File.WriteAllText(Path.Combine(s.folder, "diagnostics.txt"), diagnostics, new UTF8Encoding(false));
            var parts = SplitDiagnostics(s.id, diagnostics);
            var relay = new StringBuilder("<section id='relay'><h2>Copy diagnostics into chat</h2><p>Send the summary first, then these numbered parts in order when needed. Each part includes the suite ID. Click Copy part; if clipboard access is unavailable, the text is selected: press Ctrl+C. Screenshots of output previews help assess layout.</p>");
            for (int i = 0; i < parts.Length; i++) relay.Append("<details><summary>Diagnostic part ").Append(i + 1).Append(" / ").Append(parts.Length).Append("</summary><button type='button' data-copy='part-").Append(i).Append("'>Copy part ").Append(i + 1).Append("</button><textarea readonly id='part-").Append(i).Append("' style='width:100%;height:240px'>").Append(E(parts[i])).Append("</textarea></details>");
            relay.Append("<p id='copy-status' role='status'></p></section>");
            var nonce = Guid.NewGuid().ToString("N");
            report = report.Replace("default-src 'none';", "default-src 'none'; script-src 'nonce-" + nonce + "';");
            report = report.Replace("</body>", "<script nonce='" + nonce + "'>document.querySelectorAll('[data-copy]').forEach(function(b){b.addEventListener('click',async function(){var t=document.getElementById(b.getAttribute('data-copy'));t.focus();t.select();try{await navigator.clipboard.writeText(t.value);document.getElementById('copy-status').textContent='Copied '+b.textContent;}catch(e){document.getElementById('copy-status').textContent='Text selected. Press Ctrl+C to copy.';}});});</script></body>");
            return report.Replace("</section><h2>Timestamped suite log", "</section>" + relay + "<h2>Timestamped suite log");
        }
        public static string Create(SuiteState s, SuiteCaseResult[] results)
        {
            var html = Path.Combine(s.folder, "report.html");
            File.WriteAllText(html, BuildHtml(s, results), new UTF8Encoding(false)); return html;
        }
        public static string ToText(string html)
        {
            html = Regex.Replace(html, @"<(script|style)\b[^>]*>[\s\S]*?</\1>", "", RegexOptions.IgnoreCase);
            html = Regex.Replace(html, @"<img\b[^>]*>", "\n[Image preview: take a screenshot in the HTML report.]\n", RegexOptions.IgnoreCase);
            html = Regex.Replace(html, @"</?(?:h[1-6]|p|pre|section|article|tr|div|br)\b[^>]*>", "\n", RegexOptions.IgnoreCase);
            return WebUtility.HtmlDecode(Regex.Replace(html, "<[^>]*>", "")).Replace("\r\n", "\n").Replace("\r", "\n");
        }
        public static string[] SplitDiagnostics(string suiteId, string text)
        {
            const int size = 12000;
            var chunks = new System.Collections.Generic.List<string>();
            for (int offset = 0; offset < text.Length;) {
                int length = Math.Min(size, text.Length - offset);
                if (offset + length < text.Length && char.IsHighSurrogate(text[offset + length - 1])) length--;
                chunks.Add(text.Substring(offset, length)); offset += length;
            }
            if (chunks.Count == 0) chunks.Add("");
            return chunks.Select((chunk, i) => "SCRIBBLE DIAGNOSTICS | Suite " + suiteId + " | Part " + (i + 1) + "/" + chunks.Count + "\n" + chunk).ToArray();
        }
    }
}
