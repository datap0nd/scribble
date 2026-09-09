using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Scribble.Testing
{
    public sealed class TestLabSuiteWindow : Form
    {
        private static TestLabSuiteWindow window;
        private readonly TextBox log = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false, Dock = DockStyle.Fill };
        private readonly CancellationTokenSource cancellation = new CancellationTokenSource();
        private readonly string folder;
        private bool finished;
        public static void Open()
        {
            if (window != null && !window.IsDisposed) { window.Show(); window.Activate(); return; }
            window = new TestLabSuiteWindow(); window.Show();
        }
        public TestLabSuiteWindow()
        {
            Text = "Scribble Test Lab — running all cases"; Size = new Size(960, 700); MinimumSize = new Size(720, 500);
            Font = new Font("Segoe UI", 10); StartPosition = FormStartPosition.CenterScreen;
            folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Scribble Testcases", "suite-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(folder);
            var controls = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 80, Padding = new Padding(8) };
            Add(controls, "Stop suite", () => { cancellation.Cancel(); Append("Stopping after the current operation; evidence will be saved."); });
            Add(controls, "Copy summary", () => Clipboard.SetText(File.Exists(Path.Combine(folder, "summary.txt")) ? File.ReadAllText(Path.Combine(folder, "summary.txt")) : log.Text));
            Add(controls, "Open results folder", () => Process.Start(new ProcessStartInfo(folder) { UseShellExecute = true }));
            Add(controls, "Open final PDF", () => { var path = Path.Combine(folder, "report.pdf"); if (!File.Exists(path)) throw new InvalidOperationException("The PDF is created when the suite finishes or stops."); Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); });
            Controls.Add(log); Controls.Add(new Label { Dock = DockStyle.Top, Height = 58, Text = "Runs the synthetic suite visibly using your configured model. Prompts and outputs are recorded locally; emails remain unsent drafts.\r\nResults: " + folder }); Controls.Add(controls);
            Shown += async (s, e) => await Run();
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
            var runner = new TestLabSuiteRunner(folder, Append, cancellation.Token);
            try { await RunOnSta(runner.Run); }
            catch (Exception e) { runner.Log("Cannot run suite: " + e); }
            try {
                Append("Creating the final PDF, including errors and all available case evidence...");
                await Task.Run(() => TestLabSuiteReport.Create(runner.State, runner.Results.ToArray()));
                Append("Saved " + Path.Combine(folder, "report.pdf") + ". Use Copy summary to relay the result.");
            } catch (Exception e) {
                runner.Log("PDF export failed: " + e);
                File.AppendAllText(Path.Combine(folder, "summary.txt"), "\nPDF EXPORT FAILED: " + e.Message + "\nFull logs and HTML are preserved.\n");
            }
            finally { finished = true; Text = "Scribble Test Lab — finished"; }
        }
    }
    public static class TestLabSuiteReport
    {
        private static string E(string text) => WebUtility.HtmlEncode(text ?? "");
        public static string BuildHtml(SuiteState s, SuiteCaseResult[] results)
        {
            var summary = new StringBuilder("SCRIBBLE TEST LAB\nSuite: " + s.id + "\nMain: " + s.commit + "\nKit SHA256: " + s.kitHash + "\nResults: " + s.folder + "\n");
            summary.AppendLine("Needs review: " + results.Count(r => r.status == "needs_review") + " | Blocked/stopped: " + results.Count(r => r.status == "blocked" || r.status == "stopped" || r.status == "incomplete") + " | Not run: " + results.Count(r => r.status == "not_run"));
            summary.AppendLine("Completion is not a correctness pass. See native output and visual review in the evidence.");
            foreach (var r in results) summary.AppendLine(r.id + " " + r.host + " — " + r.status + (string.IsNullOrEmpty(r.error) ? "" : " — " + (r.error.Split('\n')[0].Length > 140 ? r.error.Split('\n')[0].Substring(0, 140) + "…" : r.error.Split('\n')[0])));
            if (results.Length == 0) summary.AppendLine("No cases ran. See the startup/download error below.");
            var logPath = Path.Combine(s.folder, "suite.log"); var log = File.Exists(logPath) ? File.ReadAllText(logPath) : "No suite log was produced.";
            if (results.Length == 0) summary.AppendLine(log.Length > 1600 ? log.Substring(0, 1600) + "\n[Full error in the PDF log]" : log);
            File.WriteAllText(Path.Combine(s.folder, "summary.txt"), summary.ToString(), new UTF8Encoding(false));
            var html = new StringBuilder("<!doctype html><html><head><meta charset='utf-8'><meta http-equiv='Content-Security-Policy' content=\"default-src 'none'; style-src 'unsafe-inline'; img-src data:\"><title>Scribble suite report</title><style>@page{size:A4;margin:14mm}body{font:10pt Arial;color:#172033}h1{font-size:22pt}h2{font-size:16pt}pre{white-space:pre-wrap;overflow-wrap:anywhere;font:8pt Consolas}table{border-collapse:collapse;width:100%}td,th{border:1px solid #ccd5e0;padding:5px;text-align:left}img{max-width:100%}.case{break-before:page}.cover{break-after:page}</style></head><body><section class='cover'><h1>Scribble Test Lab</h1><pre>");
            html.Append(E(summary.ToString())).Append("</pre><p>Share summary.txt by copy/paste, a screenshot of this page, or this PDF. Full machine-readable events and native outputs are in each case’s evidence ZIP.</p></section><h2>Timestamped suite log</h2><pre>").Append(E(log)).Append("</pre>");
            foreach (var r in results) {
                html.Append("<section class='case'><h1>").Append(E(r.id + " / " + r.host)).Append("</h1><p>").Append(E(r.status + " | " + r.started + " → " + r.finished)).Append("</p><pre>").Append(E(r.error)).Append("</pre>");
                var caseHtml = TestLab.SafeChild(s.folder, "cases/" + r.id + "/report.html");
                if (File.Exists(caseHtml)) {
                    var text = File.ReadAllText(caseHtml); var body = Regex.Match(text, @"<body[^>]*>([\s\S]*)</body>", RegexOptions.IgnoreCase);
                    html.Append(body.Success ? body.Groups[1].Value : "<pre>Case report could not be embedded. See its HTML and evidence ZIP.</pre>");
                }
                html.Append("</section>");
            }
            return html.Append("</body></html>").ToString();
        }
        public static string Create(SuiteState s, SuiteCaseResult[] results)
        {
            var html = Path.Combine(s.folder, "report.html"); var pdf = Path.Combine(s.folder, "report.pdf");
            File.WriteAllText(html, BuildHtml(s, results), new UTF8Encoding(false)); TestLabReport.RenderHtml(html, pdf); return pdf;
        }
    }
}
