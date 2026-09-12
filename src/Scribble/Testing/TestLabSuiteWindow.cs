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
using Scribble.Configuration;

namespace Scribble.Testing
{
    public sealed class TestLabSuiteWindow : Form
    {
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr window, int command);
        private static string WindowDescriptor => Path.Combine(TestLab.Root, "suite-window.json");
        private readonly TextBox log = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false, Dock = DockStyle.Fill };
        private readonly Label status = new Label { AutoSize = true };
        private readonly Label progress = new Label { AutoSize = true };
        private readonly Label currentCase = new Label { AutoSize = true };
        private readonly Label elapsed = new Label { AutoSize = true };
        private CancellationTokenSource cancellation;
        private string folder;
        private string finalPdf;
        private readonly string nonce;
        private DateTime runStarted;
        private TestLabSuiteRunner activeRunner;
        private readonly Button startButton;
        private readonly Button stopButton;
        private readonly Button reportButton;
        private readonly TextBox caseFilter = new TextBox { Width = 90 };
        private bool running;
        private bool finalizing;

        public static void Open()
        {
            Directory.CreateDirectory(TestLab.Root);
            using (var launch = new Mutex(false, @"Local\ScribbleTestLabWindow"))
            {
                bool held = false;
                try
                {
                    try { held = launch.WaitOne(TimeSpan.FromSeconds(10)); }
                    catch (AbandonedMutexException) { held = true; }
                    if (!held) throw new TimeoutException("Another Test Lab launch did not release ownership within 10 seconds. See " + TestLab.Root + ".");
                    if (ActivateExisting()) return;
                    var host = Path.Combine(Path.GetDirectoryName(typeof(TestLab).Assembly.Location), "ScribbleBrowserHost.exe");
                    if (!File.Exists(host)) throw new FileNotFoundException("Install the current Scribble build to use the standalone Test Lab runner.", host);
                    var launchNonce = Guid.NewGuid().ToString("N");
                    using (var process = Process.Start(new ProcessStartInfo(host, "--test-lab-suite " + launchNonce)
                    { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Normal }))
                    {
                        var started = DateTime.UtcNow;
                        while (DateTime.UtcNow - started < TimeSpan.FromSeconds(10))
                        {
                            process.Refresh();
                            if (process.HasExited) throw new InvalidOperationException("Test Lab exited before window_ready (exit " + process.ExitCode + ", 0x" + process.ExitCode.ToString("X8") + ").");
                            var ready = ReadWindow();
                            if (ready != null && ready.nonce == launchNonce && ready.pid == process.Id && ready.ready) return;
                            Thread.Sleep(100);
                        }
                        throw new TimeoutException("Test Lab did not acknowledge window_ready within 10 seconds. Executable: " + host);
                    }
                }
                catch (Exception error)
                {
                    var detail = "Stage: standalone runner launch\r\n" + error.Message + "\r\nDiagnostics: " + Path.Combine(TestLab.Root, "launcher-error.log");
                    File.WriteAllText(Path.Combine(TestLab.Root, "launcher-error.log"), DateTime.UtcNow.ToString("O") + " " + error);
                    MessageBox.Show(detail, "Scribble Test Lab startup error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally { if (held) launch.ReleaseMutex(); }
            }
        }

        private static bool ActivateExisting()
        {
            var existing = ReadWindow();
            if (existing == null) return false;
            try
            {
                using (var process = Process.GetProcessById(existing.pid))
                {
                    if (process.HasExited || process.StartTime.ToUniversalTime().Ticks != existing.processStart) return false;
                    if (existing.module != typeof(TestLab).Assembly.ManifestModule.ModuleVersionId.ToString())
                        throw new InvalidOperationException("A Test Lab window from a different Scribble build is already running. Close it before launching this build.");
                    File.WriteAllText(Path.Combine(TestLab.Root, "activate-window-" + existing.nonce), "activate");
                    var handle = process.MainWindowHandle;
                    if (handle != IntPtr.Zero) { ShowWindow(handle, 9); SetForegroundWindow(handle); }
                    return true;
                }
            }
            catch (ArgumentException) { return false; }
        }

        private static SuiteWindowState ReadWindow()
        {
            try { return File.Exists(WindowDescriptor) ? TestLabSuite.Read<SuiteWindowState>(WindowDescriptor) : null; }
            catch { return null; }
        }

        public static bool HasLiveWindow(int exceptPid)
        {
            var existing = ReadWindow();
            if (existing == null || existing.pid == exceptPid) return false;
            try { using (var process = Process.GetProcessById(existing.pid)) return !process.HasExited && process.StartTime.ToUniversalTime().Ticks == existing.processStart; }
            catch (ArgumentException) { return false; }
        }

        public TestLabSuiteWindow() : this(Guid.NewGuid().ToString("N")) { }

        public TestLabSuiteWindow(string launchNonce)
        {
            if (!Regex.IsMatch(launchNonce ?? "", "^[a-f0-9]{32}$")) throw new ArgumentException("Invalid Test Lab launch nonce.");
            nonce = launchNonce;
            Text = "Scribble Test Lab — idle"; Size = new Size(960, 700); MinimumSize = new Size(720, 500);
            Font = new Font("Segoe UI", 10); StartPosition = FormStartPosition.CenterScreen;
            var actions = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 58, Padding = new Padding(8) };
            startButton = Add(actions, "Start", async () => await StartRun());
            stopButton = Add(actions, "Stop", Stop);
            reportButton = Add(actions, "View final PDF", ViewReport);
            var details = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 150, Padding = new Padding(10), FlowDirection = FlowDirection.TopDown, WrapContents = false };
            var configured = "not configured";
            try { configured = new SettingsStore().Load().Model; if (string.IsNullOrWhiteSpace(configured)) configured = "not configured"; } catch { }
            details.Controls.Add(new Label { AutoSize = true, Font = new Font(Font, FontStyle.Bold), Text = "Configured model: " + configured });
            details.Controls.Add(status); details.Controls.Add(progress); details.Controls.Add(currentCase); details.Controls.Add(elapsed);
            var scope = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            scope.Controls.Add(new Label { AutoSize = true, Margin = new Padding(0, 5, 6, 0), Text = "Case ID (blank = all):" });
            scope.Controls.Add(caseFilter); details.Controls.Add(scope);
            Controls.Add(log); Controls.Add(details); Controls.Add(actions);
            finalPdf = ReadLastReport();
            var unfinished = TestLab.ActiveRunId();
            if (unfinished == null) SetIdle("Idle — Start will create a new isolated run. Opening this window made no model request.");
            else
            {
                status.Text = "Prior task still unconfirmed — click Stop to recheck and preserve its report.";
                startButton.Enabled = false; stopButton.Enabled = true; reportButton.Enabled = finalPdf != null;
                currentCase.Text = "Active capture: " + unfinished;
            }
            var timer = new System.Windows.Forms.Timer { Interval = 250 };
            timer.Tick += (s, e) => {
                if (running || finalizing) elapsed.Text = "Elapsed: " + (DateTime.UtcNow - runStarted).ToString(@"hh\:mm\:ss");
                if (activeRunner != null) {
                    var terminal = activeRunner.Results.Count(r => r.status != "not_run" && r.status != "running");
                    progress.Text = "Progress: " + terminal + " / " + activeRunner.Results.Count;
                    currentCase.Text = "Current case/stage: " + (activeRunner.State.caseId ?? "preflight") + " / " + (activeRunner.State.host ?? "preparing");
                }
                var signal = Path.Combine(TestLab.Root, "activate-window-" + nonce);
                if (!File.Exists(signal)) return;
                try { File.Delete(signal); Show(); WindowState = FormWindowState.Normal; Activate(); } catch (IOException) { }
            };
            timer.Start();
            FormClosed += (s, e) => { timer.Dispose(); RemoveWindow(); if (cancellation != null) cancellation.Dispose(); };
            Shown += (s, e) => Activate();
            FormClosing += (s, e) => { if (running || finalizing) { e.Cancel = true; Stop(); Append("The window will remain open until mandatory report finalization finishes."); } };
            RegisterWindow();
        }

        private Button Add(FlowLayoutPanel panel, string text, Func<Task> action)
        {
            var b = new Button { Text = text, AutoSize = true, Height = 34 };
            b.Click += async (s, e) => { try { await action(); } catch (Exception ex) { Append(ex.ToString()); status.Text = "Error — " + ex.Message; } };
            panel.Controls.Add(b);
            return b;
        }

        private Button Add(FlowLayoutPanel panel, string text, Action action)
        { return Add(panel, text, () => { action(); return Task.FromResult(true); }); }

        private void Append(string text)
        {
            if (InvokeRequired) { BeginInvoke((Action)(() => Append(text))); return; }
            log.AppendText(text + Environment.NewLine);
        }

        private void SetIdle(string message)
        {
            running = false; finalizing = false; activeRunner = null;
            Text = "Scribble Test Lab — idle"; status.Text = message;
            progress.Text = "Progress: 0 / 0"; currentCase.Text = "Current case/stage: none"; elapsed.Text = "Elapsed: 00:00:00";
            startButton.Enabled = TestLab.ActiveRunId() == null; stopButton.Enabled = TestLab.ActiveRunId() != null;
            reportButton.Enabled = TestLabPdfWriter.IsValid(finalPdf);
        }

        private static string CreateRunFolder()
        {
            var name = "suite-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N").Substring(0, 8);
            var candidates = new[] {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Scribble Testcases", name),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Scribble", "Testcases", name) };
            Exception failure = null;
            foreach (var candidate in candidates)
            {
                try {
                    Directory.CreateDirectory(candidate); var probe = Path.Combine(candidate, ".write-probe");
                    File.WriteAllText(probe, "ok"); if (File.ReadAllText(probe) != "ok") throw new IOException("Output probe could not be read back.");
                    File.Delete(probe); return candidate;
                } catch (Exception error) { failure = error; }
            }
            throw new IOException("No writable Test Lab output folder is available.", failure);
        }

        private async Task StartRun()
        {
            if (running || finalizing) return;
            if (TestLab.ActiveRunId() != null) { SetIdle("Prior task still unconfirmed — click Stop to recheck it first."); return; }
            cancellation = new CancellationTokenSource(); folder = CreateRunFolder(); finalPdf = null;
            runStarted = DateTime.UtcNow; running = true; Text = "Scribble Test Lab — running";
            status.Text = "Starting — acquiring run ownership and preflight."; startButton.Enabled = false; stopButton.Enabled = true; reportButton.Enabled = false;
            Append("Test Lab runner " + FileVersionInfo.GetVersionInfo(typeof(TestLab).Assembly.Location).FileVersion);
            Append("Results: " + folder);
            caseFilter.Enabled = false;
            var runner = new TestLabSuiteRunner(folder, Append, cancellation.Token, caseFilter.Text); activeRunner = runner;
            try { await RunOnSta(runner.Run); }
            catch (Exception error) { runner.Log("Cannot run suite: " + error); }
            running = false; finalizing = true; status.Text = "Finalizing — creating and validating the current run PDF.";
            try
            {
                await Task.Run(() => TestLabSuiteReport.Create(runner.State, runner.Results.ToArray()));
                finalPdf = Path.Combine(folder, "report.pdf");
                if (!TestLabPdfWriter.IsValid(finalPdf)) throw new InvalidDataException("Reporter did not produce a valid current-run PDF.");
                File.WriteAllText(Path.Combine(TestLab.Root, "last-suite-report.txt"), finalPdf, new UTF8Encoding(false));
                UpdateWindowReport();
                Append("Final PDF: " + finalPdf);
                status.Text = cancellation.IsCancellationRequested ? "Stopped — partial report preserved." : "Finished — final PDF ready for review.";
            }
            catch (Exception error)
            {
                Append("PDF FINALIZATION FAILED: " + error);
                File.AppendAllText(Path.Combine(folder, "summary.txt"), "\nPDF FINALIZATION FAILED: " + error + "\n");
                status.Text = "Reporting failed — diagnostics remain at " + folder;
            }
            finally
            {
                finalizing = false; activeRunner = null; startButton.Enabled = TestLab.ActiveRunId() == null;
                caseFilter.Enabled = true;
                stopButton.Enabled = TestLab.ActiveRunId() != null; reportButton.Enabled = TestLabPdfWriter.IsValid(finalPdf);
                Text = "Scribble Test Lab — " + (reportButton.Enabled ? "finished" : "reporting failed");
            }
        }

        private async void RecoverPrior()
        {
            try
            {
                folder = CreateRunFolder(); finalPdf = await Task.Run(() => TestLabSuite.RecoverIncomplete(folder));
                File.WriteAllText(Path.Combine(TestLab.Root, "last-suite-report.txt"), finalPdf, new UTF8Encoding(false));
                UpdateWindowReport(); status.Text = "Stopped — incomplete capture preserved in the final PDF.";
                Append("Recovered final PDF: " + finalPdf);
            }
            catch (Exception error) { status.Text = "Stop not confirmed — " + error.Message; Append(error.Message); }
            finally { SetIdle(status.Text); }
        }

        private void Stop()
        {
            if (running)
            {
                status.Text = "Stopping — preserving report"; stopButton.Enabled = false;
                if (cancellation != null) cancellation.Cancel(); Append("Stop requested. No next case will be submitted; mandatory evidence finalization continues.");
            }
            else if (finalizing) { status.Text = "Stopping — mandatory report finalization continues"; stopButton.Enabled = false; }
            else if (TestLab.ActiveRunId() != null) { stopButton.Enabled = false; status.Text = "Stopping — checking prior task and preserving report"; RecoverPrior(); }
        }

        private void ViewReport()
        {
            if (!TestLabPdfWriter.IsValid(finalPdf)) throw new InvalidOperationException("The displayed terminal run does not have a validated PDF. Path: " + (finalPdf ?? "not created"));
            Process.Start(new ProcessStartInfo(finalPdf) { UseShellExecute = true });
        }

        private static string ReadLastReport()
        {
            try { var path = File.ReadAllText(Path.Combine(TestLab.Root, "last-suite-report.txt")).Trim(); return TestLabPdfWriter.IsValid(path) ? path : null; }
            catch { return null; }
        }

        private void RegisterWindow()
        {
            using (var process = Process.GetCurrentProcess())
                File.WriteAllText(WindowDescriptor, TestLab.Serialize(new SuiteWindowState { pid = process.Id,
                    processStart = process.StartTime.ToUniversalTime().Ticks, nonce = nonce, ready = true,
                    module = typeof(TestLab).Assembly.ManifestModule.ModuleVersionId.ToString(), report = finalPdf }));
        }

        private void UpdateWindowReport() { RegisterWindow(); }

        private void RemoveWindow()
        {
            try { var state = ReadWindow(); if (state != null && state.nonce == nonce) File.Delete(WindowDescriptor); }
            catch (IOException) { }
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
    }

    public sealed class SuiteWindowState
    {
        public int pid { get; set; }
        public long processStart { get; set; }
        public string nonce { get; set; }
        public string module { get; set; }
        public bool ready { get; set; }
        public string report { get; set; }
    }
    public static class TestLabSuiteReport
    {
        private static string E(string text) => WebUtility.HtmlEncode(text ?? "");
        public static string BuildHtml(SuiteState s, SuiteCaseResult[] results)
        {
            var summary = new StringBuilder("SCRIBBLE TEST LAB\nRunner build: " + FileVersionInfo.GetVersionInfo(typeof(TestLab).Assembly.Location).FileVersion + "\nSuite: " + s.id + "\nMain: " + s.commit + "\nKit SHA256: " + s.kitHash + "\nResults: " + s.folder + "\n");
            summary.AppendLine("Needs review: " + results.Count(r => r.status == "needs_review") + " | Deterministic failures: " + results.Count(r => r.status == "failed") + " | Blocked/stopped: " + results.Count(r => r.status == "blocked" || r.status == "stopped" || r.status == "incomplete") + " | Not run: " + results.Count(r => r.status == "not_run"));
            summary.AppendLine("Completion is not a correctness pass. See native output and visual review in the evidence.");
            foreach (var r in results) summary.AppendLine(r.id + " " + r.host + " — " + r.status +
                (string.IsNullOrEmpty(r.failureKind) ? "" : " [" + r.failureKind + "]") +
                (string.IsNullOrEmpty(r.error) ? "" : " — " + (r.error.Split('\n')[0].Length > 140 ? r.error.Split('\n')[0].Substring(0, 140) + "…" : r.error.Split('\n')[0])));
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
                if (!string.IsNullOrEmpty(r.evaluation) && File.Exists(r.evaluation))
                    html.Append("<h2>Deterministic evaluation</h2><pre>").Append(E(File.ReadAllText(r.evaluation))).Append("</pre>");
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
            File.WriteAllText(html, BuildHtml(s, results), new UTF8Encoding(false));
            return TestLabPdfWriter.CreateSuite(s, results);
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
