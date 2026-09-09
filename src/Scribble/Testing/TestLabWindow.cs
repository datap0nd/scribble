using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Scribble.Testing
{
    public sealed class TestLabWindow : Form
    {
        private readonly ComboBox cases = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly TextBox details = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
        private readonly Label state = new Label { Dock = DockStyle.Top, Height = 45 };
        private readonly string host;
        private string lastRun;
        private string preparationPath;
        private DateTime preparationStarted;
        private string preparationText = "";
        private bool preparing;
        private bool exporting;
        private string reportPdf;
        private string reportSummary;
        private readonly Timer timer = new Timer { Interval = 2000 };
        public static void Open(string host, Action resetConversation = null, Action<LabCase> loadContext = null)
        {
            if (TestLab.Status() == null) return;
            var window = new TestLabWindow(host, resetConversation, loadContext); window.Show();
        }
        public TestLabWindow(string hostName, Action resetConversation = null, Action<LabCase> loadContext = null)
        {
            host = hostName; Text = "Scribble Test Lab"; Size = new Size(750, 560); MinimumSize = new Size(650, 480);
            Font = new Font("Segoe UI", 10); StartPosition = FormStartPosition.CenterScreen;
            var controls = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 140, Padding = new Padding(8), AutoScroll = true };
            Add(controls, "Prepare case", () => {
                if (preparing) throw new InvalidOperationException("Preparation is already running. Check the progress below.");
                var c = cases.SelectedItem as LabCase; if (c == null) return;
                preparationPath = TestLabPreparation.Launch(c.id); preparationStarted = DateTime.UtcNow;
                preparing = true; cases.Enabled = false; preparationText = "Opening apps and fixture files..."; RefreshDetails();
            });
            Add(controls, "Stop preparation", () => { if (preparing) TestLabPreparation.Stop(preparationPath); });
            Add(controls, "Start case", () => {
                var c = cases.SelectedItem as LabCase; if (c == null) return;
                if (preparing) throw new InvalidOperationException("Wait for preparation to finish before starting capture.");
                if (!string.Equals(host, c.host, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Start " + c.id + " from Test Lab in " + c.host + ". You can prepare it from any app.");
                if (MessageBox.Show(this, "Use a new Scribble chat and only this case's synthetic documents, emails and pages. Full prompts, sources and model replies will be recorded locally. Have you prepared that isolated context?", "Start synthetic capture", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
                if (!string.IsNullOrEmpty(TestLab.ActiveRunId())) throw new InvalidOperationException("Finish the current case first.");
                resetConversation?.Invoke();
                var run = TestLab.Start(c.id, host, true); lastRun = run.run_id; Clipboard.SetText(string.IsNullOrEmpty(c.prerequisite_prompt) ? c.prompt : c.prerequisite_prompt); RefreshState();
                loadContext?.Invoke(c);
                MessageBox.Show(this, "Case started. " + (loadContext == null ? "Close this window to return to Chrome; its prompt will be filled. The clipboard is available as a fallback. " : "The prompt is filled and the case context is loading. Check the context tray before submitting. ") + (string.IsNullOrEmpty(c.prerequisite_prompt) ? "" : "Run the prerequisite first, then use Copy prompt for the follow-up. ") + "Start your recording, then submit the prompt yourself. No model request has been sent by Test Lab.");
            });
            Add(controls, "Copy prompt", () => { var c = cases.SelectedItem as LabCase; if (c != null) Clipboard.SetText(c.prompt); });
            Add(controls, "Video marker", () => TestLab.Marker("Operator marker"));
            Add(controls, "Finish clean attempt", () => Finish(false));
            Add(controls, "Finish assisted attempt", () => Finish(true));
            Add(controls, "Collect saved outputs", Collect);
            Add(controls, "Capture new Office drafts", () => { if (lastRun == null) throw new InvalidOperationException("Start a case first."); MessageBox.Show(this, BenchmarkArtifactCollector.Capture(lastRun), "Captured run-owned documents"); });
            Add(controls, "Export PDF + evidence", Export);
            Add(controls, "Copy report summary", () => { if (reportSummary == null || !File.Exists(reportSummary)) throw new InvalidOperationException("Export the report first."); Clipboard.SetText(File.ReadAllText(reportSummary)); });
            Add(controls, "Open report PDF", () => { if (reportPdf == null || !File.Exists(reportPdf)) throw new InvalidOperationException("Export the report first."); System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(reportPdf) { UseShellExecute = true }); });
            Add(controls, "Disable Test Lab", () => { TestLab.Disable(); RefreshState(); });
            Controls.Add(details); Controls.Add(cases); Controls.Add(state); Controls.Add(controls);
            foreach (var c in TestLab.Cases()) cases.Items.Add(c);
            cases.SelectedIndexChanged += (s, e) => { preparationText = ""; RefreshDetails(); };
            if (cases.Items.Count > 0) cases.SelectedIndex = 0;
            lastRun = TestLab.ActiveRunId();
            if (lastRun == null && Directory.Exists(Path.Combine(TestLab.Root, "runs")))
                lastRun = new DirectoryInfo(Path.Combine(TestLab.Root, "runs")).GetDirectories().OrderByDescending(d => d.CreationTimeUtc).FirstOrDefault()?.Name;
            timer.Tick += (s, e) => RefreshState(); timer.Start(); RefreshState();
        }
        private void Add(FlowLayoutPanel panel, string caption, Action action)
        {
            var button = new Button { Text = caption, AutoSize = true, Height = 32 };
            button.Click += (s, e) => { try { action(); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "Test Lab"); } }; panel.Controls.Add(button);
        }
        private void RefreshState()
        {
            if (preparing)
            {
                var report = TestLabPreparation.ReadReport(preparationPath);
                if (report != null)
                {
                    preparationText = string.Join("\r\n", report.completed ?? new string[0]) + "\r\n\r\nRemaining steps:\r\n" + string.Join("\r\n", report.remaining ?? new string[0]);
                    if (report.status == "finished" || report.status == "failed") { preparing = false; cases.Enabled = true; }
                }
                if (preparing && DateTime.UtcNow - preparationStarted > TimeSpan.FromSeconds(90))
                    preparationText += "\r\nOffice startup is taking longer than expected. Check for an Office sign-in, first-run or file dialog. Capture has not started.";
                RefreshDetails();
            }
            var s = TestLab.Status();
            if (!string.IsNullOrEmpty(s?.run_id)) lastRun = s.run_id;
            state.Text = s == null ? "Test Lab disabled or expired. Existing runs can still be exported." :
                "Test Lab enabled until " + s.expires_utc.ToLocalTime().ToString("HH:mm") + " | " + (s.run_id == null ? "No active capture" : TestLab.CaptureState().ToUpperInvariant() + " " + s.run_id);
        }
        private void RefreshDetails()
        {
            var c = cases.SelectedItem as LabCase;
            details.Text = c == null ? "" : preparationText + "\r\n\r\n" + c.setup + "\r\n\r\n" + c.prompt + "\r\n\r\nRequired outputs: " + string.Join(", ", c.artifacts ?? new string[0]);
        }
        private void Finish(bool assisted) { lastRun = TestLab.ActiveRunId(); TestLab.Finish(assisted); RefreshState(); }
        private void Collect()
        {
            if (lastRun == null) throw new InvalidOperationException("Start a case first.");
            using (var picker = new OpenFileDialog { Multiselect = true, Title = "Select the actual saved test outputs", Filter = "Test outputs|*.xlsx;*.pptx;*.docx;*.pdf;*.png;*.msg;*.eml;*.html;*.json" })
                if (picker.ShowDialog(this) == DialogResult.OK) foreach (var p in picker.FileNames) TestLab.Collect(lastRun, p);
        }
        private async void Export()
        {
            if (exporting) return;
            try {
                if (lastRun == null) throw new InvalidOperationException("No run available.");
                using (var picker = new FolderBrowserDialog { Description = "Choose where to save the PDF report, pasteable summary and evidence ZIP" }) {
                    if (picker.ShowDialog(this) != DialogResult.OK) return;
                    exporting = true;
                    var id = lastRun; var destination = picker.SelectedPath;
                    var zip = await Task.Run(() => TestLab.Export(id, destination));
                    reportSummary = Path.Combine(Path.GetDirectoryName(zip), Path.GetFileNameWithoutExtension(zip) + "-summary.txt");
                    reportPdf = await Task.Run(() => TestLabReport.Create(zip));
                    MessageBox.Show(this, "PDF: " + reportPdf + "\r\n\r\nUse Open report PDF for a screenshot, or Copy report summary to paste into chat. The evidence ZIP contains the report and original outputs.", "Report exported");
                }
            } catch (Exception e) { MessageBox.Show(this, e.Message, "Report export"); }
            finally { exporting = false; }
        }
        protected override void Dispose(bool disposing) { if (disposing) timer.Dispose(); base.Dispose(disposing); }
    }
}
