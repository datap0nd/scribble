using System;
using System.Drawing;
using System.IO;
using System.Linq;
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
        private readonly Timer timer = new Timer { Interval = 2000 };
        public static void Open(string host, Action resetConversation = null)
        {
            if (TestLab.Status() == null) return;
            var window = new TestLabWindow(host, resetConversation); window.Show();
        }
        public TestLabWindow(string hostName, Action resetConversation = null)
        {
            host = hostName; Text = "Scribble Test Lab"; Size = new Size(750, 560); MinimumSize = new Size(650, 480);
            Font = new Font("Segoe UI", 10); StartPosition = FormStartPosition.CenterScreen;
            var controls = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 140, Padding = new Padding(8), AutoScroll = true };
            Add(controls, "Start case", () => {
                var c = cases.SelectedItem as LabCase; if (c == null) return;
                if (MessageBox.Show(this, "Use a new Scribble chat and only this case's synthetic documents, emails and pages. Full prompts, sources and model replies will be recorded locally. Have you prepared that isolated context?", "Start synthetic capture", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
                if (!string.IsNullOrEmpty(TestLab.ActiveRunId())) throw new InvalidOperationException("Finish the current case first.");
                resetConversation?.Invoke();
                var run = TestLab.Start(c.id, host, true); lastRun = run.run_id; Clipboard.SetText(string.IsNullOrEmpty(c.prerequisite_prompt) ? c.prompt : c.prerequisite_prompt); RefreshState();
                MessageBox.Show(this, "Case started. Add the listed test emails and attachments to the clean chat now. " + (string.IsNullOrEmpty(c.prerequisite_prompt) ? "The exact prompt is on the clipboard." : "The prerequisite prompt is on the clipboard. Run it first, then use Copy prompt for the follow-up.") + " Paste it into Scribble and run it normally. Keep this window open to mark and finish the recording.");
            });
            Add(controls, "Copy prompt", () => { var c = cases.SelectedItem as LabCase; if (c != null) Clipboard.SetText(c.prompt); });
            Add(controls, "Video marker", () => TestLab.Marker("Operator marker"));
            Add(controls, "Finish clean attempt", () => Finish(false));
            Add(controls, "Finish assisted attempt", () => Finish(true));
            Add(controls, "Collect saved outputs", Collect);
            Add(controls, "Capture new Office drafts", () => { if (lastRun == null) throw new InvalidOperationException("Start a case first."); MessageBox.Show(this, BenchmarkArtifactCollector.Capture(lastRun), "Captured run-owned documents"); });
            Add(controls, "Export run ZIP", Export);
            Add(controls, "Disable Test Lab", () => { TestLab.Disable(); RefreshState(); });
            Controls.Add(details); Controls.Add(cases); Controls.Add(state); Controls.Add(controls);
            foreach (var c in TestLab.Cases()) cases.Items.Add(c);
            cases.SelectedIndexChanged += (s, e) => { var c = cases.SelectedItem as LabCase; details.Text = c == null ? "" : c.setup + "\r\n\r\n" + c.prompt + "\r\n\r\nRequired outputs: " + string.Join(", ", c.artifacts ?? new string[0]); };
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
            var s = TestLab.Status();
            if (!string.IsNullOrEmpty(s?.run_id)) lastRun = s.run_id;
            state.Text = s == null ? "Test Lab disabled or expired. Existing runs can still be exported." :
                "Test Lab enabled until " + s.expires_utc.ToLocalTime().ToString("HH:mm") + " | " + (s.run_id == null ? "No active capture" : TestLab.CaptureState().ToUpperInvariant() + " " + s.run_id);
        }
        private void Finish(bool assisted) { lastRun = TestLab.ActiveRunId(); TestLab.Finish(assisted); RefreshState(); }
        private void Collect()
        {
            if (lastRun == null) throw new InvalidOperationException("Start a case first.");
            using (var picker = new OpenFileDialog { Multiselect = true, Title = "Select the actual saved test outputs", Filter = "Test outputs|*.xlsx;*.pptx;*.docx;*.pdf;*.png;*.msg;*.eml;*.html;*.json" })
                if (picker.ShowDialog(this) == DialogResult.OK) foreach (var p in picker.FileNames) TestLab.Collect(lastRun, p);
        }
        private void Export()
        {
            if (lastRun == null) throw new InvalidOperationException("No run available.");
            using (var picker = new FolderBrowserDialog { Description = "Choose where to write the evidence ZIP" })
                if (picker.ShowDialog(this) == DialogResult.OK) MessageBox.Show(this, TestLab.Export(lastRun, picker.SelectedPath), "Exported evidence");
        }
        protected override void Dispose(bool disposing) { if (disposing) timer.Dispose(); base.Dispose(disposing); }
    }
}
