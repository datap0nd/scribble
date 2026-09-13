using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace Scribble.Updater
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            if (args.Length != 3 || args[0] != "--update") return 2;
            using (var mutex = new Mutex(false, @"Local\ScribbleUpdater"))
            {
                bool owned;
                try { owned = mutex.WaitOne(0); } catch (AbandonedMutexException) { owned = true; }
                if (!owned) { MessageBox.Show("A Scribble update is already open. Use that update window.", "Scribble update"); return 1; }
                try { using (var window = new UpdateWindow(args[1], args[2])) Application.Run(window); }
                finally { mutex.ReleaseMutex(); }
            }
            return 0;
        }
    }
    internal sealed class UpdateWindow : Form
    {
        private readonly CancellationTokenSource cancellation = new CancellationTokenSource();
        private readonly Label status = new Label { Dock = DockStyle.Top, Height = 110, Padding = new Padding(16) };
        private readonly TextBox details = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };
        private readonly Button close = new Button { Text = "Cancel", AutoSize = true };
        private bool installing, finished;
        internal UpdateWindow(string installed, string restart)
        {
            Text = "Scribble update"; Size = new Size(700, 430); StartPosition = FormStartPosition.CenterScreen; Font = new Font("Segoe UI", 10);
            var actions = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 50, Padding = new Padding(8) };
            actions.Controls.Add(close); close.Click += (s, e) => { if (finished) Close(); else cancellation.Cancel(); };
            Controls.Add(details); Controls.Add(status); Controls.Add(actions);
            var staging = AppDomain.CurrentDomain.BaseDirectory;
            Action<string> log = text => {
                status.Text = text;
                var line = DateTime.UtcNow.ToString("O") + " " + text + Environment.NewLine;
                details.AppendText(line);
                try { File.AppendAllText(Path.Combine(staging, "update.log"), line); }
                catch (IOException error) { details.AppendText("Could not write update.log: " + error.Message + Environment.NewLine); }
                catch (UnauthorizedAccessException error) { details.AppendText("Could not write update.log: " + error.Message + Environment.NewLine); }
            };
            Shown += async (s, e) =>
            {
                try { log(await UpdateEngine.Run(installed, staging, restart, log, () => { installing = true; close.Enabled = false; }, cancellation.Token)); }
                catch (OperationCanceledException) { log("Update cancelled or download timed out. No new installer will be started."); }
                catch (Exception error) { log("Update failed: " + error); }
                finally { finished = true; installing = false; close.Enabled = true; close.Text = "Close"; details.AppendText("\r\nLogs and installer: " + staging); }
            };
            FormClosing += (s, e) => { if (!finished) { e.Cancel = true; if (!installing) cancellation.Cancel(); } };
            FormClosed += (s, e) => cancellation.Dispose();
        }
    }
}
