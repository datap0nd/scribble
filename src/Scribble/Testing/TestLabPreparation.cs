using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;

namespace Scribble.Testing
{
    // Operator entry point only: never included in a model tool catalog.
    public static class TestLabPreparation
    {
        private static readonly Dictionary<string, Process> Workers = new Dictionary<string, Process>();
        public static string Launch(string caseId, bool suite = false)
        {
            var session = TestLab.Status();
            if (session == null) throw new InvalidOperationException("Enable Test Lab first.");
            if (!string.IsNullOrEmpty(session.run_id)) throw new InvalidOperationException("Finish the current run before preparing another case.");
            if (!TestLab.Cases().Any(c => c.id == caseId)) throw new InvalidOperationException("Unknown case.");
            TestLab.VerifyKit(session.fixture_root);
            var script = TestLab.SafeChild(session.fixture_root, "operator/Prepare-ScribbleTestCase.ps1");
            if (!File.Exists(script)) throw new InvalidOperationException("Download the current test kit to use automatic preparation.");
            var directory = Path.Combine(TestLab.Root, "preparations");
            Directory.CreateDirectory(directory);
            var report = Path.Combine(directory, Guid.NewGuid().ToString("N") + ".json");
            var shell = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"WindowsPowerShell\v1.0\powershell.exe");
            var worker = new Process { StartInfo = new ProcessStartInfo(shell, "-NoProfile -STA -ExecutionPolicy Bypass -File " + Quote(script) +
                " -FixtureRoot " + Quote(session.fixture_root) + " -CaseId " + Quote(caseId) + " -AssemblyPath " + Quote(typeof(TestLab).Assembly.Location) + " -ReportPath " + Quote(report) + (suite ? " -Suite" : ""))
            { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden, RedirectStandardOutput = true, RedirectStandardError = true } };
            var sync = new object();
            DataReceivedEventHandler capture = (sender, e) => { if (e.Data != null) lock (sync) { try { File.AppendAllText(report + ".log", DateTime.UtcNow.ToString("O") + " " + e.Data + Environment.NewLine); } catch (IOException) { } } };
            worker.OutputDataReceived += capture; worker.ErrorDataReceived += capture;
            File.WriteAllText(report + ".log", DateTime.UtcNow.ToString("O") + " Preparing " + caseId + Environment.NewLine);
            try { worker.Start(); Workers[report] = worker; worker.BeginOutputReadLine(); worker.BeginErrorReadLine(); }
            catch (Exception e) { File.AppendAllText(report + ".log", e.ToString()); worker.Dispose(); throw; }
            return report;
        }
        private static string Quote(string value)
        {
            if (value.IndexOf('"') >= 0 || value.EndsWith("\\", StringComparison.Ordinal)) throw new ArgumentException("Invalid launch argument.");
            return "\"" + value + "\"";
        }
        public static PreparationReport ReadReport(string path)
        {
            PreparationReport report = null;
            try { if (File.Exists(path)) report = new JavaScriptSerializer().Deserialize<PreparationReport>(File.ReadAllText(path)); }
            catch (ArgumentException) { } // Writer may still be flushing the progress file.
            catch (IOException) { }
            Process worker;
            if (Workers.TryGetValue(path, out worker) && worker.HasExited)
            {
                worker.WaitForExit(); // Drain asynchronous stdout/stderr before publishing completion.
                if (worker.ExitCode != 0 || report == null || report.status == "preparing") report = new PreparationReport { status = "failed", completed = report?.completed,
                    remaining = new[] { "Preparation stopped before completion (exit " + worker.ExitCode + "). See the captured error below." } };
                worker.Dispose(); Workers.Remove(path);
                File.WriteAllText(path, TestLab.Serialize(report));
            }
            if (Workers.TryGetValue(path, out worker) && !worker.HasExited && report != null) report.status = "preparing";
            if (report != null) report.log = ReadLog(path);
            return report;
        }
        public static string ReadLog(string path) { try { if (!File.Exists(path + ".log")) return ""; using (var stream = new FileStream(path + ".log", FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) using (var reader = new StreamReader(stream)) return reader.ReadToEnd(); } catch (IOException) { return ""; } }
        public static void Stop(string path)
        {
            Process worker;
            if (path != null && Workers.TryGetValue(path, out worker)) { if (!worker.HasExited) worker.Kill(); if (worker.WaitForExit(2000)) ReadReport(path); }
        }
        public static string[] ContextFiles(LabCase c)
        {
            var session = TestLab.Status();
            if (session == null) throw new InvalidOperationException("Test Lab is disabled.");
            return (c.inputs ?? new string[0]).Where(p => !p.EndsWith(".eml", StringComparison.OrdinalIgnoreCase) &&
                !p.StartsWith("inputs/browser/", StringComparison.OrdinalIgnoreCase) &&
                !(c.host == "Excel" && p.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) &&
                !(c.host == "PowerPoint" && p.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase)) &&
                !(c.host == "Word" && p.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)))
                .Select(p => TestLab.SafeChild(session.fixture_root, p)).ToArray();
        }
    }
    public sealed class PreparationReport
    {
        public string log { get; set; }
        public string case_id { get; set; }
        public string status { get; set; }
        public string[] completed { get; set; }
        public string[] remaining { get; set; }
    }
}
