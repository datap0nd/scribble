using System;
using System.Diagnostics;
using System.IO;

namespace Scribble.Utilities
{
    // User-initiated only. The independent helper downloads and installs outside
    // Office and the install folder; no CMD/PowerShell association is needed.
    public static class SelfUpdater
    {
        public const string InstallerUrl = "https://github.com/datap0nd/scribble/releases/download/continuous/ScribbleSetup.exe";
        public static void LaunchUpdate(string restartExecutable)
        {
            if (Testing.TestLabSuite.Active() != null || Testing.TestLab.ActiveRunId() != null)
                throw new InvalidOperationException("Stop Test Bench and let its report finish before updating Scribble.");
            var installed = Path.GetDirectoryName(typeof(SelfUpdater).Assembly.Location);
            var helper = Path.Combine(installed, "ScribbleUpdater.exe");
            if (!File.Exists(helper)) throw new FileNotFoundException("The updater helper is missing. Run the current Scribble installer once to repair it.", helper);
            var staging = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Scribble", "Updates", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(staging);
            var staged = Path.Combine(staging, "ScribbleUpdater.exe");
            File.Copy(helper, staged);
            if (Testing.TestLab.FileHash(helper) != Testing.TestLab.FileHash(staged)) throw new IOException("Updater helper copy verification failed.");
            Process.Start(new ProcessStartInfo(staged, "--update " + Quote(installed) + " " + Quote(restartExecutable ?? ""))
            { UseShellExecute = true, WorkingDirectory = staging });
        }
        private static string Quote(string value)
        {
            if (value.Contains("\"") || value.EndsWith("\\", StringComparison.Ordinal)) throw new ArgumentException("Invalid updater argument.");
            return "\"" + value + "\"";
        }
        public static string InstalledVersion()
        {
            try { return FileVersionInfo.GetVersionInfo(typeof(SelfUpdater).Assembly.Location).FileVersion ?? "unknown"; }
            catch { return "unknown"; }
        }
    }
}
