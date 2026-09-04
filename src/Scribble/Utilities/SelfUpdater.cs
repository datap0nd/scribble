using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Scribble.Utilities
{
    // User-initiated self update. This is never exposed to the model:
    // only the Settings window calls it after an explicit confirmation.
    public static class SelfUpdater
    {
        public const string InstallerUrl =
            "https://github.com/datap0nd/scribble/releases/latest/download/ScribbleSetup.exe";

        public const int MaxInstallerBytes = 100 * 1024 * 1024;
        public const int MinInstallerBytes = 200 * 1024;

        public static async Task<string> DownloadInstallerAsync(
            CancellationToken cancellationToken)
        {
            ServicePointManager.SecurityProtocol =
                SecurityProtocolType.Tls12;
            var path = Path.Combine(
                Path.GetTempPath(),
                "Scribble-Update-" +
                Guid.NewGuid().ToString("N") +
                ".exe");
            using (var http = new HttpClient())
            {
                http.Timeout = TimeSpan.FromMinutes(5);
                using (var response = await http.GetAsync(
                    InstallerUrl,
                    HttpCompletionOption.ResponseHeadersRead,
                    cancellationToken).ConfigureAwait(true))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(
                            "The update download failed: " +
                            (int)response.StatusCode +
                            " " +
                            response.ReasonPhrase +
                            ".");
                    }

                    using (var source = await response.Content
                        .ReadAsStreamAsync()
                        .ConfigureAwait(true))
                    using (var target = File.Create(path))
                    {
                        var buffer = new byte[81920];
                        var total = 0;
                        while (true)
                        {
                            var read = await source.ReadAsync(
                                buffer,
                                0,
                                buffer.Length,
                                cancellationToken).ConfigureAwait(true);
                            if (read == 0)
                            {
                                break;
                            }

                            total += read;
                            if (total > MaxInstallerBytes)
                            {
                                throw new InvalidOperationException(
                                    "The downloaded installer is larger than expected.");
                            }

                            target.Write(buffer, 0, read);
                        }

                        if (total < MinInstallerBytes)
                        {
                            throw new InvalidOperationException(
                                "The downloaded installer looks incomplete.");
                        }
                    }
                }
            }

            using (var check = File.OpenRead(path))
            {
                var header = new byte[2];
                if (check.Read(header, 0, 2) != 2 ||
                    header[0] != 'M' ||
                    header[1] != 'Z')
                {
                    throw new InvalidOperationException(
                        "The downloaded file is not a Windows installer.");
                }
            }

            return path;
        }

        // The script receives the installer path as %1 so the file itself
        // stays pure ASCII regardless of the user's profile path. One
        // installer carries all five integrations (Chrome plus the Outlook,
        // Excel, PowerPoint, and Word add-ins); the script
        // waits only for the Office hosts whose Scribble component is
        // actually installed, since only those loaded the DLL.
        public static string BuildUpdateScript()
        {
            return BuildUpdateScript(true, true, true, true);
        }

        public static string BuildUpdateScript(
            bool waitOutlook,
            bool waitExcel,
            bool waitPowerPoint,
            bool waitWord)
        {
            if (!waitOutlook &&
                !waitExcel &&
                !waitPowerPoint &&
                !waitWord)
            {
                // Unknown installation state: waiting for every
                // host is the safe default.
                waitOutlook = true;
                waitExcel = true;
                waitPowerPoint = true;
                waitWord = true;
            }

            var hosts = new List<string>();
            if (waitOutlook)
            {
                hosts.Add("OUTLOOK.EXE");
            }

            if (waitExcel)
            {
                hosts.Add("EXCEL.EXE");
            }

            if (waitPowerPoint)
            {
                hosts.Add("POWERPNT.EXE");
            }

            if (waitWord)
            {
                hosts.Add("WINWORD.EXE");
            }

            var builder = new StringBuilder();
            builder.AppendLine("@echo off");
            builder.AppendLine("set \"installer=%~1\"");
            builder.AppendLine("set \"restart=%~2\"");
            builder.AppendLine("set tries=0");
            // Ask each host to close itself first. The user was
            // warned to save, and a polite close still lets Office
            // run its own shutdown; anything still standing after
            // the grace period below is closed for real.
            foreach (var host in hosts)
            {
                builder.AppendLine(
                    "taskkill /IM " + host + " >nul 2>&1");
            }

            builder.AppendLine(":wait");
            builder.AppendLine("set running=0");
            foreach (var host in hosts)
            {
                AppendProcessWait(builder, host);
            }

            builder.AppendLine("if %running%==0 goto install");
            builder.AppendLine("set /a tries+=1");
            builder.AppendLine("if %tries% GEQ 150 exit /b 1");
            // After about thirty seconds a host is almost certainly
            // sitting on a save prompt, so it gets closed forcibly
            // rather than stalling the update forever - which is
            // exactly how installs silently never happened before.
            builder.AppendLine("if %tries% GEQ 15 goto force");
            builder.AppendLine("timeout /T 2 /NOBREAK >nul");
            builder.AppendLine("goto wait");
            builder.AppendLine(":force");
            foreach (var host in hosts)
            {
                builder.AppendLine(
                    "taskkill /F /IM " + host + " >nul 2>&1");
            }

            builder.AppendLine("timeout /T 2 /NOBREAK >nul");
            builder.AppendLine("goto wait");
            builder.AppendLine(":install");
            builder.AppendLine(
                "\"%installer%\" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART");
            builder.AppendLine(
                "if not \"%restart%\"==\"\" start \"\" \"%restart%\"");
            builder.AppendLine("del \"%installer%\"");
            return builder.ToString();
        }

        private static void AppendProcessWait(
            StringBuilder builder,
            string executable)
        {
            builder.AppendLine(
                "tasklist /FI \"IMAGENAME eq " + executable +
                "\" | " +
                "find /I \"" + executable + "\" >nul");
            builder.AppendLine("if not errorlevel 1 set running=1");
        }

        // True when the given per-user Office add-in registration
        // exists, meaning that host has the Scribble component
        // installed and may hold the assembly loaded.
        private static bool AddinRegistered(string subkeyPath)
        {
            try
            {
                using (var key = Microsoft.Win32.Registry
                    .CurrentUser.OpenSubKey(subkeyPath))
                {
                    return key != null;
                }
            }
            catch
            {
                return false;
            }
        }

        // hostApplication may be null when the update starts from the
        // Excel/PowerPoint/Word pane settings: nothing is quit and
        // the script simply waits for the user to close the Office
        // apps.
        public static void LaunchUpdateAndQuitHost(
            object hostApplication,
            string installerPath,
            string restartExecutable)
        {
            var scriptPath = Path.Combine(
                Path.GetTempPath(),
                "Scribble-Update-" +
                Guid.NewGuid().ToString("N") +
                ".cmd");
            File.WriteAllText(
                scriptPath,
                BuildUpdateScript(
                    AddinRegistered(
                        "Software\\Microsoft\\Office\\Outlook" +
                        "\\Addins\\Scribble.AddIn"),
                    AddinRegistered(
                        "Software\\Microsoft\\Office\\Excel" +
                        "\\Addins\\Scribble.ExcelAddIn"),
                    AddinRegistered(
                        "Software\\Microsoft\\Office\\PowerPoint" +
                        "\\Addins\\Scribble.PowerPointAddIn"),
                    AddinRegistered(
                        "Software\\Microsoft\\Office\\Word" +
                        "\\Addins\\Scribble.WordAddIn")),
                Encoding.ASCII);

            // The script runs through cmd.exe invoked by full path,
            // never through the .cmd shell association: locked-down
            // machines can remap or block script associations, which
            // surfaces as "not a valid application for this OS
            // platform". /d also skips any cmd AutoRun commands.
            var comSpec = Environment.GetEnvironmentVariable(
                "ComSpec");
            if (string.IsNullOrEmpty(comSpec))
            {
                comSpec = Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.System),
                    "cmd.exe");
            }

            var start = new System.Diagnostics.ProcessStartInfo
            {
                FileName = comSpec,
                Arguments = "/d /c \"\"" + scriptPath + "\" \"" +
                    installerPath + "\" \"" +
                    (restartExecutable ?? string.Empty) + "\"\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle =
                    System.Diagnostics.ProcessWindowStyle.Hidden
            };
            try
            {
                System.Diagnostics.Process.Start(start);
            }
            catch (Exception exception)
            {
                Log.Error("UpdateScriptLaunch", exception);
                // Last resort on machines that block running
                // anything from the temp folder: open the
                // downloaded installer itself so the user can click
                // through it after closing the Office apps. Nothing
                // is quit on this path.
                System.Diagnostics.Process.Start(
                    new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = installerPath,
                        UseShellExecute = true
                    });
                return;
            }

            if (hostApplication != null)
            {
                dynamic application = hostApplication;
                application.Quit();
            }
        }

        public static string InstalledVersion()
        {
            try
            {
                return System.Diagnostics.FileVersionInfo
                    .GetVersionInfo(
                        typeof(SelfUpdater).Assembly.Location)
                    .FileVersion ?? "unknown";
            }
            catch
            {
                return "unknown";
            }
        }
    }
}
