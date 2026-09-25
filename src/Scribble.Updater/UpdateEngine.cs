using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace Scribble.Updater
{
    // Framework-only: this helper never loads Scribble.dll or Office.
    public static class UpdateEngine
    {
        public const string ReleaseRoot = "https://github.com/datap0nd/scribble/releases/download/continuous/";
        public const int MaxInstallerBytes = 100 * 1024 * 1024;
        public static string Hash(string path)
        { using (var file = File.OpenRead(path)) using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "").ToLowerInvariant(); }
        public static string VersionOf(string path)
        {
            var v = FileVersionInfo.GetVersionInfo(path);
            return v.FileMajorPart + "." + v.FileMinorPart + "." + v.FileBuildPart + "." + v.FilePrivatePart;
        }
        public static string NoUpdateMessage(string installedVersion, string publicVersion)
        {
            Version installed, available;
            if (!Version.TryParse(installedVersion, out installed) || !Version.TryParse(publicVersion, out available))
                throw new InvalidDataException("The installed or public release version is invalid.");
            if (available > installed) return null;
            var versions = " Installed: " + installedVersion + ". Public stable: " + publicVersion + ".";
            if (available == installed)
                return "Scribble is already up to date with the public stable release." + versions;
            return "No newer public update is available." + versions +
                " Your installed build was kept. Development test builds are distributed separately through GitHub Actions.";
        }
        public static Candidate ParseCandidate(string json)
        {
            var candidate = new JavaScriptSerializer().Deserialize<Candidate>(json.TrimStart('\uFEFF'));
            Version version;
            if (candidate == null || !Version.TryParse(candidate.version, out version) ||
                !Regex.IsMatch(candidate.installer_sha256 ?? "", "^[a-f0-9]{64}$") ||
                !Regex.IsMatch(candidate.commit ?? "", "^[a-f0-9]{40}$"))
                throw new InvalidDataException("The release manifest is incomplete. Try Update again after release publication finishes.");
            return candidate;
        }
        public static void ValidateInstaller(string path, Candidate candidate)
        {
            var size = new FileInfo(path).Length;
            if (size < 200 * 1024 || size > MaxInstallerBytes) throw new InvalidDataException("The installer download is incomplete or exceeds the size limit.");
            using (var file = File.OpenRead(path))
                if (file.ReadByte() != 'M' || file.ReadByte() != 'Z') throw new InvalidDataException("The downloaded file is not a Windows installer.");
            if (Hash(path) != candidate.installer_sha256) throw new InvalidDataException("Installer checksum mismatch. The release may be changing; no installer was started.");
            if (VersionOf(path) != candidate.version) throw new InvalidDataException("Installer version does not match the release manifest.");
        }
        private static async Task Download(HttpClient http, string name, string path, int limit, CancellationToken cancel)
        {
            using (var response = await http.GetAsync(ReleaseRoot + name, HttpCompletionOption.ResponseHeadersRead, cancel))
            {
                response.EnsureSuccessStatusCode();
                using (var source = await response.Content.ReadAsStreamAsync())
                using (var target = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true))
                {
                    var buffer = new byte[81920]; int count; long total = 0;
                    while ((count = await source.ReadAsync(buffer, 0, buffer.Length, cancel)) > 0)
                    {
                        total += count;
                        if (total > limit) throw new InvalidDataException("Release download exceeded the size limit.");
                        await target.WriteAsync(buffer, 0, count, cancel);
                    }
                }
            }
        }
        public static string[] HostNames()
        {
            var names = new List<string> { "ScribbleBrowserHost" };
            foreach (var pair in new[] { new[] { "Outlook", "Scribble.AddIn", "OUTLOOK" }, new[] { "Excel", "Scribble.ExcelAddIn", "EXCEL" },
                new[] { "PowerPoint", "Scribble.PowerPointAddIn", "POWERPNT" }, new[] { "Word", "Scribble.WordAddIn", "WINWORD" } })
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Office\\" + pair[0] + "\\Addins\\" + pair[1]))
                    if (key != null) names.Add(pair[2]);
            if (names.Count == 1) names.AddRange(new[] { "OUTLOOK", "EXCEL", "POWERPNT", "WINWORD" });
            return names.ToArray();
        }
        private static Process[] Hosts()
        {
            var session = Process.GetCurrentProcess().SessionId;
            var result = new List<Process>();
            foreach (var name in HostNames()) foreach (var process in Process.GetProcessesByName(name))
            {
                try { if (!process.HasExited && process.SessionId == session) { result.Add(process); continue; } }
                catch (InvalidOperationException) { }
                process.Dispose();
            }
            return result.ToArray();
        }
        public static string InstallerArguments(string installed, string log)
        {
            if (installed.Contains("\"") || log.Contains("\"")) throw new ArgumentException("Invalid installer path.");
            return "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /NOCLOSEAPPLICATIONS /NOFORCECLOSEAPPLICATIONS /RESTARTEXITCODE=3010 /DIR=\"" + installed + "\" /LOG=\"" + log + "\"";
        }
        public static async Task<string> Run(string installed, string staging, string restart, Action<string> report, Action installing, CancellationToken cancel)
        {
            if (!Path.IsPathRooted(installed) || !File.Exists(Path.Combine(installed, "Scribble.dll"))) throw new InvalidDataException("The Scribble install directory is unavailable.");
            if (!new[] { "", "outlook.exe", "excel.exe", "powerpnt.exe", "winword.exe" }.Contains(restart ?? "")) throw new InvalidDataException("Unsupported restart application.");
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            var installedVersion = VersionOf(Path.Combine(installed, "Scribble.dll"));
            Candidate candidate = null;
            var installer = Path.Combine(staging, "ScribbleSetup.exe");
            using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancel))
            using (var http = new HttpClient())
            {
                timeout.CancelAfter(TimeSpan.FromMinutes(5));
                http.DefaultRequestHeaders.UserAgent.ParseAdd("Scribble-Updater/1.0");
                // Publication replaces two assets. Retry mismatched pairs,
                // never execute bytes that failed checksum/version validation.
                for (int attempt = 0; attempt < 3; attempt++)
                {
                    report("Checking the public stable release. Installed version: " + installedVersion + ".");
                    var manifest = Path.Combine(staging, "candidate.json");
                    await Download(http, "candidate.json", manifest, 65536, timeout.Token);
                    candidate = ParseCandidate(File.ReadAllText(manifest));
                    var noUpdate = NoUpdateMessage(installedVersion, candidate.version);
                    if (noUpdate != null) return noUpdate;
                    report("Installed version: " + installedVersion + ". Public stable version: " + candidate.version + ". Downloading and verifying the update...");
                    await Download(http, "ScribbleSetup.exe", installer, MaxInstallerBytes, timeout.Token);
                    try { ValidateInstaller(installer, candidate); break; }
                    catch (InvalidDataException) when (attempt < 2) { await Task.Delay(2000, timeout.Token); }
                }
            }
            report("Verified version " + candidate.version + ". Closing Scribble's apps; respond to any Office save prompts.");
            var requested = new HashSet<string>();
            var deadline = DateTime.UtcNow.AddMinutes(10);
            while (true)
            {
                cancel.ThrowIfCancellationRequested();
                var hosts = Hosts();
                if (hosts.Length == 0) break;
                var waiting = new List<Process>();
                try
                {
                    foreach (var host in hosts)
                    {
                        try
                        {
                            // Empty Office processes can retain mapped DLLs
                            // after their last window has closed. Setup keeps
                            // those old images while replacing the payload.
                            // A headless Test Bench can still own an active run.
                            if (host.MainWindowHandle == IntPtr.Zero && host.ProcessName != "ScribbleBrowserHost") continue;
                            waiting.Add(host);
                            var identity = host.Id + ":" + host.StartTime.ToUniversalTime().Ticks;
                            if (!requested.Contains(identity) && host.CloseMainWindow()) requested.Add(identity);
                        }
                        catch (InvalidOperationException) { }
                    }
                    if (waiting.Count == 0) break;
                    report("Waiting for: " + string.Join(", ", waiting.Select(p => p.ProcessName).Distinct()) + ". Save your work and close any remaining windows. Update continues automatically.");
                }
                finally { foreach (var host in hosts) host.Dispose(); }
                if (DateTime.UtcNow >= deadline) throw new TimeoutException("Office or Test Bench is still open. Close the listed app, then retry Update. The installer and logs are preserved.");
                await Task.Delay(1000, cancel);
            }
            cancel.ThrowIfCancellationRequested();
            installing();
            report("Installing version " + candidate.version + "...");
            using (var setup = Process.Start(new ProcessStartInfo(installer, InstallerArguments(installed, Path.Combine(staging, "installer.log"))) { UseShellExecute = true }))
            {
                var minutes = 0;
                while (!await Task.Run(() => setup.WaitForExit(60 * 1000)))
                {
                    minutes++;
                    report("Setup is still running (" + minutes + " minutes). Check any visible installer prompt. This window keeps the update locked until setup exits; installer.log records its progress.");
                }
                if (setup.ExitCode == 3010) return "Restart Windows to finish installing version " + candidate.version + ".";
                if (setup.ExitCode != 0) throw new InvalidOperationException("Installer exited with code " + setup.ExitCode + ". See installer.log in this update folder.");
            }
            if (VersionOf(Path.Combine(installed, "Scribble.dll")) != candidate.version) throw new InvalidDataException("Setup exited, but the installed version did not change to " + candidate.version + ". The update is not confirmed.");
            if (!string.IsNullOrEmpty(restart)) Process.Start(new ProcessStartInfo(restart) { UseShellExecute = true });
            return "Scribble " + candidate.version + " is installed and verified. Open Test Bench and click Start.";
        }
    }
    public sealed class Candidate { public string version { get; set; } public string commit { get; set; } public string installer_sha256 { get; set; } }
}
