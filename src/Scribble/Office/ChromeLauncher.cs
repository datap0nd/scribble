using System;
using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace Scribble.Office
{
    public static class ChromeLauncher
    {
        public static void ValidateUrl(string url)
        {
            Uri parsed;
            if (string.IsNullOrWhiteSpace(url) || url.Length > 2048 || url.IndexOfAny(new[] { '"', '\r', '\n', '\0' }) >= 0 ||
                !Uri.TryCreate(url, UriKind.Absolute, out parsed) ||
                (parsed.Scheme != "http" && parsed.Scheme != "https") || parsed.UserInfo.Length != 0)
                throw new InvalidOperationException("Chrome requires an HTTP or HTTPS webpage URL without embedded credentials.");
        }
        public static void Open(string url)
        {
            ValidateUrl(url);
            string executable = null;
            foreach (var hive in new[] { Registry.CurrentUser, Registry.LocalMachine })
                using (var key = hive.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"))
                {
                    var candidate = Convert.ToString(key?.GetValue(null)).Trim('"');
                    if (File.Exists(candidate)) { executable = candidate; break; }
                }
            if (executable == null)
                foreach (var folder in new[] { Environment.SpecialFolder.LocalApplicationData, Environment.SpecialFolder.ProgramFiles, Environment.SpecialFolder.ProgramFilesX86 })
                {
                    var candidate = Path.Combine(Environment.GetFolderPath(folder), "Google", "Chrome", "Application", "chrome.exe");
                    if (File.Exists(candidate)) { executable = candidate; break; }
                }
            if (executable == null) throw new InvalidOperationException("Google Chrome is not installed on this computer.");
            Process.Start(new ProcessStartInfo { FileName = executable, Arguments = "--new-window \"" + url + "\"", UseShellExecute = false });
        }
    }
}
