using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace Scribble.BrowserHost
{
    internal static class BrowserSetup
    {
        internal static int Run(string requestedBrowser)
        {
            // Chrome is the only supported browser; the historical
            // "edge"/"auto" argument values all resolve to Chrome.
            _ = requestedBrowser;

            var extensionDirectory = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "BrowserExtension");
            if (!File.Exists(
                Path.Combine(extensionDirectory, "manifest.json")))
            {
                MessageBox.Show(
                    "The Scribble browser extension files are missing. " +
                    "Run Scribble Setup again and select Browser extension.",
                    "Scribble browser setup",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return 1;
            }

            var browser = ResolveBrowser();
            if (browser == null)
            {
                MessageBox.Show(
                    "Google Chrome was not found on this computer.",
                    "Scribble browser setup",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return 1;
            }

            var copied = CopyWithRetry(extensionDirectory);
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = browser.Executable,
                    Arguments = browser.ExtensionsPage,
                    UseShellExecute = true
                });
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = "\"" + extensionDirectory + "\"",
                    UseShellExecute = true
                });
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                    "Scribble could not open " + browser.DisplayName +
                    " setup (" + exception.GetType().Name + "). " +
                    "Open its Extensions page manually and choose Load unpacked.",
                    "Scribble browser setup",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return 1;
            }

            MessageBox.Show(
                "Two windows are open. In " + browser.DisplayName + ":\n\n" +
                "1. Turn on Developer mode.\n" +
                "2. Select Load unpacked.\n" +
                "3. Choose the BrowserExtension folder shown in File Explorer." +
                (copied
                    ? " Its path is already copied, so you can paste it into the folder box."
                    : string.Empty) +
                "\n\nThis one-time confirmation is required because Scribble is not distributed through a browser store.",
                "Finish Scribble setup in " + browser.DisplayName,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return 0;
        }

        // Scribble supports Google Chrome only; every requested
        // value resolves to the Chrome installation.
        private static BrowserChoice ResolveBrowser()
        {
            return Find(
                "Google Chrome",
                "chrome.exe",
                "chrome://extensions",
                new[]
                {
                    Combine(
                        Environment.SpecialFolder.ProgramFiles,
                        "Google", "Chrome", "Application",
                        "chrome.exe"),
                    Combine(
                        Environment.SpecialFolder.ProgramFilesX86,
                        "Google", "Chrome", "Application",
                        "chrome.exe"),
                    Combine(
                        Environment.SpecialFolder.LocalApplicationData,
                        "Google", "Chrome", "Application",
                        "chrome.exe")
                });
        }

        private static BrowserChoice Find(
            string displayName,
            string executableName,
            string extensionsPage,
            IEnumerable<string> candidates)
        {
            foreach (var candidate in candidates)
            {
                if (!string.IsNullOrWhiteSpace(candidate) &&
                    File.Exists(candidate))
                {
                    return new BrowserChoice(
                        displayName,
                        candidate,
                        extensionsPage);
                }
            }

            var registered = FindRegisteredExecutable(
                executableName);
            return registered.Length > 0
                ? new BrowserChoice(
                    displayName,
                    registered,
                    extensionsPage)
                : null;
        }

        private static string FindRegisteredExecutable(
            string executableName)
        {
            foreach (var hive in new[]
            {
                RegistryHive.CurrentUser,
                RegistryHive.LocalMachine
            })
            {
                foreach (var view in new[]
                {
                    RegistryView.Registry64,
                    RegistryView.Registry32
                })
                {
                    try
                    {
                        using (var root = RegistryKey.OpenBaseKey(
                            hive,
                            view))
                        using (var key = root.OpenSubKey(
                            "Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\" +
                            executableName))
                        {
                            var path = key?.GetValue(null) as string;
                            if (!string.IsNullOrWhiteSpace(path) &&
                                File.Exists(path))
                            {
                                return path;
                            }
                        }
                    }
                    catch
                    {
                    }
                }
            }

            return string.Empty;
        }

        private static string Combine(
            Environment.SpecialFolder root,
            params string[] parts)
        {
            var value = Environment.GetFolderPath(root);
            foreach (var part in parts)
            {
                value = Path.Combine(value, part);
            }

            return value;
        }

        private static bool CopyWithRetry(string value)
        {
            for (var attempt = 0; attempt < 5; attempt++)
            {
                try
                {
                    Clipboard.SetText(value);
                    return true;
                }
                catch
                {
                    Thread.Sleep(80);
                }
            }

            return false;
        }

        private sealed class BrowserChoice
        {
            public BrowserChoice(
                string displayName,
                string executable,
                string extensionsPage)
            {
                DisplayName = displayName;
                Executable = executable;
                ExtensionsPage = extensionsPage;
            }

            public string DisplayName { get; }

            public string Executable { get; }

            public string ExtensionsPage { get; }
        }
    }
}
