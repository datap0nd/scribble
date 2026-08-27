using System;
using System.IO;

namespace Scribble.Utilities
{
    internal static class Log
    {
        public static void Error(string operation, Exception exception)
        {
            try
            {
                var directory = Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.LocalApplicationData),
                    "Scribble",
                    "logs");
                Directory.CreateDirectory(directory);
                var path = Path.Combine(directory, "addin.log");
                var line =
                    DateTime.UtcNow.ToString("O") + " " +
                    operation + " " +
                    exception.GetType().Name + " " +
                    DiagnosticDetails.CodeForLog(exception) +
                    Environment.NewLine;
                File.AppendAllText(path, line);
            }
            catch
            {
            }
        }

        // Last lines of the diagnostic log for user-reviewed problem
        // reports. Entries carry timestamps, operation names, and
        // error codes only - never email content.
        public static string Tail(int maxLines)
        {
            try
            {
                var path = Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder
                            .LocalApplicationData),
                    "Scribble",
                    "logs",
                    "addin.log");
                if (!File.Exists(path))
                {
                    return "(no log entries)";
                }

                var lines = File.ReadAllLines(path);
                var start = Math.Max(
                    0,
                    lines.Length - Math.Max(1, maxLines));
                var builder = new System.Text.StringBuilder();
                for (var index = start;
                     index < lines.Length;
                     index++)
                {
                    builder.AppendLine(lines[index]);
                }

                return builder.ToString();
            }
            catch
            {
                return "(log unavailable)";
            }
        }
    }
}
