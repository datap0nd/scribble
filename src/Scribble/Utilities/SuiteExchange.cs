using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;
using Scribble.Security;

namespace Scribble.Utilities
{
    // Bounded local hand-off between the Scribble panes in Outlook,
    // Excel, and PowerPoint. One user-initiated snippet at a time is
    // stored beside the settings file; another pane can pull it in as
    // ordinary external context. Nothing leaves the machine and the
    // snippet is always added deliberately by the user, never
    // automatically.
    public static class SuiteExchange
    {
        public const int MaxContentCharacters = 48000;

        public sealed class Entry
        {
            public string Source { get; set; }

            public string Title { get; set; }

            public string Content { get; set; }

            public string SavedAt { get; set; }
        }

        private static string FilePath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder
                            .LocalApplicationData),
                    "Scribble",
                    "suite-exchange.json");
            }
        }

        public static void Save(
            string source,
            string title,
            string content)
        {
            try
            {
                var entry = new Entry
                {
                    Source = TextBoundary.SingleLine(source, 40),
                    Title = TextBoundary.SingleLine(title, 180),
                    Content = TextBoundary.PlainText(
                        content,
                        MaxContentCharacters),
                    SavedAt = DateTime.Now.ToString("O")
                };
                if (entry.Content.Length == 0)
                {
                    return;
                }

                var directory = Path.GetDirectoryName(FilePath);
                Directory.CreateDirectory(directory);
                File.WriteAllText(
                    FilePath,
                    new JavaScriptSerializer().Serialize(entry),
                    new UTF8Encoding(false));
            }
            catch (Exception exception)
            {
                Log.Error("SuiteExchangeSave", exception);
            }
        }

        public static Entry TryLoad()
        {
            try
            {
                if (!File.Exists(FilePath))
                {
                    return null;
                }

                var entry = new JavaScriptSerializer()
                    .Deserialize<Entry>(
                        File.ReadAllText(
                            FilePath,
                            Encoding.UTF8));
                if (entry == null)
                {
                    return null;
                }

                entry.Source = TextBoundary.SingleLine(
                    entry.Source,
                    40);
                entry.Title = TextBoundary.SingleLine(
                    entry.Title,
                    180);
                entry.Content = TextBoundary.PlainText(
                    entry.Content,
                    MaxContentCharacters);
                return entry.Content.Length == 0 ? null : entry;
            }
            catch (Exception exception)
            {
                Log.Error("SuiteExchangeLoad", exception);
                return null;
            }
        }
    }
}
