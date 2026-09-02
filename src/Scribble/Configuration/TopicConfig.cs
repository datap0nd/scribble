using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Scribble.Security;

namespace Scribble.Configuration
{
    public sealed class TopicConfig
    {
        public const int MaxTopics = 20;

        public string Id { get; set; } = string.Empty;

        public string Name { get; set; } = string.Empty;

        public string FolderPath { get; set; } = string.Empty;

        public TopicConfig Sanitized()
        {
            Guid parsed;
            var id = Guid.TryParse(Id, out parsed)
                ? parsed.ToString("N")
                : Guid.NewGuid().ToString("N");
            return new TopicConfig
            {
                Id = id,
                Name = TextBoundary.SingleLine(Name, 80).Trim(),
                FolderPath = TextBoundary.SingleLine(
                    FolderPath,
                    1000).Trim()
            };
        }

        public static IReadOnlyList<TopicConfig> Normalize(
            IEnumerable<TopicConfig> topics)
        {
            var result = new List<TopicConfig>();
            var names = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            var ids = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            foreach (var topic in topics ??
                Enumerable.Empty<TopicConfig>())
            {
                if (topic == null)
                {
                    continue;
                }

                var normalized = topic.Sanitized();
                if (normalized.Name.Length == 0 ||
                    normalized.FolderPath.Length == 0 ||
                    !names.Add(normalized.Name))
                {
                    continue;
                }

                while (!ids.Add(normalized.Id))
                {
                    normalized.Id = Guid.NewGuid().ToString("N");
                }

                result.Add(normalized);
                if (result.Count == MaxTopics)
                {
                    break;
                }
            }

            return result;
        }

        public static bool TryValidateLocalFolder(
            string value,
            out string normalized,
            out string error)
        {
            normalized = string.Empty;
            error = string.Empty;
            try
            {
                var candidate = (value ?? string.Empty).Trim();
                var candidateRoot = Path.GetPathRoot(candidate);
                if (candidate.Length == 0 ||
                    !Path.IsPathRooted(candidate) ||
                    string.IsNullOrWhiteSpace(candidateRoot) ||
                    (!candidateRoot.EndsWith(
                         Path.DirectorySeparatorChar.ToString(),
                         StringComparison.Ordinal) &&
                     !candidateRoot.EndsWith(
                         Path.AltDirectorySeparatorChar.ToString(),
                         StringComparison.Ordinal)))
                {
                    error = "Choose an absolute local folder.";
                    return false;
                }

                if (candidate.StartsWith("\\\\", StringComparison.Ordinal))
                {
                    error = "Network folders cannot be used as Topics.";
                    return false;
                }

                normalized = NormalizeFullPath(candidate);
                if (!Directory.Exists(normalized))
                {
                    error = "The Topic folder does not exist.";
                    return false;
                }

                var attributes = File.GetAttributes(normalized);
                if ((attributes & FileAttributes.ReparsePoint) != 0)
                {
                    error = "A Topic root cannot be a link or junction.";
                    return false;
                }

                var root = Path.GetPathRoot(normalized);
                if (!string.IsNullOrWhiteSpace(root))
                {
                    var drive = new DriveInfo(root);
                    if (drive.DriveType == DriveType.Network)
                    {
                        error = "Network drives cannot be used as Topics.";
                        return false;
                    }
                }

                return true;
            }
            catch (Exception exception)
            {
                error = "The Topic folder is not accessible (" +
                    exception.GetType().Name + ").";
                normalized = string.Empty;
                return false;
            }
        }

        private static string NormalizeFullPath(string value)
        {
            var full = Path.GetFullPath(value);
            var root = Path.GetPathRoot(full) ?? string.Empty;
            return full.Length > root.Length
                ? full.TrimEnd(
                    Path.DirectorySeparatorChar,
                    Path.AltDirectorySeparatorChar)
                : full;
        }
    }
}
