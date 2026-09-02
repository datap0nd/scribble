using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.Win32.SafeHandles;
using Scribble.Configuration;
using Scribble.Outlook;
using Scribble.Security;

namespace Scribble.Chat
{
    public sealed class TopicIndexStatus
    {
        public int IndexedFiles { get; set; }

        public int SkippedFiles { get; set; }

        public int FailedFiles { get; set; }

        public DateTime RefreshedUtc { get; set; }

        public bool ReusedFreshIndex { get; set; }

        public bool FolderAvailable { get; set; }

        public bool MatchesConfiguredRoot { get; set; }
    }

    internal sealed class TopicIndexEntry
    {
        public string RelativePath { get; set; }

        public long Length { get; set; }

        public long ModifiedUtcTicks { get; set; }

        public string Content { get; set; }
    }

    internal sealed class TopicIndexManifest
    {
        public int Version { get; set; }

        public string TopicId { get; set; }

        public string RootPath { get; set; }

        public long RefreshedUtcTicks { get; set; }

        public int SkippedFiles { get; set; }

        public int FailedFiles { get; set; }

        public List<TopicIndexEntry> Files { get; set; }
    }

    internal sealed class TopicSearchHit
    {
        public TopicIndexEntry Entry { get; set; }

        public bool PathPhraseMatch { get; set; }

        public bool ContentPhraseMatch { get; set; }

        public int TermCoverage { get; set; }

        public int Frequency { get; set; }

        public string Snippet { get; set; }
    }

    public sealed class TopicIndex
    {
        public const int MaxIndexedFiles = 2000;
        public const int MaxFileBytes = 25 * 1024 * 1024;
        public const int MaxCharactersPerFile = 48000;
        public const int FreshSeconds = 30;
        private const int ManifestVersion = 1;
        private const uint FileFlagBackupSemantics = 0x02000000;

        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer
            {
                MaxJsonLength = int.MaxValue,
                RecursionLimit = 100
            };

        public TopicIndexStatus Refresh(
            TopicConfig topic,
            bool force,
            CancellationToken cancellationToken)
        {
            var config = RequireTopic(topic);
            string root;
            string validationError;
            if (!TopicConfig.TryValidateLocalFolder(
                    config.FolderPath,
                    out root,
                    out validationError))
            {
                throw new InvalidOperationException(validationError);
            }

            root = ResolveFinalPath(root);

            var mutexName = "Scribble.Topic." + config.Id;
            using (var mutex = new Mutex(false, mutexName))
            {
                var entered = false;
                try
                {
                    while (!entered)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        try
                        {
                            entered = mutex.WaitOne(250);
                        }
                        catch (AbandonedMutexException)
                        {
                            entered = true;
                        }
                    }

                    var previous = LoadManifest(config.Id);
                    if (!force && IsFresh(previous, root))
                    {
                        return Status(previous, true);
                    }

                    var updated = BuildManifest(
                        config,
                        root,
                        previous,
                        cancellationToken);
                    SaveManifest(config.Id, updated);
                    return Status(updated, false);
                }
                finally
                {
                    if (entered)
                    {
                        mutex.ReleaseMutex();
                    }
                }
            }
        }

        public TopicIndexStatus GetStatus(TopicConfig topic)
        {
            var config = RequireTopic(topic);
            var manifest = LoadManifest(config.Id);
            if (manifest == null)
            {
                return null;
            }

            var status = Status(manifest, false);
            status.FolderAvailable = false;
            status.MatchesConfiguredRoot = false;
            string root;
            string error;
            if (TopicConfig.TryValidateLocalFolder(
                    config.FolderPath,
                    out root,
                    out error))
            {
                try
                {
                    root = ResolveFinalPath(root);
                    status.FolderAvailable = true;
                    status.MatchesConfiguredRoot = string.Equals(
                        manifest.RootPath,
                        root,
                        StringComparison.OrdinalIgnoreCase);
                }
                catch
                {
                }
            }

            return status;
        }

        internal IReadOnlyList<TopicSearchHit> Search(
            TopicConfig topic,
            string query,
            int maximumResults,
            CancellationToken cancellationToken)
        {
            Refresh(topic, false, cancellationToken);
            var manifest = LoadManifest(RequireTopic(topic).Id);
            if (manifest == null)
            {
                return new TopicSearchHit[0];
            }

            var boundedQuery = TextBoundary.SingleLine(query, 240).Trim();
            var terms = SplitTerms(boundedQuery);
            var hits = new List<TopicSearchHit>();
            foreach (var entry in manifest.Files ??
                new List<TopicIndexEntry>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                var hit = Rank(entry, boundedQuery, terms);
                if (boundedQuery.Length > 0 &&
                    !hit.PathPhraseMatch &&
                    !hit.ContentPhraseMatch &&
                    hit.TermCoverage == 0)
                {
                    continue;
                }

                hits.Add(hit);
            }

            return hits
                .OrderByDescending(hit => hit.PathPhraseMatch)
                .ThenByDescending(hit => hit.ContentPhraseMatch)
                .ThenByDescending(hit => hit.TermCoverage)
                .ThenByDescending(hit => hit.Frequency)
                .ThenByDescending(hit =>
                    hit.Entry.ModifiedUtcTicks)
                .ThenBy(hit => hit.Entry.RelativePath,
                    StringComparer.OrdinalIgnoreCase)
                .Take(Math.Max(1, Math.Min(10, maximumResults)))
                .ToArray();
        }

        internal TopicIndexEntry Find(
            TopicConfig topic,
            string relativePath)
        {
            var config = RequireTopic(topic);
            var manifest = LoadManifest(config.Id);
            return (manifest?.Files ??
                    new List<TopicIndexEntry>())
                .FirstOrDefault(entry => string.Equals(
                    entry.RelativePath,
                    relativePath,
                    StringComparison.OrdinalIgnoreCase));
        }

        internal bool Revalidate(
            TopicConfig topic,
            TopicIndexEntry entry,
            out string error)
        {
            error = string.Empty;
            try
            {
                var config = RequireTopic(topic);
                string root;
                string validationError;
                if (!TopicConfig.TryValidateLocalFolder(
                        config.FolderPath,
                        out root,
                        out validationError))
                {
                    error = validationError;
                    return false;
                }

                root = ResolveFinalPath(root);
                var unresolved = Path.Combine(
                    root,
                    entry.RelativePath);
                RejectReparsePoint(unresolved);
                var fullPath = SafeContainedPath(root, unresolved);
                var info = new FileInfo(fullPath);
                if (!info.Exists ||
                    (info.Attributes & FileAttributes.ReparsePoint) != 0 ||
                    info.Length != entry.Length ||
                    info.LastWriteTimeUtc.Ticks != entry.ModifiedUtcTicks)
                {
                    error = "The indexed file changed. Search the Topic again.";
                    return false;
                }

                return true;
            }
            catch
            {
                error = "The indexed file is no longer accessible.";
                return false;
            }
        }

        public void DeleteCache(string topicId)
        {
            Guid parsed;
            if (!Guid.TryParse(topicId, out parsed))
            {
                return;
            }

            var boundedId = parsed.ToString("N");
            using (var mutex = new Mutex(
                false,
                "Scribble.Topic." + boundedId))
            {
                var entered = false;
                try
                {
                    try
                    {
                        entered = mutex.WaitOne();
                    }
                    catch (AbandonedMutexException)
                    {
                        entered = true;
                    }

                    var directory = CacheDirectory(boundedId);
                    if (Directory.Exists(directory))
                    {
                        Directory.Delete(directory, true);
                    }
                }
                finally
                {
                    if (entered)
                    {
                        mutex.ReleaseMutex();
                    }
                }
            }
        }

        private TopicIndexManifest BuildManifest(
            TopicConfig topic,
            string root,
            TopicIndexManifest previous,
            CancellationToken cancellationToken)
        {
            var oldEntries = new Dictionary<string, TopicIndexEntry>(
                StringComparer.OrdinalIgnoreCase);
            if (previous != null && string.Equals(
                    previous.RootPath,
                    root,
                    StringComparison.OrdinalIgnoreCase))
            {
                foreach (var entry in previous.Files ??
                    new List<TopicIndexEntry>())
                {
                    oldEntries[entry.RelativePath] = entry;
                }
            }

            var result = new TopicIndexManifest
            {
                Version = ManifestVersion,
                TopicId = topic.Id,
                RootPath = root,
                RefreshedUtcTicks = DateTime.UtcNow.Ticks,
                Files = new List<TopicIndexEntry>()
            };
            var directories = new Stack<string>();
            var inspectedFiles = 0;
            directories.Push(root);
            while (directories.Count > 0 &&
                   inspectedFiles < MaxIndexedFiles)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var directory = directories.Pop();
                IEnumerable<string> children;
                try
                {
                    children = Directory.EnumerateDirectories(directory)
                        .OrderBy(path => path,
                            StringComparer.OrdinalIgnoreCase)
                        .ToArray();
                }
                catch
                {
                    if (string.Equals(
                        directory,
                        root,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        throw;
                    }

                    result.FailedFiles++;
                    continue;
                }

                foreach (var child in children.Reverse())
                {
                    try
                    {
                        var attributes = File.GetAttributes(child);
                        if ((attributes & (FileAttributes.Hidden |
                                           FileAttributes.System |
                                           FileAttributes.ReparsePoint)) == 0)
                        {
                            SafeContainedPath(root, child);
                            directories.Push(child);
                        }
                    }
                    catch
                    {
                        result.SkippedFiles++;
                    }
                }

                IEnumerable<string> files;
                try
                {
                    files = Directory.EnumerateFiles(directory)
                        .OrderBy(path => path,
                            StringComparer.OrdinalIgnoreCase)
                        .ToArray();
                }
                catch
                {
                    if (string.Equals(
                        directory,
                        root,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        throw;
                    }

                    result.FailedFiles++;
                    continue;
                }

                foreach (var path in files)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (inspectedFiles == MaxIndexedFiles)
                    {
                        result.SkippedFiles++;
                        break;
                    }

                    inspectedFiles++;

                    try
                    {
                        RejectReparsePoint(path);
                        var info = new FileInfo(
                            SafeContainedPath(root, path));
                        var extension = info.Extension;
                        if (info.Name.StartsWith(
                                "~$",
                                StringComparison.Ordinal) ||
                            (info.Attributes &
                             (FileAttributes.Hidden |
                              FileAttributes.System |
                              FileAttributes.ReparsePoint)) != 0 ||
                            !EmailAttachmentReader.IsSupportedExtension(
                                extension) ||
                            EmailAttachmentReader.IsImageExtension(
                                extension) ||
                            info.Length > MaxFileBytes)
                        {
                            result.SkippedFiles++;
                            continue;
                        }

                        var relative = RelativePath(root, info.FullName);
                        TopicIndexEntry old;
                        if (oldEntries.TryGetValue(relative, out old) &&
                            old.Length == info.Length &&
                            old.ModifiedUtcTicks ==
                                info.LastWriteTimeUtc.Ticks)
                        {
                            result.Files.Add(old);
                            continue;
                        }

                        var extracted =
                            EmailAttachmentReader.LoadLocalFile(
                                info.FullName);
                        RejectReparsePoint(info.FullName);
                        var after = new FileInfo(
                            SafeContainedPath(
                                root,
                                info.FullName));
                        if (!after.Exists ||
                            after.Length != info.Length ||
                            after.LastWriteTimeUtc.Ticks !=
                                info.LastWriteTimeUtc.Ticks)
                        {
                            throw new IOException(
                                "The Topic file changed during extraction.");
                        }

                        if (extracted == null ||
                            extracted.ImageDataUrl.Length > 0 ||
                            extracted.Text.Length == 0 ||
                            string.Equals(
                                extracted.Kind,
                                "unreadable",
                                StringComparison.Ordinal))
                        {
                            result.FailedFiles++;
                            continue;
                        }

                        result.Files.Add(new TopicIndexEntry
                        {
                            RelativePath = relative,
                            Length = info.Length,
                            ModifiedUtcTicks =
                                info.LastWriteTimeUtc.Ticks,
                            Content = TextBoundary.PlainText(
                                extracted.Text,
                                MaxCharactersPerFile)
                        });
                    }
                    catch
                    {
                        result.FailedFiles++;
                    }
                }
            }

            return result;
        }

        private static TopicSearchHit Rank(
            TopicIndexEntry entry,
            string query,
            IReadOnlyList<string> terms)
        {
            var path = entry.RelativePath ?? string.Empty;
            var content = entry.Content ?? string.Empty;
            var coverage = 0;
            var frequency = 0;
            foreach (var term in terms)
            {
                var pathCount = CountOccurrences(path, term);
                var contentCount = CountOccurrences(content, term);
                if (pathCount + contentCount > 0)
                {
                    coverage++;
                }

                frequency += pathCount + contentCount;
            }

            return new TopicSearchHit
            {
                Entry = entry,
                PathPhraseMatch = query.Length > 0 &&
                    path.IndexOf(
                        query,
                        StringComparison.OrdinalIgnoreCase) >= 0,
                ContentPhraseMatch = query.Length > 0 &&
                    content.IndexOf(
                        query,
                        StringComparison.OrdinalIgnoreCase) >= 0,
                TermCoverage = coverage,
                Frequency = frequency,
                Snippet = BuildSnippet(content, query, terms)
            };
        }

        private static int CountOccurrences(string value, string term)
        {
            var count = 0;
            var start = 0;
            while (term.Length > 0 && start < value.Length)
            {
                var found = value.IndexOf(
                    term,
                    start,
                    StringComparison.OrdinalIgnoreCase);
                if (found < 0)
                {
                    break;
                }

                count++;
                start = found + term.Length;
            }

            return count;
        }

        private static string BuildSnippet(
            string content,
            string query,
            IReadOnlyList<string> terms)
        {
            var plain = (content ?? string.Empty)
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Replace('\t', ' ');
            var match = query.Length > 0
                ? plain.IndexOf(query,
                    StringComparison.OrdinalIgnoreCase)
                : 0;
            if (match < 0)
            {
                foreach (var term in terms)
                {
                    match = plain.IndexOf(
                        term,
                        StringComparison.OrdinalIgnoreCase);
                    if (match >= 0)
                    {
                        break;
                    }
                }
            }

            var start = Math.Max(0, match < 0 ? 0 : match - 140);
            var length = Math.Min(500, plain.Length - start);
            var snippet = length > 0
                ? plain.Substring(start, length).Trim()
                : string.Empty;
            return start > 0 ? "..." + snippet : snippet;
        }

        private static IReadOnlyList<string> SplitTerms(string query)
        {
            return (query ?? string.Empty)
                .Split(new[]
                {
                    ' ', '\t', '\r', '\n', '.', ',', ';', ':',
                    '/', '\\', '-', '_', '(', ')', '[', ']'
                }, StringSplitOptions.RemoveEmptyEntries)
                .Where(term => term.Length > 1)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(12)
                .ToArray();
        }

        private static string SafeContainedPath(
            string root,
            string candidate)
        {
            var canonicalRoot = Path.GetFullPath(root)
                .TrimEnd(Path.DirectorySeparatorChar,
                    Path.AltDirectorySeparatorChar);
            var canonical = ResolveFinalPath(candidate);
            var prefix = canonicalRoot + Path.DirectorySeparatorChar;
            if (!canonical.StartsWith(
                    prefix,
                    StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(
                    canonical,
                    canonicalRoot,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "A Topic path escaped its configured root.");
            }

            return canonical;
        }

        private static void RejectReparsePoint(string path)
        {
            var attributes = File.GetAttributes(path);
            if ((attributes & FileAttributes.ReparsePoint) != 0)
            {
                throw new InvalidOperationException(
                    "Topic links and junctions are not allowed.");
            }
        }

        // Path.GetFullPath only resolves lexical '..' segments. Opening the
        // object and asking Windows for its final path also resolves mount
        // points and link targets before the containment check.
        private static string ResolveFinalPath(string path)
        {
            using (var handle = CreateFile(
                path,
                0,
                FileShare.Read | FileShare.Write | FileShare.Delete,
                IntPtr.Zero,
                FileMode.Open,
                FileFlagBackupSemantics,
                IntPtr.Zero))
            {
                if (handle.IsInvalid)
                {
                    throw new IOException(
                        "The Topic path could not be resolved.",
                        Marshal.GetLastWin32Error());
                }

                var buffer = new StringBuilder(1024);
                var length = GetFinalPathNameByHandle(
                    handle,
                    buffer,
                    buffer.Capacity,
                    0);
                if (length == 0)
                {
                    throw new IOException(
                        "The Topic path could not be resolved.",
                        Marshal.GetLastWin32Error());
                }

                if (length >= buffer.Capacity)
                {
                    buffer = new StringBuilder((int)length + 1);
                    length = GetFinalPathNameByHandle(
                        handle,
                        buffer,
                        buffer.Capacity,
                        0);
                    if (length == 0)
                    {
                        throw new IOException(
                            "The Topic path could not be resolved.",
                            Marshal.GetLastWin32Error());
                    }
                }

                var resolved = buffer.ToString();
                if (resolved.StartsWith(
                        "\\\\?\\UNC\\",
                        StringComparison.OrdinalIgnoreCase))
                {
                    resolved = "\\\\" + resolved.Substring(8);
                }
                else if (resolved.StartsWith(
                    "\\\\?\\",
                    StringComparison.Ordinal))
                {
                    resolved = resolved.Substring(4);
                }

                var full = Path.GetFullPath(resolved);
                var root = Path.GetPathRoot(full) ?? string.Empty;
                return full.Length > root.Length
                    ? full.TrimEnd(
                        Path.DirectorySeparatorChar,
                        Path.AltDirectorySeparatorChar)
                    : full;
            }
        }

        [DllImport(
            "kernel32.dll",
            CharSet = CharSet.Unicode,
            SetLastError = true)]
        private static extern SafeFileHandle CreateFile(
            string fileName,
            uint desiredAccess,
            FileShare shareMode,
            IntPtr securityAttributes,
            FileMode creationDisposition,
            uint flagsAndAttributes,
            IntPtr templateFile);

        [DllImport(
            "kernel32.dll",
            CharSet = CharSet.Unicode,
            SetLastError = true)]
        private static extern uint GetFinalPathNameByHandle(
            SafeFileHandle file,
            StringBuilder filePath,
            int filePathLength,
            uint flags);

        private static string RelativePath(string root, string path)
        {
            return path.Substring(
                root.TrimEnd(Path.DirectorySeparatorChar,
                    Path.AltDirectorySeparatorChar).Length)
                .TrimStart(Path.DirectorySeparatorChar,
                    Path.AltDirectorySeparatorChar);
        }

        private bool IsFresh(TopicIndexManifest manifest, string root)
        {
            return manifest != null &&
                   manifest.Version == ManifestVersion &&
                   string.Equals(
                       manifest.RootPath,
                       root,
                       StringComparison.OrdinalIgnoreCase) &&
                   new DateTime(
                       manifest.RefreshedUtcTicks,
                       DateTimeKind.Utc) >=
                       DateTime.UtcNow.AddSeconds(-FreshSeconds);
        }

        private static TopicIndexStatus Status(
            TopicIndexManifest manifest,
            bool reused)
        {
            return new TopicIndexStatus
            {
                IndexedFiles = manifest.Files?.Count ?? 0,
                SkippedFiles = manifest.SkippedFiles,
                FailedFiles = manifest.FailedFiles,
                RefreshedUtc = new DateTime(
                    manifest.RefreshedUtcTicks,
                    DateTimeKind.Utc),
                ReusedFreshIndex = reused,
                FolderAvailable = true,
                MatchesConfiguredRoot = true
            };
        }

        private TopicIndexManifest LoadManifest(string topicId)
        {
            try
            {
                var path = ManifestPath(topicId);
                if (!File.Exists(path))
                {
                    return null;
                }

                var manifest =
                    _serializer.Deserialize<TopicIndexManifest>(
                        File.ReadAllText(path));
                return IsValidManifest(manifest, topicId)
                    ? manifest
                    : null;
            }
            catch
            {
                return null;
            }
        }

        private static bool IsValidManifest(
            TopicIndexManifest manifest,
            string topicId)
        {
            if (manifest == null ||
                manifest.Version != ManifestVersion ||
                !string.Equals(
                    manifest.TopicId,
                    topicId,
                    StringComparison.OrdinalIgnoreCase) ||
                string.IsNullOrWhiteSpace(manifest.RootPath) ||
                !Path.IsPathRooted(manifest.RootPath) ||
                manifest.RootPath.StartsWith(
                    "\\\\",
                    StringComparison.Ordinal) ||
                manifest.RefreshedUtcTicks < DateTime.MinValue.Ticks ||
                manifest.RefreshedUtcTicks > DateTime.MaxValue.Ticks ||
                manifest.Files == null ||
                manifest.Files.Count > MaxIndexedFiles)
            {
                return false;
            }

            var fullRoot = Path.GetFullPath(manifest.RootPath);
            var pathRoot = Path.GetPathRoot(fullRoot) ?? string.Empty;
            var root = fullRoot.Length > pathRoot.Length
                ? fullRoot.TrimEnd(
                    Path.DirectorySeparatorChar,
                    Path.AltDirectorySeparatorChar)
                : fullRoot;
            var prefix = root.EndsWith(
                    Path.DirectorySeparatorChar.ToString(),
                    StringComparison.Ordinal)
                ? root
                : root + Path.DirectorySeparatorChar;
            foreach (var entry in manifest.Files)
            {
                if (entry == null ||
                    string.IsNullOrWhiteSpace(entry.RelativePath) ||
                    Path.IsPathRooted(entry.RelativePath) ||
                    entry.Length < 0 ||
                    entry.Length > MaxFileBytes ||
                    entry.ModifiedUtcTicks < DateTime.MinValue.Ticks ||
                    entry.ModifiedUtcTicks > DateTime.MaxValue.Ticks ||
                    (entry.Content ?? string.Empty).Length >
                        MaxCharactersPerFile)
                {
                    return false;
                }

                var combined = Path.GetFullPath(
                    Path.Combine(root, entry.RelativePath));
                if (!combined.StartsWith(
                    prefix,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            return true;
        }

        private void SaveManifest(
            string topicId,
            TopicIndexManifest manifest)
        {
            var directory = CacheDirectory(topicId);
            Directory.CreateDirectory(directory);
            var path = ManifestPath(topicId);
            var temporary = path + "." +
                Guid.NewGuid().ToString("N") + ".tmp";
            File.WriteAllText(
                temporary,
                _serializer.Serialize(manifest));
            try
            {
                if (File.Exists(path))
                {
                    File.Replace(temporary, path, null);
                }
                else
                {
                    File.Move(temporary, path);
                }
            }
            finally
            {
                if (File.Exists(temporary))
                {
                    File.Delete(temporary);
                }
            }
        }

        private static TopicConfig RequireTopic(TopicConfig topic)
        {
            if (topic == null)
            {
                throw new ArgumentNullException(nameof(topic));
            }

            var normalized = topic.Sanitized();
            if (normalized.Name.Length == 0 ||
                normalized.FolderPath.Length == 0)
            {
                throw new InvalidOperationException(
                    "The active Topic is incomplete.");
            }

            return normalized;
        }

        private static string CacheDirectory(string topicId)
        {
            return Path.Combine(
                Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData),
                "Scribble",
                "topics",
                topicId);
        }

        private static string ManifestPath(string topicId)
        {
            return Path.Combine(CacheDirectory(topicId), "index.json");
        }
    }
}
