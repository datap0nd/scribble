using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using Scribble.Security;

namespace Scribble.Configuration
{
    public sealed class SkillStore
    {
        public const int SchemaVersion = 1;
        public const int MaxLocalSkillsPerHost = 20;

        private const string PublicManifestResource =
            "Scribble.Skills.PublicSkills.json";
        private const string YesterdayFiveToken =
            "{{yesterday_5pm_local_iso}}";
        private const string NowToken = "{{now_local_iso}}";

        private readonly string _localPath;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();

        public SkillStore()
            : this(DefaultLocalPath())
        {
        }

        public SkillStore(string localPath)
        {
            _localPath = localPath ?? string.Empty;
        }

        public string LocalPath
        {
            get { return _localPath; }
        }

        public IReadOnlyList<SkillDefinition> LoadPublic()
        {
            using (var stream = typeof(SkillStore).Assembly
                .GetManifestResourceStream(PublicManifestResource))
            {
                if (stream == null)
                {
                    return new SkillDefinition[0];
                }

                using (var reader = new StreamReader(
                    stream,
                    Encoding.UTF8,
                    true))
                {
                    return ParseDocument(
                        reader.ReadToEnd(),
                        "public");
                }
            }
        }

        public IReadOnlyList<SkillDefinition> LoadLocal()
        {
            try
            {
                if (_localPath.Length == 0 ||
                    !File.Exists(_localPath))
                {
                    return new SkillDefinition[0];
                }

                return ParseDocument(
                    File.ReadAllText(_localPath, Encoding.UTF8),
                    "local");
            }
            catch
            {
                return new SkillDefinition[0];
            }
        }

        public void SaveLocal(IEnumerable<SkillDefinition> skills)
        {
            if (_localPath.Length == 0)
            {
                throw new InvalidOperationException(
                    "The local Skills path is unavailable.");
            }

            var normalized = ValidateLocal(skills).ToList();
            var document = new StoredSkillDocument
            {
                SchemaVersion = SchemaVersion,
                Skills = normalized.Select(skill =>
                    new StoredSkill
                    {
                        Id = skill.Id,
                        Name = skill.Name,
                        Description = skill.Description,
                        Prompt = skill.Prompt,
                        Host = skill.Host,
                        DisplayOrder = skill.DisplayOrder,
                        StartFresh = skill.StartFresh
                    }).ToList()
            };

            var directory = Path.GetDirectoryName(_localPath);
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new InvalidOperationException(
                    "The local Skills folder is unavailable.");
            }

            Directory.CreateDirectory(directory);
            var temporaryPath = _localPath + ".tmp";
            File.WriteAllText(
                temporaryPath,
                _serializer.Serialize(document),
                new UTF8Encoding(false));
            try
            {
                if (File.Exists(_localPath))
                {
                    File.Replace(
                        temporaryPath,
                        _localPath,
                        null,
                        true);
                }
                else
                {
                    File.Move(temporaryPath, _localPath);
                }
            }
            catch
            {
                try
                {
                    if (File.Exists(temporaryPath))
                    {
                        File.Delete(temporaryPath);
                    }
                }
                catch
                {
                }

                throw;
            }
        }

        public IReadOnlyList<SkillDefinition> GetForHost(
            string host)
        {
            var boundedHost = SkillDefinition.NormalizeHost(host);
            return LoadPublic()
                .Concat(LoadLocal())
                .Where(skill => string.Equals(
                    skill.Host,
                    boundedHost,
                    StringComparison.Ordinal))
                .ToArray();
        }

        public SkillDefinition Resolve(
            string origin,
            string id,
            string host)
        {
            var boundedOrigin = SkillDefinition.NormalizeOrigin(origin);
            var boundedId = TextBoundary.SingleLine(id, 80);
            var boundedHost = SkillDefinition.NormalizeHost(host);
            var source = boundedOrigin == "public"
                ? LoadPublic()
                : (boundedOrigin == "local"
                    ? LoadLocal()
                    : new SkillDefinition[0]);
            return source.FirstOrDefault(skill =>
                string.Equals(
                    skill.Id,
                    boundedId,
                    StringComparison.OrdinalIgnoreCase) &&
                string.Equals(
                    skill.Host,
                    boundedHost,
                    StringComparison.Ordinal));
        }

        public static SkillDefinition DuplicateToLocal(
            SkillDefinition source,
            IEnumerable<SkillDefinition> localSkills)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            var names = new HashSet<string>(
                (localSkills ?? Enumerable.Empty<SkillDefinition>())
                    .Where(skill => skill != null &&
                        string.Equals(
                            skill.Host,
                            source.Host,
                            StringComparison.Ordinal))
                    .Select(skill => skill.Name),
                StringComparer.OrdinalIgnoreCase);
            var baseName = TextBoundary.SingleLine(
                source.Name + " copy",
                SkillDefinition.MaxNameCharacters);
            var name = baseName;
            var suffix = 2;
            while (names.Contains(name))
            {
                var tail = " " + suffix.ToString(
                    CultureInfo.InvariantCulture);
                name = TextBoundary.SingleLine(
                    baseName,
                    SkillDefinition.MaxNameCharacters -
                    tail.Length) + tail;
                suffix++;
            }

            return new SkillDefinition
            {
                Id = Guid.NewGuid().ToString("N"),
                Name = name,
                Description = source.Description,
                Prompt = source.Prompt,
                Host = source.Host,
                Origin = "local",
                DisplayOrder = 0,
                StartFresh = source.StartFresh
            }.Sanitized("local");
        }

        public static string ExpandPrompt(string prompt)
        {
            return ExpandPrompt(
                prompt,
                DateTimeOffset.UtcNow,
                TimeZoneInfo.Local);
        }

        public static string ExpandPrompt(
            string prompt,
            DateTimeOffset now,
            TimeZoneInfo timeZone)
        {
            var zone = timeZone ?? TimeZoneInfo.Local;
            var localNow = TimeZoneInfo.ConvertTime(now, zone);
            var yesterdayFiveClock = new DateTime(
                localNow.Year,
                localNow.Month,
                localNow.Day,
                17,
                0,
                0,
                DateTimeKind.Unspecified).AddDays(-1);
            var yesterdayFive = new DateTimeOffset(
                yesterdayFiveClock,
                zone.GetUtcOffset(yesterdayFiveClock));
            var bounded = TextBoundary.PlainText(
                prompt,
                TextBoundary.MaxUserPromptCharacters);
            return bounded
                .Replace(
                    YesterdayFiveToken,
                    yesterdayFive.ToString(
                        "yyyy-MM-dd'T'HH:mm:sszzz",
                        CultureInfo.InvariantCulture))
                .Replace(
                    NowToken,
                    localNow.ToString(
                        "yyyy-MM-dd'T'HH:mm:sszzz",
                        CultureInfo.InvariantCulture));
        }

        private IReadOnlyList<SkillDefinition> ParseDocument(
            string json,
            string origin)
        {
            var result = new List<SkillDefinition>();
            try
            {
                var document = _serializer.DeserializeObject(json) as
                    IDictionary<string, object>;
                object versionValue;
                int version;
                if (document == null ||
                    !document.TryGetValue(
                        "SchemaVersion",
                        out versionValue) ||
                    !int.TryParse(
                        Convert.ToString(versionValue),
                        NumberStyles.Integer,
                        CultureInfo.InvariantCulture,
                        out version) ||
                    version != SchemaVersion)
                {
                    return result;
                }

                object skillsValue;
                if (!document.TryGetValue("Skills", out skillsValue))
                {
                    return result;
                }

                var entries = skillsValue as IEnumerable;
                if (entries == null)
                {
                    return result;
                }

                var ids = new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);
                var namesByHost = new HashSet<string>(
                    StringComparer.OrdinalIgnoreCase);
                var counts = new Dictionary<string, int>(
                    StringComparer.Ordinal);
                foreach (var raw in entries)
                {
                    var entry = raw as IDictionary<string, object>;
                    if (entry == null)
                    {
                        continue;
                    }

                    var skill = ReadEntry(entry).Sanitized(origin);
                    if (!SkillDefinition.IsValid(skill) ||
                        !ids.Add(skill.Id) ||
                        !namesByHost.Add(
                            skill.Host + "\n" + skill.Name))
                    {
                        continue;
                    }

                    int count;
                    counts.TryGetValue(skill.Host, out count);
                    if (skill.Origin == "local" &&
                        count >= MaxLocalSkillsPerHost)
                    {
                        continue;
                    }

                    counts[skill.Host] = count + 1;
                    result.Add(skill);
                }
            }
            catch
            {
                return new SkillDefinition[0];
            }

            return result;
        }

        private static IEnumerable<SkillDefinition> ValidateLocal(
            IEnumerable<SkillDefinition> skills)
        {
            var result = new List<SkillDefinition>();
            var ids = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            var namesByHost = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            var counts = new Dictionary<string, int>(
                StringComparer.Ordinal);
            foreach (var raw in skills ??
                Enumerable.Empty<SkillDefinition>())
            {
                var skill = raw?.Sanitized("local");
                if (!SkillDefinition.IsValid(skill))
                {
                    throw new InvalidOperationException(
                        "Each Local skill needs a name, prompt, and Office app.");
                }

                if (!ids.Add(skill.Id))
                {
                    throw new InvalidOperationException(
                        "Each Local skill needs a unique id.");
                }

                if (!namesByHost.Add(
                    skill.Host + "\n" + skill.Name))
                {
                    throw new InvalidOperationException(
                        "Local skill names must be unique within each Office app.");
                }

                int count;
                counts.TryGetValue(skill.Host, out count);
                if (count >= MaxLocalSkillsPerHost)
                {
                    throw new InvalidOperationException(
                        "Local skill limit reached for " +
                        skill.Host + " (" +
                        MaxLocalSkillsPerHost + ").");
                }

                counts[skill.Host] = count + 1;
                result.Add(skill);
            }

            return result;
        }

        private static SkillDefinition ReadEntry(
            IDictionary<string, object> entry)
        {
            return new SkillDefinition
            {
                Id = ReadString(entry, "Id"),
                Name = ReadString(entry, "Name"),
                Description = ReadString(entry, "Description"),
                Prompt = ReadString(entry, "Prompt"),
                Host = ReadString(entry, "Host"),
                DisplayOrder = ReadInteger(entry, "DisplayOrder"),
                StartFresh = ReadBoolean(entry, "StartFresh")
            };
        }

        private static string ReadString(
            IDictionary<string, object> entry,
            string key)
        {
            object value;
            return entry.TryGetValue(key, out value)
                ? Convert.ToString(value) ?? string.Empty
                : string.Empty;
        }

        private static int ReadInteger(
            IDictionary<string, object> entry,
            string key)
        {
            object value;
            int parsed;
            return entry.TryGetValue(key, out value) &&
                   int.TryParse(
                       Convert.ToString(value),
                       NumberStyles.Integer,
                       CultureInfo.InvariantCulture,
                       out parsed)
                ? parsed
                : 0;
        }

        private static bool ReadBoolean(
            IDictionary<string, object> entry,
            string key)
        {
            object value;
            bool parsed;
            return entry.TryGetValue(key, out value) &&
                   bool.TryParse(Convert.ToString(value), out parsed) &&
                   parsed;
        }

        private static string DefaultLocalPath()
        {
            return Path.Combine(
                Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData),
                "Scribble",
                "skills.json");
        }

        private sealed class StoredSkillDocument
        {
            public int SchemaVersion { get; set; }

            public List<StoredSkill> Skills { get; set; }
        }

        private sealed class StoredSkill
        {
            public string Id { get; set; }

            public string Name { get; set; }

            public string Description { get; set; }

            public string Prompt { get; set; }

            public string Host { get; set; }

            public int DisplayOrder { get; set; }

            public bool StartFresh { get; set; }
        }
    }
}
