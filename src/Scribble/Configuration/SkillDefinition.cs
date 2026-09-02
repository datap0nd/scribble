using System;
using Scribble.Security;

namespace Scribble.Configuration
{
    public sealed class SkillDefinition
    {
        public const int MaxNameCharacters = 60;
        public const int MaxDescriptionCharacters = 240;

        public string Id { get; set; } = string.Empty;

        public string Name { get; set; } = string.Empty;

        public string Description { get; set; } = string.Empty;

        public string Prompt { get; set; } = string.Empty;

        public string Host { get; set; } = string.Empty;

        public string Origin { get; set; } = string.Empty;

        public int DisplayOrder { get; set; }

        public bool StartFresh { get; set; }

        public SkillDefinition Sanitized(string origin)
        {
            var boundedOrigin = NormalizeOrigin(origin);
            return new SkillDefinition
            {
                Id = TextBoundary.SingleLine(Id, 80),
                Name = TextBoundary.SingleLine(
                    Name,
                    MaxNameCharacters),
                Description = TextBoundary.SingleLine(
                    Description,
                    MaxDescriptionCharacters),
                Prompt = TextBoundary.PlainText(
                    Prompt,
                    TextBoundary.MaxUserPromptCharacters),
                Host = NormalizeHost(Host),
                Origin = boundedOrigin,
                DisplayOrder = Math.Max(0, DisplayOrder),
                StartFresh = StartFresh
            };
        }

        public static bool IsValid(SkillDefinition skill)
        {
            return skill != null &&
                   skill.Id.Length > 0 &&
                   skill.Name.Length > 0 &&
                   skill.Prompt.Length > 0 &&
                   skill.Host.Length > 0 &&
                   skill.Origin.Length > 0;
        }

        public static string NormalizeHost(string host)
        {
            var value = TextBoundary.SingleLine(host, 20)
                .ToLowerInvariant();
            return value == "outlook" ||
                   value == "excel" ||
                   value == "powerpoint" ||
                   value == "word"
                ? value
                : string.Empty;
        }

        public static string NormalizeOrigin(string origin)
        {
            return string.Equals(
                    origin,
                    "public",
                    StringComparison.OrdinalIgnoreCase)
                ? "public"
                : (string.Equals(
                        origin,
                        "local",
                        StringComparison.OrdinalIgnoreCase)
                    ? "local"
                    : string.Empty);
        }
    }
}
