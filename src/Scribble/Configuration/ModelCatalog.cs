using System;
using System.Collections.Generic;
using System.Linq;

namespace Scribble.Configuration
{
    public sealed class ModelGuideEntry
    {
        internal ModelGuideEntry(
            string exampleId,
            string[] matchTokens,
            string listLabel,
            string speed,
            string quality,
            bool readsEmailImages,
            string notes)
        {
            ExampleId = exampleId ?? string.Empty;
            MatchTokens = matchTokens ?? new string[0];
            ListLabel = listLabel ?? string.Empty;
            Speed = speed ?? string.Empty;
            Quality = quality ?? string.Empty;
            ReadsEmailImages = readsEmailImages;
            Notes = notes ?? string.Empty;
        }

        public string ExampleId { get; }

        public string ListLabel { get; }

        public string Speed { get; }

        public string Quality { get; }

        public bool ReadsEmailImages { get; }

        public string Notes { get; }

        internal string[] MatchTokens { get; }

        public string SummaryLine
        {
            get { return ListLabel; }
        }
    }

    public static class ModelCatalog
    {
        private static readonly string[] VisionNameTokens =
        {
            "vl",
            "vision",
            "llava",
            "pixtral",
            "minicpm-v",
            "internvl",
            "moondream",
            "smolvlm",
            "multimodal",
            "gemini",
            "gemma-3",
            "gemma3",
            "gemma-4",
            "gemma4"
        };

        private static readonly ModelGuideEntry[] MasterList =
        {
            new ModelGuideEntry(
                "gemini-3.5-flash",
                new[] { "gemini", "flash", "3" },
                "Newest fast cloud Gemini",
                "Fast",
                "Excellent",
                true,
                "Latest flash generation over Google sign-in with its own quota bucket, separate from the 2.5 models. Availability depends on your organization's tier - if a request is rejected, pick another model."),
            new ModelGuideEntry(
                "gemini-2.5-flash",
                new[] { "gemini", "flash" },
                "Fast cloud Gemini - recommended",
                "Fast",
                "Excellent",
                true,
                "Best speed/quality balance over Google sign-in. Scribble disables its internal thinking for snappy replies and streams tokens live. Higher quotas than pro."),
            new ModelGuideEntry(
                "gemini-2.5-pro",
                new[] { "gemini", "pro" },
                "Deepest cloud Gemini, slower",
                "Slow",
                "Excellent",
                true,
                "Highest quality for complex drafting and reasoning. Tighter per-minute quotas and slower responses; Scribble keeps its thinking at the minimum the model allows."),
            new ModelGuideEntry(
                "gemini-2.5-flash-lite",
                new[] { "gemini", "flash", "lite" },
                "Fastest cloud Gemini",
                "Fast",
                "Good",
                true,
                "Lowest latency of the Gemini family. Fine for quick questions and summaries; weaker on multi-step tool use than flash."),
            new ModelGuideEntry(
                "qwen3-vl-30b",
                new[] { "qwen", "vl" },
                "Vision model for email images",
                "Medium",
                "Very good",
                true,
                "Preferred vision model. Receives email images through Scribble vision input. Best for screenshots, scans, and inline photos."),
            new ModelGuideEntry(
                "qwen3.6-35b-a3b",
                new[] { "qwen", "35", "a3b" },
                "Balanced everyday mailbox model",
                "Medium",
                "Very good",
                false,
                "Strong mailbox tool use and drafting. Good default for Excel/CSV attachments and everyday mail questions."),
            new ModelGuideEntry(
                "qwen3.8-27b",
                new[] { "qwen", "27" },
                "Lighter Qwen text model",
                "Medium",
                "Very good",
                false,
                "Balanced text model when the 35B route feels heavy. Handles spreadsheets; not for image understanding."),
            new ModelGuideEntry(
                "gemma-4-31b-it",
                new[] { "gemma", "31" },
                "Multimodal Gemma drafting model",
                "Medium",
                "Very good",
                true,
                "Higher-quality Gemma option for longer answers and careful drafting. " +
                "Reads email images when the server exposes Gemma's vision input."),
            new ModelGuideEntry(
                "gemma-4-26b-a4b-it",
                new[] { "gemma", "26", "a4b" },
                "Fast multimodal Gemma for quick asks",
                "Fast",
                "Good",
                true,
                "Lightweight Gemma route for quick mailbox questions and spreadsheet attachments. " +
                "Reads email images when the server exposes Gemma's vision input."),
            new ModelGuideEntry(
                "gpt-oss-20b",
                new[] { "gpt", "oss", "20" },
                "Fastest general text model",
                "Fast",
                "Good",
                false,
                "Fastest general option for simple searches, summaries, and spreadsheet attachments."),
            new ModelGuideEntry(
                "gpt-oss-120b",
                new[] { "gpt", "oss", "120" },
                "Best quality, slowest model",
                "Slow",
                "Best",
                false,
                "Highest-quality text model here, but expect much longer waits on large mail context.")
        };

        public static IReadOnlyList<ModelGuideEntry> GuideEntries
        {
            get { return MasterList; }
        }

        public static bool IsDisallowedModel(string model)
        {
            return (model ?? string.Empty).IndexOf(
                "gauss",
                StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool SupportsVision(string model)
        {
            return IsVisionCapable(model);
        }

        public static bool IsVisionCapable(string model)
        {
            var value = (model ?? string.Empty).Trim();
            if (value.Length == 0 ||
                IsDisallowedModel(value))
            {
                return false;
            }

            var profile = Resolve(value);
            if (profile?.ReadsEmailImages == true)
            {
                return true;
            }

            var normalized = Normalize(value);
            if (normalized.IndexOf(
                    "embedding",
                    StringComparison.Ordinal) >= 0)
            {
                return false;
            }

            foreach (var token in VisionNameTokens)
            {
                if (normalized.IndexOf(
                        token,
                        StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        public static ModelGuideEntry Resolve(string model)
        {
            var normalized = Normalize(model);
            if (normalized.Length == 0)
            {
                return null;
            }

            ModelGuideEntry best = null;
            var bestScore = 0;
            foreach (var entry in MasterList)
            {
                var score = MatchScore(normalized, entry.MatchTokens);
                if (score > bestScore)
                {
                    bestScore = score;
                    best = entry;
                }
            }

            return bestScore > 0 ? best : null;
        }

        public static string DescribeForSelection(string model)
        {
            var value = (model ?? string.Empty).Trim();
            if (value.Length == 0)
            {
                return BuildGuideOverview();
            }

            if (IsDisallowedModel(value))
            {
                return "This model is not available in Scribble.";
            }

            var profile = Resolve(value);
            if (profile == null)
            {
                return "Custom model. It must support OpenAI-compatible chat tool calls. " +
                    (IsVisionCapable(value)
                        ? "Its name looks vision-capable, so email images are sent as image input."
                        : "Its name does not look vision-capable, so email images stay filename-only.") +
                    " Spreadsheets are sent as extracted text for any model.";
            }

            return profile.ListLabel;
        }

        public static string FindBestVisionModel(
            IEnumerable<string> discoveredModels)
        {
            var models = (discoveredModels ?? Enumerable.Empty<string>())
                .Where(model =>
                    model != null &&
                    model.Trim().Length > 0 &&
                    !IsDisallowedModel(model) &&
                    ModelSelectionPolicy.IsGenerativeModel(model))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var entry in MasterList)
            {
                if (!entry.ReadsEmailImages)
                {
                    continue;
                }

                var match = FindMatchingDiscoveredId(entry, models);
                if (match.Length > 0)
                {
                    return match;
                }
            }

            return models
                .Where(IsVisionCandidate)
                .OrderBy(model => model, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault() ??
                string.Empty;
        }

        public static bool IsVisionCandidate(string model)
        {
            return IsVisionCapable(model);
        }

        public static string FindMatchingDiscoveredId(
            ModelGuideEntry entry,
            IEnumerable<string> discoveredModels)
        {
            if (entry == null)
            {
                return string.Empty;
            }

            foreach (var model in discoveredModels ??
                Enumerable.Empty<string>())
            {
                if (Resolve(model) == entry)
                {
                    return model;
                }
            }

            return string.Empty;
        }

        public static string BuildGuideOverview()
        {
            return "Refresh models, then pick from the guide.";
        }

        private static int MatchScore(
            string normalizedModel,
            string[] tokens)
        {
            if (tokens == null || tokens.Length == 0)
            {
                return 0;
            }

            var score = 0;
            foreach (var token in tokens)
            {
                var normalizedToken = Normalize(token);
                if (normalizedToken.Length == 0)
                {
                    continue;
                }

                if (normalizedModel.IndexOf(
                        normalizedToken,
                        StringComparison.Ordinal) >= 0)
                {
                    score++;
                }
            }

            return score == tokens.Length ? score : 0;
        }

        private static string Normalize(string value)
        {
            return (value ?? string.Empty)
                .Trim()
                .ToLowerInvariant()
                .Replace(" ", string.Empty)
                .Replace("_", "-");
        }
    }
}
