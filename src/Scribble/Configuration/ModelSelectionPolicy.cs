using System;
using System.Collections.Generic;
using System.Linq;
using Scribble.Chat;

namespace Scribble.Configuration
{
    public static class ModelSelectionPolicy
    {
        public static bool IsGenerativeModel(string model)
        {
            var value = (model ?? string.Empty).Trim();
            return value.Length > 0 &&
                   !ModelCatalog.IsDisallowedModel(value) &&
                   (!AdminPolicy.GeminiDisabled ||
                    !GeminiCodeAssistGateway.IsGeminiModel(value)) &&
                   value.IndexOf(
                       "embedding",
                       StringComparison.OrdinalIgnoreCase) < 0;
        }

        public static string DescriptionFor(string model)
        {
            return ModelCatalog.DescribeForSelection(model);
        }

        // Qwen is the reviewed local-model baseline, not a vendor
        // lock. Prefer the Qwen3.8 27B family (including endpoint
        // aliases such as "fast") when available, then another
        // suitable Qwen model. A sole non-Qwen model is safe to
        // select automatically; otherwise the user makes the choice.
        public static string PreferredModel(
            IEnumerable<string> models)
        {
            var candidates = (models ??
                    Enumerable.Empty<string>())
                .Select(model => (model ?? string.Empty).Trim())
                .Where(IsGenerativeModel)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var qwen = candidates
                .Where(model =>
                    IsQwen(model) &&
                    !IsUnsuitableAutomaticChoice(model))
                .OrderByDescending(QwenPreferenceScore)
                .ThenBy(
                    model => model,
                    StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();
            if (qwen != null)
            {
                return qwen;
            }

            return candidates.Count == 1 &&
                   !IsUnsuitableAutomaticChoice(candidates[0])
                ? candidates[0]
                : string.Empty;
        }

        private static int QwenPreferenceScore(string model)
        {
            var score = 100;
            if (Contains(model, "qwen3.8") &&
                (Contains(model, "27b") ||
                 Contains(model, "-27")))
            {
                score += 180;
            }
            else if (Contains(model, "qwen3.8"))
            {
                score += 140;
            }
            else if (Contains(model, "qwen3.6"))
            {
                score += 100;
            }
            else if (Contains(model, "qwen3.5"))
            {
                score += 80;
            }

            if (Contains(model, "a3b"))
            {
                score += 20;
            }

            if (Contains(model, "27b") ||
                Contains(model, "-27"))
            {
                score += 30;
            }

            if (Contains(model, "instruct") ||
                Contains(model, "chat"))
            {
                score += 20;
            }

            return score;
        }

        private static bool IsUnsuitableAutomaticChoice(string model)
        {
            return Contains(model, "rerank") ||
                   Contains(model, "reward") ||
                   Contains(model, "distill") ||
                   Contains(model, "coder") ||
                   Contains(model, "image") ||
                   Contains(model, "audio") ||
                   Contains(model, "omni") ||
                   Contains(model, "tts") ||
                   Contains(model, "thinking") ||
                   Contains(model, "reasoning") ||
                   Contains(model, "-base") ||
                   Contains(model, "/base") ||
                   Contains(model, "_base");
        }

        private static bool IsQwen(string model)
        {
            return Contains(model, "qwen");
        }

        private static bool Contains(string value, string token)
        {
            return value.IndexOf(
                token,
                StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
