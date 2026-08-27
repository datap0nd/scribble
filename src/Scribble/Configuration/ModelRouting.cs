using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Scribble.Chat;
using Scribble.Outlook;

namespace Scribble.Configuration
{
    public static class ModelRouting
    {
        public static bool ContextMayIncludeImages(
            MessageSnapshot selectedMessage,
            IReadOnlyList<MessageSnapshot> workingMessages)
        {
            if (MessageHasImageAttachment(selectedMessage))
            {
                return true;
            }

            var workingSet = MailboxWorkingSet.Normalize(
                workingMessages);
            foreach (var message in workingSet)
            {
                if (MessageHasImageAttachment(message))
                {
                    return true;
                }
            }

            return false;
        }

        public static bool ToolResultsIncludeImages(
            IReadOnlyList<MailboxToolResult> toolResults)
        {
            if (toolResults == null)
            {
                return false;
            }

            foreach (var result in toolResults)
            {
                if (result?.VisionImages != null &&
                    result.VisionImages.Count > 0)
                {
                    return true;
                }
            }

            return false;
        }

        public static bool MessageHasImageAttachment(
            MessageSnapshot message)
        {
            if (message?.AttachmentNames == null)
            {
                return false;
            }

            foreach (var name in message.AttachmentNames)
            {
                if (IsImageAttachmentName(name))
                {
                    return true;
                }
            }

            return false;
        }

        public static string ResolveForRequest(
            AppSettings settings,
            bool imagesExpected,
            IReadOnlyList<MailboxToolResult> loadedImages = null)
        {
            var preferred = (settings?.Model ?? string.Empty).Trim();
            if (preferred.Length == 0 || settings == null)
            {
                return preferred;
            }

            var needsVision =
                imagesExpected ||
                ToolResultsIncludeImages(loadedImages);
            if (!needsVision)
            {
                return preferred;
            }

            if (ModelCatalog.IsVisionCapable(preferred))
            {
                return preferred;
            }

            if (!settings.SwitchToVisionModelForImages)
            {
                return preferred;
            }

            var candidates = (settings.DiscoveredModels ??
                new List<string>())
                .Concat(new[] { preferred });
            var visionModel = ModelCatalog.FindBestVisionModel(
                candidates);
            return visionModel.Length > 0
                ? visionModel
                : preferred;
        }

        public static bool IsTemporaryVisionSwitch(
            AppSettings settings,
            string activeModel)
        {
            if (settings == null ||
                !settings.SwitchToVisionModelForImages)
            {
                return false;
            }

            var preferred = settings.Model.Trim();
            var active = (activeModel ?? string.Empty).Trim();
            return preferred.Length > 0 &&
                   active.Length > 0 &&
                   !string.Equals(
                       preferred,
                       active,
                       StringComparison.OrdinalIgnoreCase) &&
                   ModelCatalog.IsVisionCapable(active);
        }

        private static bool IsImageAttachmentName(string fileName)
        {
            return EmailAttachmentReader.IsImageExtension(
                Path.GetExtension(fileName ?? string.Empty));
        }
    }
}
