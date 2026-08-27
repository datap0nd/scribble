using System.Collections.Generic;

namespace Scribble.Chat
{
    public static class DraftToolCatalog
    {
        public const string CreateDraft = "create_draft";
        public const string UpdateDraft = "update_draft";

        public static ChatToolDefinition CreateDefinition()
        {
            return Definition(
                CreateDraft,
                "Create, save, and display exactly one unsent Outlook draft. " +
                "This tool is available only after explicit one-request user " +
                "authorization and cannot send email.",
                true);
        }

        public static ChatToolDefinition UpdateDefinition()
        {
            return Definition(
                UpdateDraft,
                "Update and redisplay the single unsent Outlook draft already linked " +
                "to this chat. This cannot create another draft or send email.",
                false);
        }

        public static bool IsDraftTool(string name)
        {
            return IsCreateDraft(name) || IsUpdateDraft(name);
        }

        public static bool IsCreateDraft(string name)
        {
            return string.Equals(
                CreateDraft,
                name,
                System.StringComparison.Ordinal);
        }

        public static bool IsUpdateDraft(string name)
        {
            return string.Equals(
                UpdateDraft,
                name,
                System.StringComparison.Ordinal);
        }

        private static ChatToolDefinition Definition(
            string name,
            string description,
            bool includeKind)
        {
            // No maxLength/uniqueItems in these schemas: llama.cpp
            // compiles tool schemas into a sampling grammar, and
            // large bounded repetitions make sampler initialization
            // fail with HTTP 400. All limits are enforced locally by
            // DraftToolHost via TextBoundary, so the schema caps were
            // advisory only.
            var properties = new Dictionary<string, object>
            {
                {
                    "body",
                    BoundedString(
                        "Complete draft body. Use # heading, ## subheading, - list item, 1. numbered item, --- divider, and | cell | cell | table rows with a | --- | separator under the header row when a visual layout is requested. Never return raw HTML. Do not include commentary outside the draft.")
                },
                {
                    "subject",
                    BoundedString(
                        "Optional subject. Omit it to preserve the current subject during an update.")
                },
                {
                    "to",
                    BoundedString(
                        "Optional semicolon-separated recipients. Omit it to preserve current recipients during an update.")
                },
                {
                    "cc",
                    BoundedString(
                        "Optional semicolon-separated CC recipients. Omit it to preserve current recipients during an update.")
                },
                {
                    "bold_phrases",
                    new Dictionary<string, object>
                    {
                        { "type", "array" },
                        {
                            "description",
                            "Optional exact phrases from body to bold, at most twelve. Formatting is applied by the local host."
                        },
                        { "maxItems", 12 },
                        {
                            "items",
                            BoundedString(
                                "Exact body phrase to bold.")
                        }
                    }
                }
            };

            if (includeKind)
            {
                properties.Add(
                    "reply_handle",
                    BoundedString(
                        "Required when kind is reply. Use the exact opaque handle returned by search_mailbox, read_messages, read_thread, or selected. Omit for a new message. Never invent a handle."));
                properties.Add(
                    "kind",
                    new Dictionary<string, object>
                    {
                        { "type", "string" },
                        {
                            "description",
                            "Use reply for an exact email identified by reply_handle, or new for a new message."
                        },
                        { "enum", new[] { "new", "reply" } }
                    });
            }

            return new ChatToolDefinition
            {
                type = "function",
                function = new ChatToolFunctionDefinition
                {
                    name = name,
                    description = description +
                        " It must be the only tool call in its response.",
                    parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", properties },
                        {
                            "required",
                            includeKind
                                ? new[] { "kind", "body" }
                                : new[] { "body" }
                        },
                        { "additionalProperties", false }
                    }
                }
            };
        }

        private static Dictionary<string, object> BoundedString(
            string description)
        {
            return new Dictionary<string, object>
            {
                { "type", "string" },
                { "description", description }
            };
        }
    }
}
