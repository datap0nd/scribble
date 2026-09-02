using System;
using System.Collections.Generic;
using Scribble.Outlook;

namespace Scribble.Chat
{
    public static class MailboxToolCatalog
    {
        public const string SearchMailbox = "search_mailbox";
        public const string ReadMessages = "read_messages";
        public const string ReadThread = "read_thread";

        public static readonly IReadOnlyList<string> ApprovedNames =
            new[]
            {
                SearchMailbox,
                ReadMessages,
                ReadThread
            };

        public static List<ChatToolDefinition> CreateDefinitions(
            bool workingSetOnly = false)
        {
            var definitions = new List<ChatToolDefinition>
            {
                new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = SearchMailbox,
                        description =
                            "Search the user's local primary Outlook Inbox and Sent Items. " +
                            "Returns bounded metadata, snippets, and temporary message " +
                            "handles, newest first. Only one search is allowed per request, " +
                            "so choose the query carefully before calling.",
                        parameters = ObjectSchema(
                            new Dictionary<string, object>
                            {
                                {
                                    "query",
                                    StringSchema(
                                        "Words or a phrase to find in subject, body, sender, or recipients. " +
                                        "Use an empty string to list recent mail.")
                                },
                                {
                                    "folder",
                                    EnumSchema(
                                        "Mailbox folders to search.",
                                        "all",
                                        "inbox",
                                        "sent")
                                },
                                {
                                    "days_back",
                                    IntegerSchema(
                                        "Only include messages this many days old or newer. " +
                                        "Defaults to 365 when omitted; stay generous unless " +
                                        "the user names a time range.",
                                        1,
                                        3650)
                                },
                                {
                                    "received_after",
                                    StringSchema(
                                        "Optional inclusive ISO-8601 lower timestamp with UTC offset. When supplied, it replaces days_back as the lower bound.")
                                },
                                {
                                    "received_before",
                                    StringSchema(
                                        "Optional inclusive ISO-8601 upper timestamp with UTC offset.")
                                },
                                {
                                    "unread_only",
                                    BooleanSchema(
                                        "When true, return only messages that are currently unread. Reading through this tool never changes their read state.")
                                },
                                {
                                    "max_results",
                                    IntegerSchema(
                                        "Maximum number of result summaries to return.",
                                        1,
                                        MailboxWorkingSet.MaxMessages)
                                }
                            },
                            "query")
                    }
                },
                new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = ReadMessages,
                        description =
                            "Load the bounded plain-text bodies of messages returned by " +
                            "search_mailbox, the selected email, or the user-approved " +
                            "working set. At most " +
                            MailboxWorkingSet.MaxMessages +
                            " unique message bodies " +
                            "can be loaded in one request. Attachments are included as " +
                            "extracted text where possible (Excel, PDF, PowerPoint, " +
                            "Word including legacy formats, RTF, and text files); " +
                            "unreadable types are noted. Image attachments are " +
                            "delivered to vision-capable models as image input right " +
                            "after this tool result, so call this tool when the user " +
                            "asks about an image in a message.",
                        parameters = ObjectSchema(
                            new Dictionary<string, object>
                            {
                                {
                                    "handles",
                                    new Dictionary<string, object>
                                    {
                                        { "type", "array" },
                                        {
                                            "items",
                                            StringSchema(
                                                "A temporary handle returned by search_mailbox, " +
                                                "selected, or context1 through context" +
                                                MailboxWorkingSet.MaxMessages +
                                                " from the " +
                                                "user-approved working set.")
                                        },
                                        { "minItems", 1 },
                                        { "maxItems", MailboxWorkingSet.MaxMessages }
                                    }
                                }
                            },
                            "handles")
                    }
                },
                new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = ReadThread,
                        description =
                            "Load bounded messages in the Outlook conversation containing " +
                            "a searched or selected message, subject to the " +
                            MailboxWorkingSet.MaxMessages +
                            "-message request-wide context cap.",
                        parameters = ObjectSchema(
                            new Dictionary<string, object>
                            {
                                {
                                    "handle",
                                    StringSchema(
                                        "A temporary handle returned by search_mailbox, " +
                                        "or selected for the currently selected email.")
                                }
                            },
                            "handle")
                    }
                }
            };

            return workingSetOnly
                ? new List<ChatToolDefinition>
                {
                    definitions[1]
                }
                : definitions;
        }

        public static bool IsApproved(string name)
        {
            foreach (var approved in ApprovedNames)
            {
                if (string.Equals(
                    approved,
                    name,
                    StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        private static Dictionary<string, object> ObjectSchema(
            Dictionary<string, object> properties,
            params string[] required)
        {
            return new Dictionary<string, object>
            {
                { "type", "object" },
                { "properties", properties },
                { "required", required },
                { "additionalProperties", false }
            };
        }

        private static Dictionary<string, object> StringSchema(
            string description)
        {
            return new Dictionary<string, object>
            {
                { "type", "string" },
                { "description", description }
            };
        }

        private static Dictionary<string, object> EnumSchema(
            string description,
            params string[] values)
        {
            return new Dictionary<string, object>
            {
                { "type", "string" },
                { "description", description },
                { "enum", values }
            };
        }

        private static Dictionary<string, object> IntegerSchema(
            string description,
            int minimum,
            int maximum)
        {
            return new Dictionary<string, object>
            {
                { "type", "integer" },
                { "description", description },
                { "minimum", minimum },
                { "maximum", maximum }
            };
        }

        private static Dictionary<string, object> BooleanSchema(
            string description)
        {
            return new Dictionary<string, object>
            {
                { "type", "boolean" },
                { "description", description }
            };
        }
    }
}
