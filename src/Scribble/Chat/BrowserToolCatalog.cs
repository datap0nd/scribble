using System;
using System.Collections.Generic;

namespace Scribble.Chat
{
    // Tools the browser companion exposes to the model. Navigation
    // and page reading execute inside the extension (the side panel
    // drives the user's own visible tab through chrome.tabs); the
    // Outlook draft tool executes in the native host and only ever
    // opens an unsent draft window for the user's review.
    public static class BrowserToolCatalog
    {
        public const string NavigatePage = "browser_navigate";
        public const string ReadPage = "browser_read_page";
        public const string OpenOutlookDraft = "open_outlook_draft";

        public static readonly IReadOnlyList<string> ApprovedNames =
            new[]
            {
                NavigatePage,
                ReadPage,
                OpenOutlookDraft
            };

        // Tools the extension itself must execute with chrome APIs.
        public static bool IsBrowserExecuted(string name)
        {
            return string.Equals(
                    name,
                    NavigatePage,
                    StringComparison.Ordinal) ||
                string.Equals(
                    name,
                    ReadPage,
                    StringComparison.Ordinal);
        }

        public static bool IsApproved(string name)
        {
            foreach (var approved in ApprovedNames)
            {
                if (string.Equals(
                    name,
                    approved,
                    StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        public static List<ChatToolDefinition> CreateDefinitions(
            bool allowOutlookDraft)
        {
            var definitions = new List<ChatToolDefinition>
            {
                new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = NavigatePage,
                        description =
                            "Navigate the user's active browser tab to an http or " +
                            "https URL and return the loaded page's readable text, " +
                            "title, final URL, and a bounded list of the page's " +
                            "links. The navigation is visible to the user in their " +
                            "own tab. For a multi-step task, work one page at a " +
                            "time: open the site's own search-results URL first " +
                            "(for example /s?k=your+terms on Amazon), then call " +
                            "this tool again with an exact product or article URL " +
                            "picked from the returned <links> list. " +
                            "It cannot click, fill forms, sign in, purchase, " +
                            "download, or upload.",
                        parameters = new Dictionary<string, object>
                        {
                            { "type", "object" },
                            {
                                "properties",
                                new Dictionary<string, object>
                                {
                                    {
                                        "url",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "string" },
                                            {
                                                "description",
                                                "Absolute http:// or https:// URL to open."
                                            }
                                        }
                                    }
                                }
                            },
                            { "required", new[] { "url" } }
                        }
                    }
                },
                new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = ReadPage,
                        description =
                            "Re-read the active browser tab's current readable text, " +
                            "title, URL, and link list. Use it when the page has " +
                            "changed since the attached context was captured.",
                        parameters = new Dictionary<string, object>
                        {
                            { "type", "object" },
                            {
                                "properties",
                                new Dictionary<string, object>()
                            }
                        }
                    }
                }
            };

            if (allowOutlookDraft)
            {
                definitions.Add(new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = OpenOutlookDraft,
                        description =
                            "Open one unsent Outlook draft window on the user's " +
                            "desktop with the given recipients, subject, and plain-" +
                            "text body. The draft is never sent; the user reviews, " +
                            "edits, and sends it themselves. Available only because " +
                            "the user's own latest message asked for a draft; at " +
                            "most one draft per request.",
                        parameters = new Dictionary<string, object>
                        {
                            { "type", "object" },
                            {
                                "properties",
                                new Dictionary<string, object>
                                {
                                    {
                                        "to",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "string" },
                                            {
                                                "description",
                                                "Semicolon-separated recipients; may be empty."
                                            }
                                        }
                                    },
                                    {
                                        "cc",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "string" },
                                            {
                                                "description",
                                                "Semicolon-separated CC recipients; may be empty."
                                            }
                                        }
                                    },
                                    {
                                        "subject",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "string" },
                                            { "description", "Draft subject line." }
                                        }
                                    },
                                    {
                                        "body",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "string" },
                                            {
                                                "description",
                                                "Plain-text draft body. Markdown pipe tables render as text."
                                            }
                                        }
                                    }
                                }
                            },
                            { "required", new[] { "body" } }
                        }
                    }
                });
            }

            return definitions;
        }
    }
}
