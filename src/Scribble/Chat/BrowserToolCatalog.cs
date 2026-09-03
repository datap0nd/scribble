using System;
using System.Collections.Generic;

namespace Scribble.Chat
{
    // Tools the browser companion exposes to the model. Navigation
    // and page reading execute inside the extension (the active tab
    // is read-only context; actions use inactive work tabs); the
    // Outlook draft tool executes in the native host and only ever
    // opens an unsent draft window for the user's review.
    public static class BrowserToolCatalog
    {
        public const string NavigatePage = "browser_navigate";
        public const string ReadPage = "browser_read_page";
        public const string SearchGoogle = "browser_search_google";
        public const string SnapshotPage = "browser_snapshot";
        public const string ActOnPage = "browser_act";
        public const string AskUser = PromptHelperTool.Name;
        public const string OpenOutlookDraft = "open_outlook_draft";
        public const string OpenExcelTable = "open_excel_table";

        public static readonly IReadOnlyList<string> ApprovedNames =
            new[]
            {
                NavigatePage,
                ReadPage,
                SearchGoogle,
                SnapshotPage,
                ActOnPage,
                AskUser,
                OpenOutlookDraft,
                OpenExcelTable
            };

        // Tools the extension itself must execute with chrome APIs
        // (or, for ask_user, by waiting on the person at the panel).
        public static bool IsBrowserExecuted(string name)
        {
            return string.Equals(
                    name,
                    NavigatePage,
                    StringComparison.Ordinal) ||
                string.Equals(
                    name,
                    ReadPage,
                    StringComparison.Ordinal) ||
                string.Equals(
                    name,
                    SearchGoogle,
                    StringComparison.Ordinal) ||
                string.Equals(
                    name,
                    SnapshotPage,
                    StringComparison.Ordinal) ||
                string.Equals(
                    name,
                    ActOnPage,
                    StringComparison.Ordinal) ||
                string.Equals(
                    name,
                    AskUser,
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

        public static List<ChatToolDefinition> CreateDefinitions()
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
                            "Open an http or https URL in one of Scribble's own " +
                            "work tabs (up to 5, background tabs next to the " +
                            "user's - their current tab is never navigated away) " +
                            "and return the loaded page. The URL must have appeared " +
                            "literally in the user's request or a clarification " +
                            "answer; model-constructed search and destination URLs " +
                            "are refused. Use browser_search_google for discovery, " +
                            "then browser_snapshot and browser_act to click an " +
                            "observed result by ref. It cannot sign in, purchase, " +
                            "download, upload, or act on the user's active tab.",
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
                                    },
                                    {
                                        "tab",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "integer" },
                                            {
                                                "description",
                                                "Work tab number 1-5. Defaults to the last used " +
                                                "work tab (or opens tab 1)."
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
                            "Re-read a page's current readable text, title, URL, " +
                            "and link list: a Scribble work tab when tab is given " +
                            "(default: the last used work tab), or the user's " +
                            "active tab when no work tab exists. Use it when the " +
                            "page has changed since it was last read.",
                        parameters = new Dictionary<string, object>
                        {
                            { "type", "object" },
                            {
                                "properties",
                                new Dictionary<string, object>
                                {
                                    {
                                        "tab",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "integer" },
                                            {
                                                "description",
                                                "Work tab number 1-5; omit for the last used " +
                                                "work tab or the user's active tab."
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                },
                new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = SearchGoogle,
                        description =
                            "Search Google through its visible UI in a Scribble-owned " +
                            "background work tab. The extension opens Google's home " +
                            "page, types the query into the search field, submits it, " +
                            "and returns an interactive snapshot of the results. Query " +
                            "words must come from the user's request or clarification " +
                            "answers. Analyze the results, then click the best result " +
                            "with browser_act instead of guessing its URL.",
                        parameters = new Dictionary<string, object>
                        {
                            { "type", "object" },
                            { "additionalProperties", false },
                            {
                                "properties",
                                new Dictionary<string, object>
                                {
                                    {
                                        "query",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "string" },
                                            {
                                                "description",
                                                "Google query, at most 200 characters, composed only from user-supplied words."
                                            }
                                        }
                                    },
                                    {
                                        "tab",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "integer" },
                                            {
                                                "description",
                                                "Work tab number 1-5; defaults to the last work tab or tab 1."
                                            }
                                        }
                                    }
                                }
                            },
                            { "required", new[] { "query" } }
                        }
                    }
                },
                new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = SnapshotPage,
                        description =
                            "Inspect a Scribble-owned work tab and return bounded " +
                            "visible page text plus visible interactive controls. " +
                            "Each control has an opaque ref, role, accessible name, " +
                            "state, safe value summary, link target, and viewport " +
                            "status. Use an optional query to return only matching " +
                            "controls. Refs expire when the document navigates. " +
                            "Sensitive field values are never read.",
                        parameters = new Dictionary<string, object>
                        {
                            { "type", "object" },
                            { "additionalProperties", false },
                            {
                                "properties",
                                new Dictionary<string, object>
                                {
                                    {
                                        "tab",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "integer" },
                                            { "description", "Open work tab number 1-5." }
                                        }
                                    },
                                    {
                                        "query",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "string" },
                                            { "description", "Optional visible-name or role filter." }
                                        }
                                    }
                                }
                            }
                        }
                    }
                },
                new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = ActOnPage,
                        description =
                            "Perform one bounded action in a Scribble-owned work tab, " +
                            "using a ref from the latest browser_snapshot, then return " +
                            "a fresh snapshot. Actions: click, type, select, check, " +
                            "press, hover, scroll, wait. Typed text must be at most " +
                            "200 characters and copied directly from the user's request " +
                            "or a clarification answer. Credential, personal-data, " +
                            "payment, booking, purchase, messaging, upload, download, " +
                            "and destructive actions are refused.",
                        parameters = new Dictionary<string, object>
                        {
                            { "type", "object" },
                            { "additionalProperties", false },
                            {
                                "properties",
                                new Dictionary<string, object>
                                {
                                    { "tab", new Dictionary<string, object> { { "type", "integer" } } },
                                    {
                                        "action",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "string" },
                                            { "enum", new[] { "click", "type", "select", "check", "press", "hover", "scroll", "wait" } }
                                        }
                                    },
                                    { "ref", new Dictionary<string, object> { { "type", "string" } } },
                                    { "value", new Dictionary<string, object> { { "type", "string" } } },
                                    {
                                        "source",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "string" },
                                            { "description", "user_prompt or clarification_answer; used for typed-value provenance." }
                                        }
                                    },
                                    {
                                        "key",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "string" },
                                            { "description", "For press: Enter, Escape, Tab, Backspace, Delete, Space, arrow, Home, End, PageUp, or PageDown." }
                                        }
                                    },
                                    { "direction", new Dictionary<string, object> { { "type", "string" }, { "enum", new[] { "up", "down", "left", "right" } } } },
                                    {
                                        "amount",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "integer" },
                                            { "description", "Pixels for scroll (100-2000) or milliseconds for wait (250-5000)." }
                                        }
                                    }
                                }
                            },
                            { "required", new[] { "action" } }
                        }
                    }
                },
                PromptHelperTool.CreateDefinition()
            };

            {
                definitions.Add(new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = OpenExcelTable,
                        description =
                            "Open a brand-new, unsaved Excel workbook on the user's " +
                            "desktop containing one table built from your rows, with " +
                            "an optional native chart. Use it whenever the user asks " +
                            "to put results in Excel, a spreadsheet, a workbook, or " +
                            "a table outside the chat. Nothing is saved; the user " +
                            "reviews the workbook themselves. At most one workbook " +
                            "per request.",
                        parameters = new Dictionary<string, object>
                        {
                            { "type", "object" },
                            {
                                "properties",
                                new Dictionary<string, object>
                                {
                                    {
                                        "title",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "string" },
                                            {
                                                "description",
                                                "Optional title written above the table."
                                            }
                                        }
                                    },
                                    {
                                        "columns",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "array" },
                                            {
                                                "items",
                                                new Dictionary<string, object>
                                                {
                                                    { "type", "string" }
                                                }
                                            },
                                            {
                                                "description",
                                                "Column headers, up to 20."
                                            }
                                        }
                                    },
                                    {
                                        "rows",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "array" },
                                            {
                                                "items",
                                                new Dictionary<string, object>
                                                {
                                                    { "type", "array" },
                                                    {
                                                        "items",
                                                        new Dictionary<string, object>
                                                        {
                                                            { "type", "string" }
                                                        }
                                                    }
                                                }
                                            },
                                            {
                                                "description",
                                                "Data rows matching the headers, up to 500. " +
                                                "Numeric strings become real numbers."
                                            }
                                        }
                                    },
                                    {
                                        "chart_kind",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "string" },
                                            {
                                                "description",
                                                "Optional chart: column, bar, line, or pie. " +
                                                "Omit or use none for no chart."
                                            }
                                        }
                                    },
                                    {
                                        "chart_title",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "string" },
                                            { "description", "Optional chart title." }
                                        }
                                    }
                                }
                            },
                            { "required", new[] { "columns", "rows" } }
                        }
                    }
                });
            }

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
                            "text body. Use it whenever the user asks to email, " +
                            "message, or send something to someone. The draft is " +
                            "never sent by this tool; the user reviews, edits, and " +
                            "sends it themselves, so drafting is always safe. " +
                            "Recipients may be plain names - Outlook resolves them " +
                            "from the user's contacts. At most one draft per " +
                            "request.",
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
