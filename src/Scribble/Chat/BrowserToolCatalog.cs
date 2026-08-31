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
        public const string ClickControl = "browser_click";
        public const string AskUser = "ask_user";
        public const string OpenOutlookDraft = "open_outlook_draft";
        public const string OpenExcelTable = "open_excel_table";

        public static readonly IReadOnlyList<string> ApprovedNames =
            new[]
            {
                NavigatePage,
                ReadPage,
                ClickControl,
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
                    ClickControl,
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
                },
                new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = ClickControl,
                        description =
                            "Click one visible button or link on the current page by " +
                            "its exact visible text, then return the page as it looks " +
                            "after the click. Use it only to get past benign " +
                            "interstitials that block reading: cookie or consent " +
                            "banners, location/country/language choosers (pick the " +
                            "option matching the user's request), continue/accept/" +
                            "close. Clicks that buy, pay, check out, add to cart, " +
                            "sign in, register, subscribe, or delete are refused, and " +
                            "typing into fields is impossible.",
                        parameters = new Dictionary<string, object>
                        {
                            { "type", "object" },
                            {
                                "properties",
                                new Dictionary<string, object>
                                {
                                    {
                                        "text",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "string" },
                                            {
                                                "description",
                                                "The visible text of the control to click, e.g. " +
                                                "\"United Arab Emirates\" or \"Accept all\"."
                                            }
                                        }
                                    }
                                }
                            },
                            { "required", new[] { "text" } }
                        }
                    }
                },
                new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = AskUser,
                        description =
                            "Ask the user one short clarifying question and wait for " +
                            "their answer. Provide 2-6 short options they can click; " +
                            "they may also type a different answer. Use it BEFORE " +
                            "doing many steps when something important is ambiguous - " +
                            "location or country, which recipient, budget, scope, " +
                            "which product variant. Ask one question at a time.",
                        parameters = new Dictionary<string, object>
                        {
                            { "type", "object" },
                            {
                                "properties",
                                new Dictionary<string, object>
                                {
                                    {
                                        "question",
                                        new Dictionary<string, object>
                                        {
                                            { "type", "string" },
                                            {
                                                "description",
                                                "The question to show the user, one sentence."
                                            }
                                        }
                                    },
                                    {
                                        "options",
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
                                                "2-6 short clickable answers, most likely first."
                                            }
                                        }
                                    }
                                }
                            },
                            { "required", new[] { "question" } }
                        }
                    }
                }
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
