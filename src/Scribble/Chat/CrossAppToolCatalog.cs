using System;
using System.Collections.Generic;
using Scribble.Office;

namespace Scribble.Chat
{
    // Cross-application draft tools for the Excel and PowerPoint
    // panes: hand content to a sibling Office app as a clearly
    // marked, unsent, unsaved draft. Email drafts open for review in
    // Outlook and can never be sent; slide and sheet drafts follow
    // the same marked-draft rules as the in-host writers. All three
    // are offered only when the user's own prompt authorized a
    // draft, and they share the same one-shot permission.
    public static class CrossAppToolCatalog
    {
        public const string CreateEmailDraft = "create_email_draft";
        public const string SendToPowerPoint = "send_to_powerpoint";
        public const string SendToExcel = "send_to_excel";
        public const string SendToWord = "send_to_word";
        public const string OpenInChrome = "open_in_chrome";

        // The pilot keeps the existing handoff name. The model supplies
        // presentation wording and fact IDs; the host compiles all numeric
        // values, citations and workbook formulas from retained evidence.
        public static ChatToolDefinition AnalysisDeckDefinition()
        {
            var textPart = ToolSchema.Build(new Dictionary<string, object>
            {
                { "Text", ToolSchema.String("Prose only; no numeric claim.") },
                { "FactId", ToolSchema.String("Exact host-issued fact ID when showing a value.") },
                { "IncludeUnit", new Dictionary<string, object> { { "type", "boolean" } } }
            });
            var textParts = Array(textPart, 24);
            var cell = ToolSchema.Build(new Dictionary<string, object>
            {
                { "Text", ToolSchema.String("Label or prose without numeric claims.") },
                { "FactId", ToolSchema.String("Exact host-issued fact ID for this value.") }
            });
            var row = ToolSchema.Build(new Dictionary<string, object>
            {
                { "Cells", Array(cell, 30) }
            }, "Cells");
            var series = ToolSchema.Build(new Dictionary<string, object>
            {
                { "Name", ToolSchema.String("Metric label.") },
                { "FactIds", Array(ToolSchema.String("Exact fact ID in category order."), 30) }
            }, "Name", "FactIds");
            var chart = ToolSchema.Build(new Dictionary<string, object>
            {
                { "Type", ToolSchema.String("Native chart type supported by the draft writer.") },
                { "Title", ToolSchema.String("Concise chart title including units where requested.") },
                { "Categories", Array(ToolSchema.String("Category label."), 30) },
                { "Series", Array(series, 8) }
            }, "Type", "Title", "Categories", "Series");
            var card = ToolSchema.Build(new Dictionary<string, object>
            {
                { "Heading", ToolSchema.String("Card heading.") },
                { "Points", textParts }
            }, "Heading", "Points");
            var slide = ToolSchema.Build(new Dictionary<string, object>
            {
                { "Id", ToolSchema.String("Stable logical slide ID.") },
                { "Layout", ToolSchema.String("Supported Scribble native layout.") },
                { "Title", ToolSchema.String("Decision-oriented title without numeric literals.") },
                { "Subtitle", textParts },
                { "Takeaway", textParts },
                { "TableHeaders", Array(ToolSchema.String("Column heading."), 30) },
                { "TableRows", Array(row, 40) },
                { "Chart", chart },
                { "Cards", Array(card, 8) }
            }, "Id", "Layout", "Title");
            return new ChatToolDefinition
            {
                type = "function",
                function = new ChatToolFunctionDefinition
                {
                    name = SendToPowerPoint,
                    description = "Create a new unsaved, marked PowerPoint deck from the retained verified analysis. Supply only narrative/layout choices and host-issued fact IDs. Never retype numeric values, citations, formulas or source data. Use one call for the complete deck; native review may require a targeted correction.",
                    parameters = ToolSchema.Build(new Dictionary<string, object>
                    {
                        { "AnalysisId", ToolSchema.String("Exact analysis_id from the typed read_cells result.") },
                        { "WorkbookTitle", ToolSchema.String("Concise report title.") },
                        { "Slides", Array(slide, PresentationDraftWriter.MaxDraftSlides) }
                    }, "AnalysisId", "Slides")
                }
            };
        }

        private static Dictionary<string, object> Array(Dictionary<string, object> item,
            int maximum)
        {
            return new Dictionary<string, object>
            {
                { "type", "array" }, { "minItems", 1 },
                { "maxItems", maximum }, { "items", item }
            };
        }

        public static List<ChatToolDefinition> CreateDefinitions(
            string hostKind)
        {
            var definitions = new List<ChatToolDefinition>();
            // The Outlook pane has its own richer create_draft /
            // update_draft email tools; only the document panes get
            // the cross-app email tool.
            if (!string.Equals(
                hostKind,
                "outlook",
                StringComparison.Ordinal))
            {
                definitions.Add(new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = CreateEmailDraft,
                        description =
                            "Open one unsent Outlook email draft with the " +
                            "given recipients, subject, and body for the " +
                            "user to review. The draft is never sent and " +
                            "sending is impossible. When the current file " +
                            "is saved on disk it can be attached with " +
                            "attach_current_file. For a visual email use " +
                            "only these layout lines in body: # heading, " +
                            "## subheading, - list item, 1. numbered item, " +
                            "--- divider, and | cell | cell | table rows " +
                            "with a | --- | --- | separator under the " +
                            "header row.",
                        parameters = ToolSchema.Build(
                            new Dictionary<string, object>
                            {
                                {
                                    "to",
                                    ToolSchema.String(
                                        "Recipient addresses separated by " +
                                        "semicolons. May be empty for the " +
                                        "user to fill in.")
                                },
                                {
                                    "cc",
                                    ToolSchema.String(
                                        "Cc addresses separated by semicolons.")
                                },
                                {
                                    "subject",
                                    ToolSchema.String(
                                        "Email subject line.")
                                },
                                {
                                    "body",
                                    ToolSchema.String(
                                        "Plain-text email body using only the " +
                                        "allowed layout lines.")
                                },
                                {
                                    "attach_current_file",
                                    new Dictionary<string, object>
                                    {
                                        { "type", "boolean" },
                                        {
                                            "description",
                                            "Attach the current document file " +
                                            "when it is saved on disk."
                                        }
                                    }
                                }
                            },
                            "body")
                    }
                });
            }

            if (!string.Equals(
                hostKind,
                "powerpoint",
                StringComparison.Ordinal))
            {
                definitions.Add(new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = SendToPowerPoint,
                        description =
                            "Create a brand-new unsaved PowerPoint draft " +
                            "presentation with clearly marked [Scribble " +
                            "draft] slides. PowerPoint starts if needed, " +
                            "a fresh deck opens for this task and later batches continue it, existing " +
                            "files are never touched, and nothing is " +
                            "saved. Use it for any request to create a " +
                            "powerpoint, deck, or slides. At most " +
                            PresentationDraftWriter.MaxDraftSlides +
                            " slides per call. " + SamsungPresentationReview.AuthoringInstructions,
                        parameters =
                            PresentationToolCatalog.DraftDefinition()
                                .function.parameters as
                                Dictionary<string, object>
                    }
                });
            }

            if (!string.Equals(
                hostKind,
                "excel",
                StringComparison.Ordinal))
            {
                definitions.Add(new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = SendToExcel,
                        description =
                            "Create a brand-new unsaved Excel draft " +
                            "workbook and write the table (with live " +
                            "formulas and an optional native chart) into " +
                            "its 'Scribble Draft' worksheet. Excel starts if " +
                            "needed, a fresh workbook opens every time, " +
                            "existing files are never touched, and " +
                            "nothing is saved. Use it for any request to " +
                            "create an excel, spreadsheet, or workbook.",
                        parameters =
                            WorkbookToolCatalog.DraftDefinition()
                                .function.parameters as
                                Dictionary<string, object>
                    }
                });
            }

            if (!string.Equals(
                hostKind,
                "word",
                StringComparison.Ordinal))
            {
                definitions.Add(new ChatToolDefinition
                {
                    type = "function",
                    function = new ChatToolFunctionDefinition
                    {
                        name = SendToWord,
                        description =
                            "Create a brand-new unsaved Word draft " +
                            "document containing the given text " +
                            "for the user to review. Word starts " +
                            "if needed, a fresh document opens " +
                            "every time with an [Scribble draft] " +
                            "heading, existing documents are never " +
                            "modified, and nothing is saved. Use " +
                            "it for any request to create a word " +
                            "file or document.",
                        parameters =
                            WordToolCatalog.DraftDefinition()
                                .function.parameters as
                                Dictionary<string, object>
                    }
                });
            }

            if (hostKind != "chrome") definitions.Add(new ChatToolDefinition {
                type = "function", function = new ChatToolFunctionDefinition {
                    name = OpenInChrome,
                    description = "Open an HTTP or HTTPS webpage explicitly requested by the user in a new Chrome window. Chrome starts if needed. Supply the exact URL from the user's request; this does not upload the Office document or execute scripts.",
                    parameters = ToolSchema.Build(new Dictionary<string, object> { { "url", ToolSchema.String("Exact HTTP/HTTPS URL supplied in the user request.") } }, "url")
                }
            });
            return definitions;
        }

        public static bool IsCrossAppTool(string name)
        {
            return name == OpenInChrome || string.Equals(
                       name,
                       CreateEmailDraft,
                       StringComparison.Ordinal) ||
                   string.Equals(
                       name,
                       SendToPowerPoint,
                       StringComparison.Ordinal) ||
                   string.Equals(
                       name,
                       SendToExcel,
                       StringComparison.Ordinal) ||
                   string.Equals(
                       name,
                       SendToWord,
                       StringComparison.Ordinal);
        }
    }
}
