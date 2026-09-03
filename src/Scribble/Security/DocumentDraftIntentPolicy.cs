using System;

namespace Scribble.Security
{
    // Local intent gate for the Excel and PowerPoint panes. Document
    // draft surfaces (the Scribble Draft worksheet, appended draft
    // slides, and unsent Outlook email drafts) unlock only when the
    // user's own latest prompt contains an explicit produce-something
    // phrase; model or document text can never grant the permission.
    public static class DocumentDraftIntentPolicy
    {
        private static readonly string[] DraftPhrases =
        {
            "create a draft",
            "create draft",
            "draft sheet",
            "draft slide",
            "draft a slide",
            "draft an email",
            "draft the email",
            "draft this",
            "create a slide",
            "create slides",
            "create a sheet",
            "create a table in",
            "create an email",
            "add a slide",
            "add slides",
            "add a sheet",
            "insert a slide",
            "insert slides",
            "insert a sheet",
            "insert a table",
            "insert into",
            "make a slide",
            "make slides",
            "make a sheet",
            "make a table in",
            "write it to",
            "write this to",
            "write these to",
            "write to the draft",
            "write an email",
            "compose an email",
            "put this in",
            "put it in",
            "put these in",
            "put that in",
            "export to",
            "email this",
            "email these",
            "email it",
            "email that",
            "email the",
            "mail this",
            "mail these",
            "send this to",
            "send these to",
            "send it to",
            "send that to",
            "send to outlook",
            "send to excel",
            "send to powerpoint",
            "prepare a draft",
            "prepare an email",
            "prepare a slide",
            "turn this into",
            "turn it into",
            "turn that into",
            "build a slide",
            "build slides",
            "build a deck",
            "build a table",
            "build a chart",
            "build a report",
            "make a chart",
            "make a graph",
            "do a chart",
            "bar chart",
            "line chart",
            "pie chart",
            "column chart",
            "a chart of",
            "a chart with",
            "a chart from",
            "a graph of",
            "a graph with",
            "chart this",
            "graph this",
            "plot this",
            "visualize",
            "put this",
            "put that",
            "put these",
            "put it",
            "put the",
            "into word",
            "into excel",
            "into powerpoint",
            "into outlook",
            "into an email",
            "into a slide",
            "into a sheet",
            "into a document",
            "in a slide",
            "in a sheet",
            "in a draft",
            "in a document",
            "in a new document",
            "to word",
            "to a slide",
            "as a slide",
            "as a table",
            "as a document",
            "as an email",
            "send to word",
            "write a report",
            "write a summary",
            "write a document",
            "write the report",
            "create an excel",
            "create a excel",
            "create excel",
            "create a word",
            "create word",
            "create a powerpoint",
            "create powerpoint",
            "create a presentation",
            "create a document",
            "create a workbook",
            "create a spreadsheet",
            "create a deck",
            "create a file",
            "create a report",
            "make an excel",
            "make a word",
            "make a powerpoint",
            "make a presentation",
            "make a document",
            "make a workbook",
            "make a spreadsheet",
            "new excel",
            "new word file",
            "new word document",
            "new powerpoint",
            "new presentation",
            "new workbook",
            "new spreadsheet",
            "new document",
            "an excel file",
            "an excel sheet",
            "a word file",
            "a word doc",
            "a powerpoint file",
            "a powerpoint deck",
            "in excel",
            "in word",
            "in powerpoint"
        };

        // Editing verbs authorize a draft only together with a
        // document reference, mirroring the Outlook update policy:
        // "update the table" or "fix column B" unlocks the marked
        // draft surface, while "update me on the project" does not.
        private static readonly string[] EditActions =
        {
            "edit",
            "update",
            "change",
            "fix",
            "correct",
            "fill",
            "populate",
            "revise",
            "rewrite",
            "modify",
            "adjust",
            "calculate",
            "compute",
            "sort",
            "reformat",
            "format",
            "clean up",
            "improve",
            "translate",
            "shorten",
            "expand",
            "build",
            "generate",
            "produce",
            "convert",
            "draw",
            "plot",
            "chart",
            "graph",
            "highlight",
            "restructure",
            "reorganize",
            "create",
            "make",
            "write",
            "draft",
            "add",
            "insert",
            "prepare",
            "compose",
            "start",
            "open",
            "give me"
        };

        private static readonly string[] DocumentReferences =
        {
            "sheet",
            "cell",
            "cells",
            "column",
            "row",
            "rows",
            "table",
            "workbook",
            "spreadsheet",
            "range",
            "formula",
            "slide",
            "slides",
            "deck",
            "presentation",
            "document",
            "paragraph",
            "section",
            "text",
            "draft",
            "data",
            "numbers",
            "values",
            "chart",
            "graph",
            "report",
            "summary",
            "analysis",
            "email",
            "excel",
            "powerpoint",
            "power point",
            "word file",
            "word doc",
            "a word",
            "file",
            "xlsx",
            "docx",
            "pptx"
        };

        public static bool AllowsDraft(string userPrompt)
        {
            return AllowsDraft(userPrompt, false);
        }

        // A deliberately attached Excel selection is a local user
        // gesture, so it can satisfy the document-reference half of
        // the gate. An edit verb is still required: attaching cells
        // by itself never grants write permission, and document or
        // model text can never set this flag.
        public static bool AllowsDraft(
            string userPrompt,
            bool hasAttachedExcelSelection)
        {
            var prompt = TextBoundary.PlainText(
                    userPrompt,
                    TextBoundary.MaxUserPromptCharacters)
                .ToLowerInvariant();
            foreach (var phrase in DraftPhrases)
            {
                if (prompt.IndexOf(
                        phrase,
                        StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return ContainsAny(prompt, EditActions) &&
                   (hasAttachedExcelSelection ||
                    ContainsAny(prompt, DocumentReferences));
        }

        private static bool ContainsAny(
            string prompt,
            string[] phrases)
        {
            foreach (var phrase in phrases)
            {
                if (prompt.IndexOf(
                        phrase,
                        StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
