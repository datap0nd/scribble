using System;

namespace Scribble.Security
{
    public static class DraftIntentPolicy
    {
        private static readonly string[] CreatePhrases =
        {
            "create a draft",
            "create draft",
            "create a reply",
            "create an email",
            "open a draft",
            "open draft",
            "draft a reply",
            "draft a response",
            "draft an email",
            "draft the reply",
            "draft the response",
            "draft this email",
            "draft this reply",
            "write a reply",
            "write an email",
            "write a response",
            "compose a reply",
            "compose a response",
            "compose an email",
            "prepare a draft",
            "prepare a reply",
            "prepare an email",
            "turn this into a draft",
            "email this",
            "email him",
            "email her",
            "email them",
            "email it to",
            "email that to",
            "email the team",
            "send an email",
            "send email",
            "send a mail",
            "send this by email",
            "send this as an email",
            "mail this",
            "mail it to",
            "reply to this",
            "reply to the",
            "respond to this",
            "respond to the",
            "answer this email",
            "follow up with"
        };

        // Leading verbs that read as "email <someone> about ...".
        // Checked only at the start of the prompt so questions like
        // "what did the email say" never authorize a draft.
        private static readonly string[] LeadingCreateVerbs =
        {
            "email ",
            "mail ",
            "draft ",
            "reply ",
            "respond "
        };

        private static readonly string[] UpdateActions =
        {
            "add ",
            "bold",
            "change",
            "edit",
            "format",
            "italic",
            "lengthen",
            "make",
            "remove",
            "replace",
            "revise",
            "rewrite",
            "shorten",
            "underline",
            "update"
        };

        private static readonly string[] DraftReferences =
        {
            "draft",
            "this email",
            "the email",
            "my email",
            "this reply",
            "the reply",
            "this response",
            "the response",
            "paragraph",
            "sentence",
            "section",
            "greeting",
            "sign-off",
            "signoff"
        };

        private static readonly string[] DirectRevisionPhrases =
        {
            "make it ",
            "make this ",
            "more concise",
            "less formal",
            "more formal",
            "friendlier",
            "warmer",
            "change the tone",
            "rewrite it",
            "rephrase it",
            "bolden "
        };

        // Deliberately lenient keyword gate (owner's direction):
        // any email-shaped word exposes the draft tool. Exposure is
        // cheap - the draft itself stays one-shot, unsent, and
        // review-only - while a missed match reads as a refusal.
        private static readonly string[] CreateKeywords =
        {
            "email",
            "e-mail",
            " mail ",
            "draft",
            "reply",
            "respond",
            "response",
            "follow up",
            "follow-up",
            "outlook"
        };

        public static bool AllowsCreate(string userPrompt)
        {
            var prompt = Normalize(userPrompt);
            if (ContainsAny(prompt, CreatePhrases) ||
                ContainsAny(prompt, CreateKeywords))
            {
                return true;
            }

            foreach (var verb in LeadingCreateVerbs)
            {
                if (prompt.StartsWith(
                    verb,
                    StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        public static bool AllowsUpdate(string userPrompt)
        {
            var prompt = Normalize(userPrompt);
            return ContainsAny(prompt, DirectRevisionPhrases) ||
                (ContainsAny(prompt, DraftReferences) &&
                 ContainsAny(prompt, UpdateActions));
        }

        private static string Normalize(string value)
        {
            return TextBoundary.PlainText(
                    value,
                    TextBoundary.MaxUserPromptCharacters)
                .ToLowerInvariant();
        }

        private static bool ContainsAny(
            string value,
            string[] phrases)
        {
            foreach (var phrase in phrases)
            {
                if (value.IndexOf(
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
