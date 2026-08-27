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
            "turn this into a draft"
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

        public static bool AllowsCreate(string userPrompt)
        {
            return ContainsAny(
                Normalize(userPrompt),
                CreatePhrases);
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
