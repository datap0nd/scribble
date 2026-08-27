using System;
using System.Collections.Generic;

namespace Scribble.Security
{
    // Cross-checks model-chosen email recipients against the user's
    // own latest prompt. The draft still opens either way - sending
    // stays impossible and the user reviews it - but a recipient the
    // user never mentioned gets called out explicitly, so a model
    // that hallucinated or swapped an address is caught before the
    // user hits Send in Outlook.
    public static class RecipientIntentCheck
    {
        private const int MaxReportedRecipients = 5;

        // Returns an empty string when every recipient is grounded
        // in the prompt; otherwise a single warning sentence naming
        // the unmentioned addresses.
        public static string Warn(
            string to,
            string cc,
            string userPrompt)
        {
            var prompt = TextBoundary.PlainText(
                    userPrompt,
                    TextBoundary.MaxUserPromptCharacters)
                .ToLowerInvariant();
            if (prompt.Length == 0)
            {
                return string.Empty;
            }

            var unmentioned = new List<string>();
            foreach (var recipient in SplitRecipients(to, cc))
            {
                if (unmentioned.Count == MaxReportedRecipients)
                {
                    break;
                }

                if (!IsMentioned(prompt, recipient))
                {
                    unmentioned.Add(recipient);
                }
            }

            if (unmentioned.Count == 0)
            {
                return string.Empty;
            }

            return " Check the recipients before sending: " +
                string.Join(", ", unmentioned) +
                (unmentioned.Count == 1
                    ? " was"
                    : " were") +
                " not mentioned in your request.";
        }

        private static IEnumerable<string> SplitRecipients(
            string to,
            string cc)
        {
            var seen = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            var combined = (to ?? string.Empty) + ";" +
                (cc ?? string.Empty);
            foreach (var part in combined.Split(';', ','))
            {
                var recipient = TextBoundary.SingleLine(
                    part,
                    120);
                if (recipient.Length > 0 &&
                    seen.Add(recipient))
                {
                    yield return recipient;
                }
            }
        }

        // A recipient counts as mentioned when the prompt contains
        // the full address, its local part, or any local-part token
        // of three or more characters ("john.smith@x.com" matches a
        // prompt saying "email this to john").
        private static bool IsMentioned(
            string prompt,
            string recipient)
        {
            var address = recipient.ToLowerInvariant();
            if (prompt.IndexOf(
                    address,
                    StringComparison.Ordinal) >= 0)
            {
                return true;
            }

            var at = address.IndexOf('@');
            var local = at > 0
                ? address.Substring(0, at)
                : address;
            if (local.Length >= 3 &&
                prompt.IndexOf(
                    local,
                    StringComparison.Ordinal) >= 0)
            {
                return true;
            }

            foreach (var token in local.Split(
                '.',
                '_',
                '-'))
            {
                if (token.Length >= 3 &&
                    prompt.IndexOf(
                        token,
                        StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
