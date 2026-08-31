using System;

namespace Scribble.Outlook
{
    // A mailbox sweep is a read-only metadata scan, so it is
    // budgeted by the text it returns rather than by the
    // approval-bound working set in MailboxWorkingSet. The user
    // decides how wide a sweep is - up to MaxResults summaries in
    // one pass - and the packing helpers here shrink each summary
    // so that a wide sweep still fits the request's tool-result
    // budget. Message bodies stay on the narrower read budget
    // below: bodies are the expensive intake, summaries are not.
    public static class MailboxSearchBudget
    {
        public const int MaxResults = 500;
        public const int RecommendedResults = 25;
        public const int MaxSearchesPerRequest = 4;
        public const int MaxBodyMessages = 25;
        public const int MaxThreadMessages = 20;
        public const int MaxSnippetCharacters = 500;

        // Up to this many results a sweep carries full summaries.
        // Beyond it the summaries go compact - shorter subject, no
        // recipient list, a trimmed or dropped snippet - because a
        // wide sweep is a list of what is there, not a quotation of
        // it. read_messages loads the bodies worth reading.
        private const int DetailedResults = 50;
        private const int CompactResults = 150;

        public static int SnippetCharacters(int results)
        {
            if (results <= DetailedResults)
            {
                return MaxSnippetCharacters;
            }

            return results <= CompactResults ? 200 : 0;
        }

        public static int SubjectCharacters(int results)
        {
            return results <= DetailedResults ? 240 : 140;
        }

        public static int SenderCharacters(int results)
        {
            return results <= DetailedResults ? 240 : 120;
        }

        public static bool IncludesRecipients(int results)
        {
            return results <= DetailedResults;
        }

        // Outlook's DASL restriction filters server-side when a
        // query is present; the manual fallback scan has to widen
        // with the sweep or a 500-result request quietly returns a
        // handful of matches.
        public static int ScannedItemsPerFolder(int maxResults)
        {
            var scaled = Math.Max(1, maxResults) * 6;
            return Math.Max(600, Math.Min(6000, scaled));
        }

        // Per-message body budget for one read_messages call. Ten
        // handles or fewer keep the full reviewed body length; a
        // larger read shares the request's tool-result budget out
        // so the result is still returnable.
        public static int BodyCharacters(
            int messageCount,
            int resultBudget,
            int maximum)
        {
            var share = resultBudget /
                Math.Max(1, messageCount) -
                400;
            return Math.Max(
                1200,
                Math.Min(maximum, share));
        }
    }
}
