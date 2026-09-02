using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Scribble.Security;

namespace Scribble.Outlook
{
    internal sealed class MailboxContextService
    {
        private const int InboxFolder = 6;
        private const int SentItemsFolder = 5;
        private const int MaxFallbackScannedItemsPerFolder = 600;

        private readonly object _outlookApplication;

        public MailboxContextService(object outlookApplication)
        {
            _outlookApplication = outlookApplication ??
                throw new ArgumentNullException(nameof(outlookApplication));
        }

        public IReadOnlyList<MailboxSearchHit> Search(
            string query,
            string folderScope,
            int daysBack,
            int maxResults,
            DateTime? receivedAfter = null,
            DateTime? receivedBefore = null,
            bool unreadOnly = false)
        {
            var boundedQuery = TextBoundary.PlainText(query, 240);
            var boundedDays = Math.Max(1, Math.Min(3650, daysBack));
            var boundedResults = Math.Max(
                1,
                Math.Min(
                    MailboxWorkingSet.MaxMessages + 1,
                    maxResults));
            var cutoff = receivedAfter ??
                DateTime.Now.AddDays(-boundedDays);
            var exactTimeRange = receivedAfter.HasValue ||
                receivedBefore.HasValue;
            var hits = new List<MailboxSearchHit>();

            if (folderScope == "inbox" || folderScope == "all")
            {
                SearchDefaultFolder(
                    InboxFolder,
                    "Inbox",
                    boundedQuery,
                    cutoff,
                    receivedBefore,
                    exactTimeRange,
                    unreadOnly,
                    boundedResults,
                    hits);
            }

            if (folderScope == "sent" || folderScope == "all")
            {
                SearchDefaultFolder(
                    SentItemsFolder,
                    "Sent Items",
                    boundedQuery,
                    cutoff,
                    receivedBefore,
                    exactTimeRange,
                    unreadOnly,
                    boundedResults,
                    hits);
            }

            return hits
                .OrderByDescending(hit =>
                    hit.Message.ReceivedAt ?? DateTime.MinValue)
                .Take(boundedResults)
                .ToArray();
        }

        public IReadOnlyList<EmailAttachmentContent> ReadAttachments(
            MessageSnapshot message)
        {
            return EmailAttachmentReader.Read(
                _outlookApplication,
                message);
        }

        public IReadOnlyList<MessageSnapshot> ReadConversation(
            MessageSnapshot source,
            int maxMessages)
        {
            if (source == null || source.EntryId.Length == 0)
            {
                throw new InvalidOperationException(
                    "The source message is unavailable.");
            }

            var limit = Math.Max(
                1,
                Math.Min(
                    Math.Max(20, MailboxWorkingSet.MaxMessages),
                    maxMessages));
            var messages = new List<MessageSnapshot>();
            object session = null;
            object sourceItem = null;
            object conversation = null;
            object table = null;

            try
            {
                dynamic application = _outlookApplication;
                session = application.Session;
                dynamic outlookSession = session;
                sourceItem = source.StoreId.Length > 0
                    ? outlookSession.GetItemFromID(
                        source.EntryId,
                        source.StoreId)
                    : outlookSession.GetItemFromID(source.EntryId);
                dynamic mail = sourceItem;
                conversation = mail.GetConversation();
                if (conversation == null)
                {
                    return new[] { source };
                }

                dynamic outlookConversation = conversation;
                table = outlookConversation.GetTable();
                dynamic conversationTable = table;
                while (!conversationTable.EndOfTable &&
                       messages.Count < limit)
                {
                    object row = null;
                    object item = null;
                    try
                    {
                        row = conversationTable.GetNextRow();
                        dynamic conversationRow = row;
                        var entryId =
                            Convert.ToString(
                                conversationRow["EntryID"]) ??
                            string.Empty;
                        if (entryId.Length == 0)
                        {
                            continue;
                        }

                        try
                        {
                            item = source.StoreId.Length > 0
                                ? outlookSession.GetItemFromID(
                                    entryId,
                                    source.StoreId)
                                : outlookSession.GetItemFromID(entryId);
                        }
                        catch
                        {
                            item = outlookSession.GetItemFromID(entryId);
                        }

                        if (MessageReader.IsReadableItem(item))
                        {
                            messages.Add(
                                MessageReader.CaptureItem(item));
                        }
                    }
                    finally
                    {
                        Release(item);
                        Release(row);
                    }
                }
            }
            finally
            {
                Release(table);
                Release(conversation);
                Release(sourceItem);
                Release(session);
            }

            if (messages.Count == 0)
            {
                messages.Add(source);
            }

            return messages
                .OrderBy(message =>
                    message.ReceivedAt ?? DateTime.MinValue)
                .Take(limit)
                .ToArray();
        }

        private void SearchDefaultFolder(
            int folderKind,
            string folderName,
            string query,
            DateTime cutoff,
            DateTime? receivedBefore,
            bool exactTimeRange,
            bool unreadOnly,
            int maxResults,
            List<MailboxSearchHit> output)
        {
            object session = null;
            object folder = null;
            object items = null;
            object searchItems = null;
            var folderHitCount = 0;

            try
            {
                dynamic application = _outlookApplication;
                session = application.Session;
                dynamic outlookSession = session;
                folder = outlookSession.GetDefaultFolder(folderKind);
                dynamic outlookFolder = folder;
                items = outlookFolder.Items;

                if (query.Length > 0 || unreadOnly)
                {
                    try
                    {
                        dynamic outlookItems = items;
                        searchItems = query.Length > 0
                            ? outlookItems.Restrict(
                                BuildDaslFilter(query))
                            : outlookItems.Restrict(
                                "[UnRead] = true");
                    }
                    catch
                    {
                        searchItems = null;
                    }
                }

                var collection = searchItems ?? items;
                dynamic candidates = collection;
                var sortedNewest = false;
                try
                {
                    candidates.Sort(
                        folderKind == SentItemsFolder
                            ? "[SentOn]"
                            : "[ReceivedTime]",
                        true);
                    sortedNewest = true;
                }
                catch
                {
                }

                var count = Convert.ToInt32(candidates.Count);
                var scanLimit = searchItems != null
                    ? count
                    : Math.Min(
                        count,
                        MaxFallbackScannedItemsPerFolder);

                for (var index = 1;
                     index <= scanLimit &&
                     folderHitCount < maxResults;
                     index++)
                {
                    object item = null;
                    try
                    {
                        item = candidates.Item(index);
                        if (!MessageReader.IsReadableItem(item))
                        {
                            continue;
                        }

                        var snapshot = MessageReader.CaptureItem(item);
                        if (exactTimeRange &&
                            !snapshot.ReceivedAt.HasValue)
                        {
                            continue;
                        }

                        if (receivedBefore.HasValue &&
                            snapshot.ReceivedAt.HasValue &&
                            snapshot.ReceivedAt.Value >
                                receivedBefore.Value)
                        {
                            continue;
                        }

                        if (snapshot.ReceivedAt.HasValue &&
                            snapshot.ReceivedAt.Value < cutoff)
                        {
                            if (sortedNewest)
                            {
                                break;
                            }

                            continue;
                        }

                        if (unreadOnly && !snapshot.IsUnread)
                        {
                            continue;
                        }

                        if (searchItems == null &&
                            query.Length > 0 &&
                            !Matches(snapshot, query))
                        {
                            continue;
                        }

                        output.Add(new MailboxSearchHit(
                            snapshot,
                            folderName,
                            BuildSnippet(snapshot.Body, query)));
                        folderHitCount++;
                    }
                    finally
                    {
                        Release(item);
                    }
                }
            }
            finally
            {
                if (!ReferenceEquals(searchItems, items))
                {
                    Release(searchItems);
                }

                Release(items);
                Release(folder);
                Release(session);
            }
        }

        private static string BuildDaslFilter(string query)
        {
            var escaped = query.Replace("'", "''");
            return
                "@SQL=(" +
                "\"urn:schemas:httpmail:subject\" ci_phrasematch '" +
                escaped +
                "') OR (" +
                "\"urn:schemas:httpmail:textdescription\" ci_phrasematch '" +
                escaped +
                "') OR (" +
                "\"urn:schemas:httpmail:fromname\" ci_phrasematch '" +
                escaped +
                "') OR (" +
                "\"urn:schemas:httpmail:displayto\" ci_phrasematch '" +
                escaped +
                "')";
        }

        private static bool Matches(
            MessageSnapshot message,
            string query)
        {
            return Contains(message.Subject, query) ||
                   Contains(message.Sender, query) ||
                   Contains(message.Recipients, query) ||
                   Contains(message.Body, query);
        }

        private static bool Contains(string value, string query)
        {
            return (value ?? string.Empty).IndexOf(
                query,
                StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string BuildSnippet(string body, string query)
        {
            var plain = TextBoundary.PlainText(body, 4000)
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("\t", " ");
            while (plain.Contains("  "))
            {
                plain = plain.Replace("  ", " ");
            }

            var start = 0;
            if (query.Length > 0)
            {
                var match = plain.IndexOf(
                    query,
                    StringComparison.OrdinalIgnoreCase);
                if (match > 120)
                {
                    start = match - 120;
                }
            }

            if (start >= plain.Length)
            {
                return string.Empty;
            }

            var length = Math.Min(500, plain.Length - start);
            var snippet = plain.Substring(start, length).Trim();
            return start > 0 ? "..." + snippet : snippet;
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }
    }

    internal sealed class MailboxSearchHit
    {
        public MailboxSearchHit(
            MessageSnapshot message,
            string folderName,
            string snippet)
        {
            Message = message ??
                throw new ArgumentNullException(nameof(message));
            FolderName = folderName ?? string.Empty;
            Snippet = snippet ?? string.Empty;
        }

        public MessageSnapshot Message { get; }

        public string FolderName { get; }

        public string Snippet { get; }
    }
}
