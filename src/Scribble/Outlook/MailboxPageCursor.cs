using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace Scribble.Outlook
{
    // A table cursor does not reopen a sorted Items collection at a numerical
    // offset: arrivals and deletions cannot shift that offset between pages.
    internal sealed class MailboxPageCursor : IDisposable
    {
        private readonly object _application;
        private readonly string _query;
        private readonly DateTime _after;
        private readonly DateTime _before;
        private readonly bool _unread;
        private readonly Queue<int> _folders;
        private readonly HashSet<string> _seen = new HashSet<string>(StringComparer.Ordinal);
        private object _table;
        private string _storeId;
        private string _folderName;
        public bool Complete { get; private set; }

        public MailboxPageCursor(object application, string query, string folder,
            DateTime after, DateTime before, bool unread)
        {
            _application = application;
            _query = query ?? "";
            _after = after;
            _before = before;
            _unread = unread;
            _folders = new Queue<int>(folder == "inbox" ? new[] { 6 } :
                folder == "sent" ? new[] { 5 } : new[] { 6, 5 });
        }

        private bool OpenNextFolder()
        {
            Release(_table);
            _table = null;
            if (_folders.Count == 0) { Complete = true; return false; }
            object session = null;
            object folder = null;
            try
            {
                dynamic application = _application;
                session = application.Session;
                dynamic outlookSession = session;
                var kind = _folders.Peek();
                folder = outlookSession.GetDefaultFolder(kind);
                dynamic source = folder;
                _storeId = Convert.ToString(source.StoreID);
                _folderName = kind == 6 ? "Inbox" : "Sent Items";
                // Outlook evaluates the full-text predicate without materializing bodies.
                // If the provider cannot evaluate it, fail explicitly; never silently
                // fall back to a fixed number of recent items.
                _table = source.GetTable(_query.Length == 0 ? "" :
                    MailboxContextService.BuildDaslFilter(_query));
                _folders.Dequeue();
                return true;
            }
            finally { Release(folder); Release(session); }
        }

        public async Task<IReadOnlyList<MailboxSearchHit>> ReadAsync(int pageSize,
            CancellationToken cancellationToken)
        {
            var hits = new List<MailboxSearchHit>();
            var reader = new MessageReader(_application);
            // A page also bounds nonmatching rows, so sparse searches yield to the UI.
            var scanned = 0;
            var characters = 0;
            while (hits.Count < pageSize && scanned < 100 && characters < 12000)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (_table == null && !OpenNextFolder()) break;
                dynamic table = _table;
                if (table.EndOfTable)
                {
                    if (!OpenNextFolder()) break;
                    continue;
                }
                object row = null;
                try
                {
                    row = table.GetNextRow();
                    dynamic entry = row;
                    string id = Convert.ToString(entry["EntryID"]);
                    scanned++;
                    if (!MessageReader.IsReadableItemClass(Convert.ToString(entry["MessageClass"]))) continue;
                    if (!_seen.Add(_storeId + "\n" + id)) continue;
                    // Metadata only. Missing/unreadable items are errors, not coverage.
                    var message = reader.CaptureById(id, _storeId, true);
                    if (!message.ReceivedAt.HasValue || message.ReceivedAt.Value < _after ||
                        message.ReceivedAt.Value > _before || (_unread && !message.IsUnread)) continue;
                    hits.Add(new MailboxSearchHit(message, _folderName, ""));
                    characters += message.EntryId.Length + message.StoreId.Length +
                        message.Subject.Length + message.Sender.Length + message.Recipients.Length + 512;
                }
                finally
                {
                    Release(row);
                    await Task.Yield();
                }
            }
            if (_table != null && _folders.Count == 0)
            {
                dynamic table = _table;
                if (table.EndOfTable) { Complete = true; Dispose(); }
            }
            return hits;
        }

        public void Dispose() { Release(_table); _table = null; }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }
}
