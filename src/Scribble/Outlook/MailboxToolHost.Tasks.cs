using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Scribble.Chat;

namespace Scribble.Outlook
{
    public sealed class MailboxTaskMessage
    {
        public string Handle { get; set; }
        public SavedMessage Source { get; set; }
        public string BodyEvidence { get; set; }
        public int BodyLength { get; set; } = -1;
        public int ReadUntil { get; set; }
        public int AttachmentCount { get; set; } = -1;
        public Dictionary<string, int> AttachmentOffsets { get; set; } = new Dictionary<string, int>();
        public Dictionary<string, string> AttachmentHashes { get; set; } = new Dictionary<string, string>();
        public List<int> CompleteAttachments { get; set; } = new List<int>();
        public bool Analysed { get; set; }
        public string Summary { get; set; }
        public string Id { get { return "mail:" + Source.StoreId + ":" + Source.EntryId; } }
    }
    public sealed class MailboxTaskSearch
    {
        public string Id { get; set; }
        public string Query { get; set; }
        public string Folder { get; set; }
        public DateTime After { get; set; }
        public DateTime Before { get; set; }
        public bool Unread { get; set; }
        public bool Complete { get; set; }
        public string StoreIdentity { get; set; }
    }
    public sealed class MailboxTaskCheckpoint
    {
        public List<MailboxTaskMessage> Messages { get; set; } = new List<MailboxTaskMessage>();
        public List<MailboxTaskSearch> Searches { get; set; } = new List<MailboxTaskSearch>();
    }

    public sealed partial class MailboxToolHost
    {
        private TaskContextManager _task;
        private MailboxTaskCheckpoint _ledger;
        private bool _reviewAll;
        private bool _requiresEnumeration;

        public async Task BindTaskAsync(TaskContextManager task, CancellationToken token)
        {
            _task = task;
            _reviewAll = Regex.IsMatch(task.State.Objective ?? "", @"\b(all|every|entire|unread|morning)\b", RegexOptions.IgnoreCase);
            _requiresEnumeration = !_workingSetOnly && _reviewAll &&
                !Regex.IsMatch(task.State.Objective ?? "", @"\b(selected|this email|this message)\b", RegexOptions.IgnoreCase);
            string saved;
            _ledger = task.State.HostData.TryGetValue("mailbox_ledger", out saved)
                ? _serializer.Deserialize<MailboxTaskCheckpoint>(saved) : new MailboxTaskCheckpoint();
            foreach (var entry in _ledger.Messages)
            {
                token.ThrowIfCancellationRequested();
                var current = new MessageReader(_application).CaptureById(entry.Source.EntryId, entry.Source.StoreId);
                if (entry.BodyEvidence != null && TaskCheckpointStore.Fingerprint(current.Body) != entry.BodyEvidence)
                    throw new InvalidOperationException("Message content changed since checkpoint: " + entry.Source.Subject);
                _handles[entry.Handle] = current;
                foreach (var attachment in entry.AttachmentHashes)
                {
                    var page = await MailboxAttachmentPages.ReadAsync(_application, current, int.Parse(attachment.Key), 0, token);
                    if (page.Fingerprint != attachment.Value) throw new InvalidOperationException("An attachment changed since checkpoint: " + page.FileName);
                }
                await Task.Yield();
            }
            foreach (var search in _ledger.Searches.Where(s => !s.Complete))
            {
                if (StoreIdentity(search.Folder) != search.StoreIdentity) throw new InvalidOperationException("The Outlook mailbox/profile changed. Reopen the original mailbox to resume.");
                _cursors[search.Id] = new MailboxPageCursor(_application, search.Query, search.Folder, search.After, search.Before, search.Unread);
            }
            _nextHandle = _handles.Count + 1;
            foreach (var handle in _handles.Keys.ToArray()) { _metadataHandles.Add(handle); if (_workingSetOnly) RegisterCoverage(handle, _handles[handle]); }
            SaveCoverage();
        }

        private string StoreIdentity(string folder)
        {
            dynamic app = _application;
            dynamic session = app.Session;
            object inbox = null, sent = null;
            try
            {
                var identity = "";
                if (folder != "sent") { inbox = session.GetDefaultFolder(6); dynamic item = inbox; identity += Convert.ToString(item.StoreID); }
                if (folder != "inbox") { sent = session.GetDefaultFolder(5); dynamic item = sent; identity += ":" + Convert.ToString(item.StoreID); }
                return identity;
            }
            finally
            {
                ReleaseTaskCom(inbox); ReleaseTaskCom(sent); ReleaseTaskCom((object)session);
            }
        }
        private static void ReleaseTaskCom(object value)
        {
            if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value)) System.Runtime.InteropServices.Marshal.ReleaseComObject(value);
        }

        private MailboxTaskMessage RegisterCoverage(string handle, MessageSnapshot source)
        {
            if (_ledger == null) return null;
            var entry = _ledger.Messages.FirstOrDefault(m => m.Source.EntryId == source.EntryId && m.Source.StoreId == source.StoreId);
            if (entry == null)
            {
                var saved = TaskRecoveryInput.Copy<SavedMessage>(source);
                saved.Body = "";
                entry = new MailboxTaskMessage { Handle = handle, Source = saved };
                _ledger.Messages.Add(entry);
                if (!_task.State.ExpectedSourceIds.Contains(entry.Id)) _task.State.ExpectedSourceIds.Add(entry.Id);
                _metadataHandles.Add(handle);
            }
            return entry;
        }

        public string CompletionBlocker
        {
            get
            {
                if (_task == null) return null;
                if (_requiresEnumeration && _ledger.Searches.Count == 0)
                    return "Mailbox coverage is incomplete: enumerate the user's requested mailbox/time window with search_mailbox before answering.";
                var cursor = _ledger.Searches.FirstOrDefault(s => !s.Complete);
                if (cursor != null) return "Mailbox enumeration is incomplete. Continue search_mailbox with cursor " + cursor.Id + " (empty pages do not mean completion).";
                var pending = _ledger.Messages.Where(m => !m.Analysed && !_task.State.Exclusions.ContainsKey(m.Id)).ToArray();
                if (pending.Length > 0) return pending.Length + " messages still need analysis receipts. Read every body/attachment page, then call record_mailbox_analysis. Next handles: " + string.Join(", ", pending.Take(10).Select(m => m.Handle));
                return null;
            }
        }

        public IReadOnlyList<MailboxTaskMessage> AnalysisReport { get { return _ledger?.Messages.Where(m => m.Analysed).ToArray() ?? new MailboxTaskMessage[0]; } }

        private void SaveCoverage()
        {
            if (_task == null) return;
            _task.State.HostData["mailbox_ledger"] = _serializer.Serialize(_ledger);
            _task.Checkpoint();
        }

        private MailboxToolResult RecordAnalysis(string callId, IDictionary<string, object> arguments)
        {
            if (_task == null) return Error(callId, "MAILBOX_TASK_REQUIRED", "Analysis receipts require an active task.");
            var handle = GetString(arguments, "handle", "");
            var entry = _ledger.Messages.FirstOrDefault(m => m.Handle == handle);
            if (entry == null) return Error(callId, "MAILBOX_HANDLE_UNKNOWN", "Unknown message handle.");
            var exclusion = GetString(arguments, "exclusion_reason", "");
            if (exclusion.Length > 0)
            {
                if (_reviewAll) return Error(callId, "MAILBOX_REVIEW_ALL", "This request requires every matching message; relevance exclusions are not allowed.");
                _task.State.Exclusions[entry.Id] = exclusion;
            }
            else
            {
                var current = new MessageReader(_application).CaptureById(entry.Source.EntryId, entry.Source.StoreId);
                if ((entry.BodyEvidence != null && TaskCheckpointStore.Fingerprint(current.Body) != entry.BodyEvidence) ||
                    (entry.AttachmentCount >= 0 && MailboxAttachmentPages.Count(_application, current) != entry.AttachmentCount))
                    return Error(callId, "MAILBOX_SOURCE_CHANGED", "The message body or attachment collection changed during analysis. Its old coverage cannot be used.");
                if (entry.BodyLength < 0 || entry.ReadUntil < entry.BodyLength || entry.AttachmentCount < 0 ||
                    entry.CompleteAttachments.Count != entry.AttachmentCount)
                    return Error(callId, "MAILBOX_COVERAGE_INCOMPLETE", "Read the complete body and every attachment before recording analysis. Body offset " + entry.ReadUntil + "; attachments complete " + entry.CompleteAttachments.Count + " of " + entry.AttachmentCount + ".");
                var summary = GetString(arguments, "summary", "");
                if (summary.Length == 0) return Error(callId, "MAILBOX_ANALYSIS_REQUIRED", "Provide a source-grounded summary, including actions and important evidence.");
                entry.Analysed = true;
                entry.Summary = summary;
                _task.State.Exclusions.Remove(entry.Id);
                if (!_task.State.Batches.Any(b => b.Id == entry.Id)) _task.State.Batches.Add(new TaskBatchResult
                {
                    Id = entry.Id, CoveredSourceIds = new List<string> { entry.Id }, Output = summary,
                    EvidenceReferences = new List<string> { entry.BodyEvidence }
                });
            }
            SaveCoverage();
            return Success(callId, new { ok = true, source_id = entry.Id, analysed = entry.Analysed,
                analysed_count = _ledger.Messages.Count(m => m.Analysed), matched_count = _ledger.Messages.Count,
                outstanding = CompletionBlocker }, "Mailbox analysis coverage saved");
        }

        private async Task<MailboxToolResult> ReadAttachmentAsync(string callId, IDictionary<string, object> arguments, CancellationToken token)
        {
            var handle = GetString(arguments, "handle", "");
            MessageSnapshot source;
            if (!_handles.TryGetValue(handle, out source)) return Error(callId, "MAILBOX_HANDLE_UNKNOWN", "Unknown message handle.");
            var index = GetInteger(arguments, "attachment_index", 1, 1, int.MaxValue);
            var offset = GetInteger(arguments, "offset", 0, 0, int.MaxValue);
            var page = await MailboxAttachmentPages.ReadAsync(_application, source, index, offset, token);
            var entry = RegisterCoverage(handle, source);
            if (entry != null)
            {
                var key = index.ToString();
                string fingerprint;
                if (entry.AttachmentHashes.TryGetValue(key, out fingerprint) && fingerprint != page.Fingerprint)
                    return Error(callId, "ATTACHMENT_CHANGED", "The attachment changed during processing. Its old analysis was not reused.");
                entry.AttachmentHashes[key] = page.Fingerprint;
                int readUntil;
                entry.AttachmentOffsets.TryGetValue(key, out readUntil);
                if (offset > readUntil) return Error(callId, "ATTACHMENT_PAGE_GAP", "Continue this attachment at offset " + readUntil);
                entry.AttachmentOffsets[key] = Math.Max(readUntil, offset + page.Text.Length);
                if (!page.NextOffset.HasValue && !entry.CompleteAttachments.Contains(index)) entry.CompleteAttachments.Add(index);
                SaveCoverage();
            }
            var images = string.IsNullOrEmpty(page.ImageDataUrl) ? new VisionImagePayload[0] : new[] { new VisionImagePayload(page.FileName, page.ImageDataUrl) };
            return Success(callId, new { untrusted_attachment_data = true, handle, attachment_index = index,
                file_name = page.FileName, offset, next_offset = page.NextOffset, complete = !page.NextOffset.HasValue,
                content = page.Text, kind = page.Kind }, "Read attachment " + index + ": " + page.FileName, images);
        }
    }
}
