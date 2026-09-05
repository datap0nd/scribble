using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using Scribble.Office;
using Scribble.Outlook;

namespace Scribble.Chat
{
    public sealed class TaskRecoveryInput
    {
        public static readonly string ProcessSession = Guid.NewGuid().ToString("N");
        public string Prompt { get; set; }
        public string SelectionHandle { get; set; }
        public SavedSelection Selection { get; set; }
        public bool ReplaceSource { get; set; }
        public string KoreanHandle { get; set; }
        public SavedKoreanWorkbook Korean { get; set; }
        public SavedMessage Selected { get; set; }
        public List<SavedMessage> Working { get; set; } = new List<SavedMessage>();
        public List<SavedReference> Documents { get; set; } = new List<SavedReference>();
        public List<SavedImage> Images { get; set; } = new List<SavedImage>();

        public static TaskRecoveryInput Read(DurableTaskState state)
        {
            return new JavaScriptSerializer { MaxJsonLength = int.MaxValue }
                .Deserialize<TaskRecoveryInput>(state.HostData["recovery_input"]);
        }

        public void PersistTo(DurableTaskState state)
        {
            state.HostData["recovery_input"] = new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.Serialize(this);
        }

        public static T Copy<T>(object value)
        {
            if (value == null) return default(T);
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            return json.Deserialize<T>(json.Serialize(value));
        }
    }

    public sealed class SavedReference
    {
        public string Name { get; set; }
        public string Content { get; set; }
    }
    public sealed class SavedImage
    {
        public string FileName { get; set; }
        public string DataUrl { get; set; }
    }
    public sealed class SavedMessage
    {
        public string EntryId { get; set; }
        public string StoreId { get; set; }
        public string Subject { get; set; }
        public string Sender { get; set; }
        public string Recipients { get; set; }
        public DateTime? ReceivedAt { get; set; }
        public string Body { get; set; }
        public List<string> AttachmentNames { get; set; } = new List<string>();
        public int RemoteImageCount { get; set; }
        public bool IsUnread { get; set; }
        public MessageSnapshot Restore()
        {
            return new MessageSnapshot(EntryId, StoreId, Subject, Sender, Recipients, ReceivedAt,
                Body, AttachmentNames, RemoteImageCount, IsUnread);
        }
    }
    public sealed class SavedSelection
    {
        public string AttachmentId { get; set; }
        public bool Saved { get; set; }
        public string WorkbookIdentity { get; set; }
        public string WorkbookName { get; set; }
        public int WindowHandle { get; set; }
        public string WorksheetName { get; set; }
        public string Address { get; set; }
        public int StartRow { get; set; }
        public int StartColumn { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public string Preview { get; set; }
        public bool PreviewTruncated { get; set; }
        public ExcelSelectionSnapshot Restore()
        {
            return new ExcelSelectionSnapshot(AttachmentId, Saved, WorkbookIdentity, WorkbookName,
                WindowHandle, WorksheetName, Address, StartRow, StartColumn, RowCount, ColumnCount, Preview, PreviewTruncated);
        }
    }
    public sealed class SavedKoreanCell
    {
        public string WorksheetName { get; set; }
        public string Address { get; set; }
        public string SourceText { get; set; }
    }
    public sealed class SavedKoreanWorkbook
    {
        public bool Saved { get; set; }
        public string WorkbookIdentity { get; set; }
        public string WorkbookName { get; set; }
        public int WindowHandle { get; set; }
        public List<SavedKoreanCell> Cells { get; set; } = new List<SavedKoreanCell>();
        public int SkippedFormulaCells { get; set; }
        public int SkippedMergedCells { get; set; }
        public KoreanWorkbookSnapshot Restore()
        {
            return new KoreanWorkbookSnapshot(Saved, WorkbookIdentity, WorkbookName, WindowHandle,
                Cells.Select(c => new KoreanWorkbookCellSnapshot(c.WorksheetName, c.Address, c.SourceText)).ToArray(),
                SkippedFormulaCells, SkippedMergedCells);
        }
    }

    internal static class OfficeTaskBinding
    {
        internal static TaskSourceBinding Capture(string host, object application)
        {
            if (host == "outlook") return null;
            dynamic app = application;
            dynamic document = host == "excel" ? app.ActiveWorkbook : host == "word" ? app.ActiveDocument : app.ActivePresentation;
            if (document == null) return null;
            string name = Convert.ToString(document.Name);
            string path = Convert.ToString(document.Path);
            string fullName = Convert.ToString(document.FullName);
            var saved = !string.IsNullOrEmpty(path);
            return new TaskSourceBinding
            {
                Id = host + ":" + (saved ? fullName.ToUpperInvariant() : name),
                Location = saved ? fullName : name, Saved = saved, SessionId = TaskRecoveryInput.ProcessSession,
                Fingerprint = saved && File.Exists(fullName) ?
                    TaskCheckpointStore.Fingerprint(new FileInfo(fullName).Length + ":" + File.GetLastWriteTimeUtc(fullName).Ticks) :
                    TaskCheckpointStore.Fingerprint(name + TaskRecoveryInput.ProcessSession)
            };
        }

        internal static void Validate(DurableTaskState state, string host, object application)
        {
            if (state.ProcessSession != TaskRecoveryInput.ProcessSession && state.Sources.Any(s => !s.Saved))
                throw new InvalidOperationException("An unsaved source disappeared when Office closed. Reopen/reselect its original content in a new task; it cannot be rebound by document name.");
            if (state.Sources.Count == 0 || host == "outlook") return;
            dynamic app = application;
            dynamic documents = host == "excel" ? app.Workbooks : host == "word" ? app.Documents : app.Presentations;
            var candidates = new List<object>();
            for (var index = 1; index <= (int)documents.Count; index++)
            {
                dynamic document = documents.Item(index);
                string location = Convert.ToString(document.FullName);
                if (string.Equals(location, state.Sources[0].Location, StringComparison.OrdinalIgnoreCase)) candidates.Add((object)document);
            }
            if (candidates.Count == 1)
            {
                dynamic document = candidates[0];
                if (host == "powerpoint") document.Windows.Item(1).Activate(); else document.Activate();
                if (host == "excel")
                {
                    var recovery = TaskRecoveryInput.Read(state);
                    if (recovery.Selection != null) document.Worksheets.Item(recovery.Selection.WorksheetName).Activate();
                }
            }
            var current = Capture(host, application);
            // Excel's durable transform verifies every source and destination cell,
            // including user-saved output from before interruption.
            var cellValidated = host == "excel" && state.HostData.ContainsKey("excel_transform");
            if (candidates.Count != 1 || !state.Sources.All(s => s.Matches(current) ||
                (cellValidated && current != null && s.Id == current.Id && s.Location == current.Location && s.Saved == current.Saved)))
                throw new InvalidOperationException("Reopen the original uniquely identified unchanged document before resuming: " + state.Sources[0].Location);
        }
    }
}
