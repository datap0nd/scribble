using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Scribble.Chat;
using Scribble.Office;
using Scribble.Outlook;

namespace GuardrailTests
{
    internal static class DurableTransformTests
    {
        private static void Check(bool value, string message) { if (!value) throw new InvalidOperationException(message); }
        private static ChatCompletionRequest Request() { return new ChatCompletionRequest { model = "test", messages = new List<object> { new ChatCompletionInputMessage { role = "user", content = "Translate every row preserving blanks and formulas" } } }; }
        public static void TwentyThousandRows()
        {
            var root = Path.Combine(Path.GetTempPath(), "scribble-excel-" + Guid.NewGuid().ToString("N"));
            var store = new TaskCheckpointStore(root);
            try
            {
                var task = new TaskContextManager(Request(), "excel", "Translate every row preserving blanks and formulas", store);
                var target = new Target(20000);
                var transform = new DurableExcelTransform(task, 20000);
                transform.CaptureAsync(target, CancellationToken.None).GetAwaiter().GetResult();
                for (var offset = 0; offset < 20000; offset += 100)
                    transform.StageReviewed(offset, Enumerable.Range(offset, 100).Select(i => i % 17 == 0 ? "" : i == 19999 ? new string('L', 25000) : "Reviewed " + i).ToArray(), "B", false);
                target.InterruptBefore = 700;
                try { transform.CommitAsync(target, CancellationToken.None).GetAwaiter().GetResult(); throw new Exception("Missing before-write interruption"); }
                catch (OperationCanceledException) { }
                Check(target.WrittenRows == 700, "Writes continued after cancellation.");
                task = new TaskContextManager(Request(), "excel", task.State.Objective, store, store.Load(task.State.Id));
                transform = new DurableExcelTransform(task, 20000);
                target.InterruptBefore = -1; target.InterruptAfter = 1600;
                try { transform.CommitAsync(target, CancellationToken.None).GetAwaiter().GetResult(); throw new Exception("Missing after-write interruption"); }
                catch (OperationCanceledException) { }
                Check(target.WrittenRows == 1700, "After-write journal did not match applied rows.");
                task = new TaskContextManager(Request(), "excel", task.State.Objective, store, store.Load(task.State.Id));
                transform = new DurableExcelTransform(task, 20000); target.InterruptAfter = -1;
                transform.CommitAsync(target, CancellationToken.None).GetAwaiter().GetResult();
                Check(transform.State.Committed && target.WrittenRows == 20000 && target.WriteCounts.All(c => c == 1), "Restart duplicated or omitted output writes.");
                Check(target.Destination[19999].Value.Length == 25000 && target.Source[1].HasFormula && target.Destination[17].Value == "", "Long cells, blanks or formulas were lost.");
                task.State.EnumerationComplete = true;
                Check(task.State.CanComplete(true) && task.State.Batches.Count == 200, "Final write ledger does not cover all rows.");
                transform.CommitAsync(target, CancellationToken.None).GetAwaiter().GetResult();
                Check(target.WrittenRows == 20000, "Readback reapplied verified output.");
            }
            finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        }

        public static void ChangedRangesFailClosed()
        {
            foreach (var sourceChanged in new[] { false, true })
            {
                var root = Path.Combine(Path.GetTempPath(), "scribble-excel-conflict-" + Guid.NewGuid().ToString("N"));
                try
                {
                    var task = new TaskContextManager(Request(), "excel", "Translate", new TaskCheckpointStore(root));
                    var target = new Target(200); var transform = new DurableExcelTransform(task, 200);
                    transform.CaptureAsync(target, CancellationToken.None).GetAwaiter().GetResult();
                    transform.StageReviewed(0, Enumerable.Repeat("Translation", 200).ToArray(), "B", false);
                    (sourceChanged ? target.Source : target.Destination)[199].Value = "User edit";
                    var blocked = false;
                    try { transform.CommitAsync(target, CancellationToken.None).GetAwaiter().GetResult(); }
                    catch (InvalidOperationException) { blocked = true; }
                    Check(blocked && target.WrittenRows == 0, "A changed source or occupied destination was overwritten.");
                }
                finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
            }
        }

        public static void AttachmentTail()
        {
            var path = Path.Combine(Path.GetTempPath(), "scribble-tail-" + Guid.NewGuid().ToString("N") + ".txt");
            try
            {
                var original = new string('a', 130000) + "IMPORTANT ATTACHMENT TAIL";
                File.WriteAllText(path, original);
                var result = new System.Text.StringBuilder(); int? offset = 0;
                do
                {
                    var page = EmailAttachmentReader.LoadLocalPage(path, offset.Value, 6000, CancellationToken.None);
                    result.Append(page.Text); offset = page.NextOffset;
                } while (offset.HasValue);
                Check(result.ToString() == original, "Attachment paging truncated or duplicated text past the old context cap.");
            }
            finally { File.Delete(path); }
        }

        public static void EveryAttachmentIsRequired()
        {
            var root = Path.Combine(Path.GetTempPath(), "scribble-mail-attachments-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                var shortFile = Path.Combine(root, "short.txt"); File.WriteAllText(shortFile, "Short evidence");
                var longFile = Path.Combine(root, "long.txt"); File.WriteAllText(longFile, new string('a', 65000) + "CRITICAL TWELFTH ATTACHMENT");
                var mail = new FakeSelectedMailItem("attachments", "Attachment review");
                for (var i = 1; i <= 12; i++) mail.Attachments.Add(new FakeOutlookAttachment("file" + i + ".txt", i == 12 ? longFile : shortFile));
                var app = new FakeOutlookApplication(); app.Session.RegisterFolder(6, new FakeMailFolder { Items = new FakeSearchItems(new[] { mail }) });
                var selected = new MessageReader(app).CaptureById(mail.EntryID, "store", true);
                var task = new TaskContextManager(Request(), "outlook", "Review this selected email and all attachments", new TaskCheckpointStore(root));
                var json = new System.Web.Script.Serialization.JavaScriptSerializer();
                using (var host = new MailboxToolHost(app, selected))
                {
                    host.BindTaskAsync(task, CancellationToken.None).GetAwaiter().GetResult();
                    Func<string, object, MailboxToolResult> call = (name, args) => host.ExecuteAsync(new ChatToolCall { id = Guid.NewGuid().ToString("N"), type = "function", function = new ChatToolCallFunction { name = name, arguments = json.Serialize(args) } }, CancellationToken.None).GetAwaiter().GetResult();
                    var body = call(MailboxToolCatalog.ReadMessages, new { handles = new[] { "selected" }, body_offset = 0 });
                    Check(body.Content.Contains("\"attachment_count\":12"), "Attachment enumeration was capped at ten.");
                    Check(call(MailboxToolCatalog.RecordAnalysis, new { handle = "selected", summary = "Premature" }).Content.Contains("MAILBOX_COVERAGE_INCOMPLETE"), "Unread attachments were marked reviewed.");
                    var tail = false;
                    for (var index = 1; index <= 12; index++)
                    {
                        int? offset = 0;
                        do
                        {
                            var page = call(MailboxToolCatalog.ReadAttachment, new { handle = "selected", attachment_index = index, offset = offset.Value });
                            var data = json.Deserialize<Dictionary<string, object>>(page.Content);
                            Check(!data.ContainsKey("error_code"), page.Content);
                            tail |= Convert.ToString(data["content"]).Contains("CRITICAL TWELFTH ATTACHMENT");
                            offset = data["next_offset"] == null ? null : (int?)Convert.ToInt32(data["next_offset"]);
                        } while (offset.HasValue);
                    }
                    Check(tail && !call(MailboxToolCatalog.RecordAnalysis, new { handle = "selected", summary = "CRITICAL TWELFTH ATTACHMENT" }).Content.Contains("error_code") && host.CompletionBlocker == null, "Full attachment review did not reconcile.");
                }
            }
            finally { Directory.Delete(root, true); }
        }

        private sealed class Target : IExcelTransformTarget
        {
            public ExcelTaskCell[] Source, Destination;
            public int[] WriteCounts;
            public int WrittenRows;
            public int InterruptBefore = -1, InterruptAfter = -1;
            public Target(int count)
            {
                Source = Enumerable.Range(0, count).Select(i => new ExcelTaskCell { Id = "row:" + i, Sheet = "Data", Address = "A" + (i + 1), Value = i % 17 == 0 ? "" : "원본 " + i, HasFormula = i == 1, Formula = i == 1 ? "=1+1" : "" }).ToArray();
                Destination = Enumerable.Range(0, count).Select(i => new ExcelTaskCell { Id = "out:" + i, Sheet = "Data", Address = "B" + (i + 1), Value = "", Formula = "" }).ToArray();
                WriteCounts = new int[count];
            }
            private ExcelTaskCell Clone(ExcelTaskCell c) { return new ExcelTaskCell { Id = c.Id, Sheet = c.Sheet, Address = c.Address, Value = c.Value, Formula = c.Formula, HasFormula = c.HasFormula, Merged = c.Merged }; }
            public IReadOnlyList<ExcelTaskCell> ReadSources(int offset, int count) { return Source.Skip(offset).Take(count).Select(Clone).ToArray(); }
            public IReadOnlyList<ExcelTaskCell> ReadDestination(int offset, int count, string column, bool replace) { return (replace ? Source : Destination).Skip(offset).Take(count).Select(Clone).ToArray(); }
            public void Write(int offset, IReadOnlyList<string> values, string column, bool replace)
            {
                if (offset == InterruptBefore) throw new OperationCanceledException();
                for (var i = 0; i < values.Count; i++) { (replace ? Source : Destination)[offset + i].Value = values[i]; WriteCounts[offset + i]++; WrittenRows++; }
                if (offset == InterruptAfter) throw new OperationCanceledException();
            }
        }
    }
}
