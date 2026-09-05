using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Outlook;

namespace GuardrailTests
{
    internal static class ScaleTaskTests
    {
        private static void Check(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }

        public static void MailboxPagination()
        {
            foreach (var count in new[] { 500, 1000 })
            {
                var now = DateTime.Now;
                var app = new FakeOutlookApplication();
                var items = Enumerable.Range(0, count).Select(i => new FakeSelectedMailItem(
                    "scale-" + i, "Subject " + i, now.AddMinutes(-i), true)
                {
                    Body = i == count - 1 ? new string('x', 24000) + "IMPORTANT LATE EVIDENCE" : "Short body"
                }).ToArray();
                // Include a duplicate table row: identity reconciliation must remove it.
                app.Session.RegisterFolder(6, new FakeMailFolder { Items = new FakeSearchItems(items.Concat(new[] { items[0] })) });
                var json = new JavaScriptSerializer();
                var seen = new HashSet<string>();
                var cursor = "";
                var pages = 0;
                var foundLateEvidence = false;
                using (var host = new MailboxToolHost(app, null))
                {
                    do
                    {
                        var response = host.ExecuteAsync(Call(MailboxToolCatalog.SearchMailbox, json.Serialize(new
                        {
                            query = "", folder = "inbox", days_back = 10, max_results = 37, cursor
                        })), CancellationToken.None).GetAwaiter().GetResult();
                        var page = (Dictionary<string, object>)json.DeserializeObject(response.Content);
                        Check(!page.ContainsKey("error_code"), response.Content);
                        foreach (var raw in (object[])page["results"])
                        {
                            var hit = (Dictionary<string, object>)raw;
                            Check(seen.Add((string)hit["source_id"]), "Duplicate source escaped pagination.");
                            int? offset = 0;
                            var reconstructed = new StringBuilder();
                            do
                            {
                                var read = host.Execute(Call(MailboxToolCatalog.ReadMessages, json.Serialize(new
                                {
                                    handles = new[] { (string)hit["handle"] }, body_offset = offset.Value
                                })));
                                var result = (Dictionary<string, object>)json.DeserializeObject(read.Content);
                                var body = (Dictionary<string, object>)((object[])result["messages"])[0];
                                Check(!body.ContainsKey("error_code"), read.Content);
                                reconstructed.Append(body["body"]);
                                offset = body["next_body_offset"] == null ? null : (int?)Convert.ToInt32(body["next_body_offset"]);
                            } while (offset.HasValue);
                            foundLateEvidence |= reconstructed.ToString().EndsWith("IMPORTANT LATE EVIDENCE");
                        }
                        cursor = (string)page["next_cursor"];
                        Check(++pages < 100, "Cursor did not terminate.");
                    } while (cursor.Length > 0);
                }
                Check(seen.Count == count && foundLateEvidence, "Mailbox coverage or long-body paging was incomplete.");
            }
        }

        public static void RestartAndCoverage()
        {
            var root = Path.Combine(Path.GetTempPath(), "scribble-task-test-" + Guid.NewGuid().ToString("N"));
            var store = new TaskCheckpointStore(root);
            var state = new DurableTaskState { Host = "test", Objective = "Review all", EnumerationComplete = true };
            state.ExpectedSourceIds.AddRange(Enumerable.Range(0, 20000).Select(i => "row-" + i));
            var coordinator = new TaskCoordinator(state, store);
            try
            {
                var evidence = store.PutEvidence(state.Id, "Confidential source evidence");
                Check(store.ReadEvidence(state.Id, evidence) == "Confidential source evidence", "Evidence did not round trip.");
                Check(!Encoding.UTF8.GetString(File.ReadAllBytes(Path.Combine(root, state.Id, evidence + ".dat")))
                    .Contains("Confidential"), "Evidence was stored in plaintext.");
                coordinator.Checkpoint();
                Check(!state.CanComplete(true), "An unprocessed task completed.");
                var batches = state.ExpectedSourceIds.Select((id, i) => new { id, i }).GroupBy(x => x.i / 100)
                    .Select(g => g.Select(x => x.id).ToArray()).ToArray();
                using (var stop = new CancellationTokenSource())
                {
                    var calls = 0;
                    try
                    {
                        coordinator.RunBatchesAsync(batches, (ids, token) =>
                        {
                            if (++calls == 7) stop.Cancel();
                            token.ThrowIfCancellationRequested();
                            return Task.FromResult(new TaskBatchResult { Id = ids[0], CoveredSourceIds = ids.ToList() });
                        }, null, stop.Token).GetAwaiter().GetResult();
                    }
                    catch (OperationCanceledException) { }
                }
                var resumed = store.Load(state.Id);
                Check(resumed.Lifecycle == TaskLifecycle.Paused && resumed.Outstanding().Length == 19400,
                    "Stop lost the completed batch ledger.");
                coordinator = new TaskCoordinator(resumed, store);
                Check(coordinator.Resume(binding => true), "Validated task could not resume.");
                coordinator.RunBatchesAsync(batches, (ids, token) => Task.FromResult(new TaskBatchResult
                {
                    Id = ids[0], CoveredSourceIds = ids.ToList()
                }), null, CancellationToken.None).GetAwaiter().GetResult();
                Check(resumed.Batches.Count == 200 && resumed.CanComplete(true), "Restart duplicated or omitted rows.");
                resumed.Writes.Add(new TaskWriteRecord { Id = "pending-write", Status = "pending" });
                Check(!resumed.CanComplete(true) && !coordinator.Resume(binding => true), "Uncertain writes were retried without readback.");
                resumed.Writes[0].Status = "verified";
                Check(coordinator.Resume(binding => true), "Verified writes prevented resumption.");
                coordinator.Complete(true);
                Check(store.Load(state.Id).Lifecycle == TaskLifecycle.Completed, "Completion did not persist.");
                var original = new TaskSourceBinding { Id = "Book1", Location = "", Saved = false, SessionId = "old" };
                Check(!original.Matches(new TaskSourceBinding { Id = "Book1", Location = "", Saved = false, SessionId = "new" }),
                    "A disappeared unsaved workbook rebound to another session.");
                coordinator.Discard();
                Check(!Directory.Exists(Path.Combine(root, state.Id)), "Discard left private task data behind.");
            }
            finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        }

        private static ChatToolCall Call(string name, string arguments)
        {
            return new ChatToolCall { id = Guid.NewGuid().ToString("N"), type = "function",
                function = new ChatToolCallFunction { name = name, arguments = arguments } };
        }
    }
}
