using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace Scribble.Chat
{
    public enum TaskLifecycle { Running, AwaitingUser, Paused, Completed, Discarded }

    public sealed class TaskSourceBinding
    {
        public string Id { get; set; }
        public string Location { get; set; }
        public string Fingerprint { get; set; }
        public bool Saved { get; set; }
        public string SessionId { get; set; }

        public bool Matches(TaskSourceBinding current)
        {
            return current != null && !string.IsNullOrEmpty(Id) &&
                !string.IsNullOrEmpty(Fingerprint) && (!Saved || !string.IsNullOrEmpty(Location)) &&
                Id == current.Id && Location == current.Location &&
                Fingerprint == current.Fingerprint && Saved == current.Saved &&
                (Saved || (!string.IsNullOrEmpty(SessionId) && SessionId == current.SessionId));
        }
    }

    // Issued by host code from the original instruction; never inferred from a summary.
    public sealed class TaskAuthorization
    {
        public string OriginalInstruction { get; set; }
        public string Operation { get; set; }
        public TaskSourceBinding Source { get; set; }
        public TaskSourceBinding Destination { get; set; }

        public bool Allows(string operation, TaskSourceBinding source, TaskSourceBinding destination)
        {
            return !string.IsNullOrWhiteSpace(OriginalInstruction) && Operation == operation &&
                Source != null && Source.Matches(source) &&
                Destination != null && Destination.Matches(destination);
        }
    }

    public sealed class TaskBatchResult
    {
        public string Id { get; set; }
        public List<string> CoveredSourceIds { get; set; } = new List<string>();
        public string Output { get; set; }
        public List<string> EvidenceReferences { get; set; } = new List<string>();
        public List<string> Failures { get; set; } = new List<string>();
    }

    public sealed class TaskWriteRecord
    {
        public string Id { get; set; }
        public string BeforeFingerprint { get; set; }
        public string AfterFingerprint { get; set; }
        // Pending must be read back after interruption before any retry.
        public string Status { get; set; }
    }

    public sealed class DurableTaskState
    {
        public int Version { get; set; } = 1;
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string Host { get; set; }
        public string Objective { get; set; }
        public List<string> OriginalDecisions { get; set; } = new List<string>();
        public TaskLifecycle Lifecycle { get; set; } = TaskLifecycle.Running;
        public List<TaskSourceBinding> Sources { get; set; } = new List<TaskSourceBinding>();
        public List<string> ExpectedSourceIds { get; set; } = new List<string>();
        public bool EnumerationComplete { get; set; }
        public Dictionary<string, string> Exclusions { get; set; } = new Dictionary<string, string>();
        public List<TaskBatchResult> Batches { get; set; } = new List<TaskBatchResult>();
        public List<TaskWriteRecord> Writes { get; set; } = new List<TaskWriteRecord>();
        public List<TaskAuthorization> Authorizations { get; set; } = new List<TaskAuthorization>();
        public string Cursor { get; set; }
        public string Blocker { get; set; }

        public string[] Outstanding()
        {
            var covered = new HashSet<string>(Batches.Where(b => b.Failures.Count == 0)
                .SelectMany(b => b.CoveredSourceIds), StringComparer.Ordinal);
            return ExpectedSourceIds.Where(id => !covered.Contains(id) && !Exclusions.ContainsKey(id))
                .Distinct(StringComparer.Ordinal).ToArray();
        }

        public bool CanComplete(bool reviewAll)
        {
            var expected = new HashSet<string>(ExpectedSourceIds, StringComparer.Ordinal);
            var covered = Batches.SelectMany(b => b.CoveredSourceIds).ToArray();
            return EnumerationComplete && Outstanding().Length == 0 &&
                covered.Distinct(StringComparer.Ordinal).Count() == covered.Length &&
                Batches.All(b => b.Failures.Count == 0 && b.CoveredSourceIds.All(expected.Contains)) &&
                Writes.All(w => w.Status == "verified") &&
                (!reviewAll || Exclusions.Count == 0) &&
                Exclusions.All(e => expected.Contains(e.Key) && !string.IsNullOrWhiteSpace(e.Value));
        }
    }

    // Checkpoints contain mailbox/document data, so protect both state and evidence
    // with DPAPI CurrentUser. A failed replacement leaves the previous checkpoint intact.
    public sealed class TaskCheckpointStore
    {
        private readonly string _root;
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        public TaskCheckpointStore(string root = null)
        {
            _root = root ?? Path.Combine(Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData), "Scribble", "Tasks");
        }

        private string TaskDirectory(string id)
        {
            Guid parsed;
            if (!Guid.TryParseExact(id, "N", out parsed)) throw new ArgumentException("Invalid task ID.");
            return Path.Combine(_root, parsed.ToString("N"));
        }

        public void Save(DurableTaskState state)
        {
            if (state == null || state.Version != 1) throw new ArgumentException("Unsupported checkpoint.");
            if (state.Lifecycle == TaskLifecycle.Discarded) throw new InvalidOperationException("Discarded tasks cannot be checkpointed.");
            WriteProtected(Path.Combine(TaskDirectory(state.Id), "state.dat"), _json.Serialize(state));
        }

        public DurableTaskState Load(string id)
        {
            var state = _json.Deserialize<DurableTaskState>(ReadProtected(
                Path.Combine(TaskDirectory(id), "state.dat")));
            if (state == null || state.Version != 1 || state.Id != id)
                throw new InvalidDataException("Unsupported or mismatched checkpoint.");
            return state;
        }

        public string PutEvidence(string taskId, string text)
        {
            var id = Fingerprint(text);
            WriteProtected(Path.Combine(TaskDirectory(taskId), id + ".dat"), text);
            return id;
        }

        public string ReadEvidence(string taskId, string id)
        {
            if (id == null || id.Length != 64 || id.Any(c => !Uri.IsHexDigit(c)))
                throw new ArgumentException("Invalid evidence ID.");
            var text = ReadProtected(Path.Combine(TaskDirectory(taskId), id + ".dat"));
            if (Fingerprint(text) != id) throw new InvalidDataException("Evidence fingerprint mismatch.");
            return text;
        }

        public void Discard(string id)
        {
            var path = TaskDirectory(id);
            if (Directory.Exists(path)) Directory.Delete(path, true);
        }

        public static string Fingerprint(string text)
        {
            using (var hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(text ?? "")))
                    .Replace("-", "").ToLowerInvariant();
        }

        private static string ReadProtected(string path)
        {
            return Encoding.UTF8.GetString(ProtectedData.Unprotect(File.ReadAllBytes(path),
                null, DataProtectionScope.CurrentUser));
        }

        private static void WriteProtected(string path, string text)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                var bytes = ProtectedData.Protect(Encoding.UTF8.GetBytes(text ?? ""),
                    null, DataProtectionScope.CurrentUser);
                using (var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write,
                    FileShare.None, 4096, FileOptions.WriteThrough))
                {
                    file.Write(bytes, 0, bytes.Length);
                    file.Flush(true);
                }
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }

    public sealed class TaskCoordinator
    {
        private readonly TaskCheckpointStore _store;
        public DurableTaskState State { get; private set; }

        public TaskCoordinator(DurableTaskState state, TaskCheckpointStore store)
        {
            State = state ?? throw new ArgumentNullException(nameof(state));
            _store = store ?? throw new ArgumentNullException(nameof(store));
        }

        public void Checkpoint() { _store.Save(State); }

        public void Pause(string blocker)
        {
            State.Blocker = blocker;
            State.Lifecycle = TaskLifecycle.Paused;
            Checkpoint();
        }

        public bool Resume(Func<TaskSourceBinding, bool> validate)
        {
            if (State.Lifecycle == TaskLifecycle.Completed || State.Lifecycle == TaskLifecycle.Discarded)
                return false;
            if (validate == null || State.Sources.Any(s => !validate(s)))
            {
                Pause("Reopen or reselect the original sources; their identities could not be validated.");
                return false;
            }
            if (State.Writes.Any(w => w.Status != "verified"))
            {
                Pause("Read back the pending destination writes before resuming.");
                return false;
            }
            State.Blocker = null;
            State.Lifecycle = TaskLifecycle.Running;
            Checkpoint();
            return true;
        }

        public async Task RunBatchesAsync(IEnumerable<string[]> batches,
            Func<string[], CancellationToken, Task<TaskBatchResult>> process,
            Action<int, int> progress, CancellationToken cancellationToken)
        {
            if (State.Lifecycle != TaskLifecycle.Running) throw new InvalidOperationException("Task is not running.");
            try
            {
                foreach (var ids in batches)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var outstanding = new HashSet<string>(State.Outstanding(), StringComparer.Ordinal);
                    var pending = ids.Where(outstanding.Contains).Distinct(StringComparer.Ordinal).ToArray();
                    if (pending.Length == 0) continue;
                    var result = await process(pending, cancellationToken).ConfigureAwait(true);
                    if (result == null || string.IsNullOrEmpty(result.Id) ||
                        result.Failures.Count > 0 ||
                        result.CoveredSourceIds.Count != pending.Length ||
                        !new HashSet<string>(result.CoveredSourceIds, StringComparer.Ordinal).SetEquals(pending) ||
                        State.Batches.Any(b => b.Id == result.Id))
                        throw new InvalidOperationException("Batch coverage did not reconcile; retry the outstanding sources.");
                    State.Batches.Add(result);
                    Checkpoint();
                    progress?.Invoke(State.ExpectedSourceIds.Count - State.Outstanding().Length,
                        State.ExpectedSourceIds.Count);
                    // Captures the owning Office synchronization context; never moves COM onto a worker.
                    await Task.Yield();
                }
            }
            catch (OperationCanceledException) { Pause("Stopped by user. Resume to continue outstanding batches."); throw; }
            catch (Exception ex) { Pause(ex.Message); throw; }
        }

        public void Complete(bool reviewAll)
        {
            if (State.Lifecycle != TaskLifecycle.Running) throw new InvalidOperationException("Resume the task before completing it.");
            if (!State.CanComplete(reviewAll)) throw new InvalidOperationException("Task coverage is incomplete.");
            State.Lifecycle = TaskLifecycle.Completed;
            Checkpoint();
        }

        public void Discard()
        {
            _store.Discard(State.Id);
            State.Lifecycle = TaskLifecycle.Discarded;
        }
    }
}
