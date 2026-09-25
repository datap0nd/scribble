using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace Scribble.Chat
{
    // Local, DPAPI-protected flight recorder. No credentials or HTTP authorization
    // headers are accepted. The rotating slots bound additional disk use per task.
    public sealed class TaskDiagnostics
    {
        public const int Slots = 128;
        public const int MaxDetailCharacters = 524288;
        private readonly TaskCheckpointStore _store;
        private readonly DurableTaskState _state;
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        private readonly object _gate = new object();
        private readonly string _benchmarkRun = Scribble.Testing.TestLab.ActiveRunId();

        public TaskDiagnostics(TaskCheckpointStore store, DurableTaskState state) { _store = store; _state = state; }
        public string Id { get { return _state.Id; } }

        // Reserve at the transport boundary, so authoring, internal reviewers,
        // compaction and HTTP retries share one durable allowance.
        public void ReserveModelRequest()
        {
            lock (_gate)
            {
                string rawLimit;
                if (!_state.HostData.TryGetValue("delivery_request_limit", out rawLimit)) return;
                int limit;
                if (!int.TryParse(rawLimit, out limit) || limit != 18)
                    throw new InvalidOperationException("DELIVERY_REQUEST_BUDGET_INVALID");
                string rawCount;
                var count = 0;
                if (_state.HostData.TryGetValue("delivery_request_count", out rawCount) &&
                    (!int.TryParse(rawCount, out count) || count < 0 || count > limit))
                    throw new InvalidOperationException("DELIVERY_REQUEST_BUDGET_INVALID");
                if (count >= limit)
                    throw new InvalidOperationException("DELIVERY_REQUEST_LIMIT: The complete task reached its model-request budget. Retain the draft and inspect the first unresolved stage; do not restart the same repair loop.");
                _state.HostData["delivery_request_count"] = (count + 1).ToString();
                _store.Save(_state);
            }
        }

        public void Record(string stage, object detail)
        {
            Scribble.Testing.TestLab.Record(_benchmarkRun, Id, stage, detail);
            lock (_gate)
            {
                string previous;
                long sequence = 0;
                if (_state.HostData.TryGetValue("diagnostic_sequence", out previous)) long.TryParse(previous, out sequence);
                sequence++;
                var text = _json.Serialize(detail);
                var entry = new TaskDiagnosticEntry { Sequence = sequence, Stage = stage, TaskId = Id,
                    Utc = DateTime.UtcNow.ToString("O"), Hash = TaskCheckpointStore.Fingerprint(text),
                    Truncated = text.Length > MaxDetailCharacters,
                    Detail = text.Substring(0, Math.Min(text.Length, MaxDetailCharacters)) };
                _store.WriteDiagnostic(Id, (int)(sequence % Slots), _json.Serialize(entry));
                _state.HostData["diagnostic_sequence"] = sequence.ToString();
                _state.HostData["diagnostic_id"] = Id;
                _store.Save(_state);
            }
        }

        // Offline replay inputs; this API never calls tools, models or applications.
        public IReadOnlyList<TaskDiagnosticEntry> ReadLocalReplay()
        {
            var result = new List<TaskDiagnosticEntry>();
            for (var i = 0; i < Slots; i++)
            {
                var text = _store.ReadDiagnostic(Id, i);
                if (text != null) result.Add(_json.Deserialize<TaskDiagnosticEntry>(text));
            }
            return result.OrderBy(e => e.Sequence).ToArray();
        }

        // Suitable for preview/export: deliberately excludes source text and URLs.
        public string RedactedReport()
        {
            return _json.Serialize(new { diagnostic_id = Id, host = _state.Host,
                lifecycle = _state.Lifecycle.ToString(), entries = ReadLocalReplay().Select(e => new
                { e.Sequence, e.Stage, e.Utc, e.Hash, e.Truncated }) });
        }
    }

    public sealed class TaskDiagnosticEntry
    {
        public long Sequence { get; set; }
        public string Stage { get; set; }
        public string TaskId { get; set; }
        public string Utc { get; set; }
        public string Hash { get; set; }
        public bool Truncated { get; set; }
        public string Detail { get; set; }
    }
}
