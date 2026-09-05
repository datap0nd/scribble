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

        public TaskDiagnostics(TaskCheckpointStore store, DurableTaskState state) { _store = store; _state = state; }
        public string Id { get { return _state.Id; } }

        public void Record(string stage, object detail)
        {
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
