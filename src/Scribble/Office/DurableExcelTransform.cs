using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Scribble.Chat;

namespace Scribble.Office
{
    public sealed class ExcelTaskCell
    {
        public string Id { get; set; }
        public string Sheet { get; set; }
        public string Address { get; set; }
        public string Value { get; set; }
        public string Formula { get; set; }
        public bool HasFormula { get; set; }
        public bool Merged { get; set; }
    }

    public interface IExcelTransformTarget
    {
        IReadOnlyList<ExcelTaskCell> ReadSources(int offset, int count);
        IReadOnlyList<ExcelTaskCell> ReadDestination(int offset, int count, string column, bool replaceSource);
        void Write(int offset, IReadOnlyList<string> values, string column, bool replaceSource);
    }

    public sealed class ExcelTransformPart
    {
        public int Offset { get; set; }
        public int Count { get; set; }
        public string Evidence { get; set; }
        public string Status { get; set; }
    }

    public sealed class ExcelTransformCheckpoint
    {
        public int Expected { get; set; }
        public int Staged { get; set; }
        public string Destination { get; set; }
        public bool ReplaceSource { get; set; }
        public bool SourceComplete { get; set; }
        public bool Committed { get; set; }
        public List<ExcelTransformPart> Sources { get; set; } = new List<ExcelTransformPart>();
        public List<ExcelTransformPart> Outputs { get; set; } = new List<ExcelTransformPart>();
        public List<ExcelTransformPart> Writes { get; set; } = new List<ExcelTransformPart>();
        public string Terminology { get; set; } = "";
    }

    // Source data and reviewed output live in encrypted chunks. The small state
    // record contains only offsets and evidence references, not 20,000 cells per write.
    public sealed class DurableExcelTransform
    {
        private readonly TaskContextManager _task;
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        private readonly List<ExcelTaskCell> _sources = new List<ExcelTaskCell>();
        private readonly List<string> _values = new List<string>();
        public ExcelTransformCheckpoint State { get; private set; }
        public IReadOnlyList<ExcelTaskCell> Sources { get { return _sources; } }
        public IReadOnlyList<string> Values { get { return _values; } }

        public DurableExcelTransform(TaskContextManager task, int expected)
        {
            _task = task;
            string saved;
            State = task.State.HostData.TryGetValue("excel_transform", out saved)
                ? _json.Deserialize<ExcelTransformCheckpoint>(saved) : new ExcelTransformCheckpoint { Expected = expected };
            if (State.Expected != expected) throw new InvalidOperationException("The captured Excel range size changed.");
            foreach (var part in State.Sources) _sources.AddRange(_json.Deserialize<List<ExcelTaskCell>>(Read(part.Evidence)));
            foreach (var part in State.Outputs) _values.AddRange(_json.Deserialize<List<string>>(Read(part.Evidence)));
            if (_values.Count != State.Staged) throw new InvalidOperationException("The Excel staging ledger is inconsistent.");
        }

        public async Task CaptureAsync(IExcelTransformTarget target, CancellationToken token)
        {
            if (State.SourceComplete) return;
            for (var offset = _sources.Count; offset < State.Expected;)
            {
                token.ThrowIfCancellationRequested();
                var count = Math.Min(100, State.Expected - offset);
                var rows = target.ReadSources(offset, count);
                if (rows.Count != count) throw new InvalidOperationException("Source capture omitted rows.");
                if (rows.Any(r => r.Merged)) throw new InvalidOperationException("Unmerge the captured source cells before transforming this range.");
                State.Sources.Add(new ExcelTransformPart { Offset = offset, Count = count, Evidence = Put(rows) });
                _sources.AddRange(rows);
                foreach (var row in rows)
                    if (!_task.State.ExpectedSourceIds.Contains(row.Id)) _task.State.ExpectedSourceIds.Add(row.Id);
                offset += count;
                Save();
                await Task.Yield();
            }
            State.SourceComplete = true;
            Save();
        }

        public void StageReviewed(int offset, IReadOnlyList<string> values, string destination, bool replaceSource)
        {
            if (!State.SourceComplete || values == null || values.Count == 0 || offset < 0 || offset + values.Count > State.Expected)
                throw new InvalidOperationException("Reviewed output does not cover a valid source window.");
            var literals = values.Select(ExcelSelectionOutputPolicy.SanitizeLiteral).ToArray();
            if (values.Any(value => value != null && value.Length > ExcelSelectionOutputPolicy.MaxCellCharacters))
                throw new InvalidOperationException("A reviewed value exceeds Excel's cell length. Repair it explicitly; output was not silently truncated.");
            if (offset < State.Staged)
            {
                if (offset + literals.Length <= State.Staged && _values.Skip(offset).Take(literals.Length).SequenceEqual(literals)) return;
                throw new InvalidOperationException("A retry changed already reviewed output.");
            }
            if (offset != State.Staged) throw new InvalidOperationException("Review all preceding rows before this batch.");
            if (State.Staged > 0 && (State.Destination != destination || State.ReplaceSource != replaceSource))
                throw new InvalidOperationException("The captured output destination cannot change between batches.");
            State.Destination = destination;
            State.ReplaceSource = replaceSource;
            State.Outputs.Add(new ExcelTransformPart { Offset = offset, Count = literals.Length, Evidence = Put(literals), Status = "reviewed" });
            _values.AddRange(literals);
            State.Staged += literals.Length;
            Save();
        }

        public async Task CommitAsync(IExcelTransformTarget target, CancellationToken token, Action<int, int> progress = null)
        {
            if (State.Staged != State.Expected || State.Outputs.Any(p => p.Status != "reviewed"))
                throw new InvalidOperationException("Every source row must be staged and reviewed before writing.");
            // Validate the entire range before the first mutation, including cells
            // already written before interruption. Reopening an unsaved output file
            // restores missing writes only when its original source still matches.
            for (var offset = 0; offset < State.Expected; offset += 100)
            {
                token.ThrowIfCancellationRequested();
                ValidateWindow(target, offset, Math.Min(100, State.Expected - offset));
                await Task.Yield();
            }
            for (var offset = 0; offset < State.Expected; offset += 100)
            {
                token.ThrowIfCancellationRequested();
                var count = Math.Min(100, State.Expected - offset);
                var part = State.Writes.FirstOrDefault(p => p.Offset == offset);
                if (part == null)
                {
                    part = new ExcelTransformPart { Offset = offset, Count = count, Status = "pending" };
                    State.Writes.Add(part);
                }
                var current = target.ReadDestination(offset, count, State.Destination, State.ReplaceSource);
                var output = _values.Skip(offset).Take(count).ToArray();
                if (!OutputMatches(current, output))
                {
                    ValidateWindow(target, offset, count);
                    part.Status = "pending";
                    Save(); // Must be durable before changing NumberFormat or Value2.
                    target.Write(offset, output, State.Destination, State.ReplaceSource);
                    current = target.ReadDestination(offset, count, State.Destination, State.ReplaceSource);
                    if (!OutputMatches(current, output)) throw new InvalidOperationException("Excel output readback differs at row offset " + offset + ". The write remains uncertain.");
                }
                part.Status = "verified";
                var batchId = "excel-write:" + offset;
                if (!_task.State.Batches.Any(b => b.Id == batchId))
                    _task.State.Batches.Add(new TaskBatchResult { Id = batchId,
                        CoveredSourceIds = _sources.Skip(offset).Take(count).Select(c => c.Id).ToList(),
                        EvidenceReferences = new List<string> { Put(current) } });
                Save();
                progress?.Invoke(offset + count, State.Expected);
                await Task.Yield();
            }
            State.Committed = true;
            Save();
        }

        private void ValidateWindow(IExcelTransformTarget target, int offset, int count)
        {
            var sources = target.ReadSources(offset, count);
            var destination = target.ReadDestination(offset, count, State.Destination, State.ReplaceSource);
            if (sources.Count != count || destination.Count != count) throw new InvalidOperationException("The source/destination range changed size.");
            var journaled = State.Writes.Any(p => p.Offset == offset);
            for (var i = 0; i < count; i++)
            {
                var original = _sources[offset + i];
                var source = sources[i];
                var cell = destination[i];
                var outputPresent = journaled && LiteralMatches(cell.Value, _values[offset + i]) && !cell.HasFormula && !cell.Merged;
                if (source.Id != original.Id || source.Merged ||
                    (!(State.ReplaceSource && outputPresent) && (source.Value != original.Value || source.Formula != original.Formula || source.HasFormula != original.HasFormula)))
                    throw new InvalidOperationException("Source changed at " + original.Sheet + "!" + original.Address + ". Reconcile it before resuming.");
                if (State.ReplaceSource && source.HasFormula)
                    throw new InvalidOperationException("Source formulas must be preserved. Choose an adjacent empty column for this transformation.");
                if (!State.ReplaceSource && !outputPresent && (cell.HasFormula || cell.Merged || !string.IsNullOrEmpty(cell.Value)))
                    throw new InvalidOperationException("Destination occupied or changed at row offset " + (offset + i) + ". No conflicting value was overwritten.");
            }
        }

        public static bool LiteralMatches(string actual, string expected)
        {
            return actual == expected || (expected != null && expected.Length > 1 && expected[0] == '\'' &&
                "=+-@".IndexOf(expected[1]) >= 0 && actual == expected.Substring(1));
        }
        private static bool OutputMatches(IReadOnlyList<ExcelTaskCell> cells, IReadOnlyList<string> values)
        {
            return cells.Count == values.Count && cells.Select((c, i) => !c.HasFormula && !c.Merged && LiteralMatches(c.Value, values[i])).All(v => v);
        }
        public void Save()
        {
            _task.State.HostData["excel_transform"] = _json.Serialize(State);
            _task.Checkpoint();
        }
        private string Put(object value) { return _task.Store.PutEvidence(_task.State.Id, _json.Serialize(value)); }
        private string Read(string id) { return _task.Store.ReadEvidence(_task.State.Id, id); }
    }
}
