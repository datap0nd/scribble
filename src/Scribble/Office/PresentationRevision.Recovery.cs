using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;
using Scribble.Chat;
namespace Scribble.Office
{
    internal sealed partial class PresentationRevision
    {
        internal string BatchId = Guid.NewGuid().ToString("N");
        public sealed class RecoveryItem
        {
            public int SlideId { get; set; }
            public int LiveId { get; set; }
            public int BackupId { get; set; }
            public int Index { get; set; }
            public string Before { get; set; }
            public string After { get; set; }
            public string LastKnown { get; set; }
            public string BackupFingerprint { get; set; }
            public bool Started { get; set; }
            public bool Applied { get; set; }
            public bool Deleted { get; set; }
            public List<Dictionary<string, object>> Operations { get; set; }
            public int[] InsertedIds { get; set; }
            public int[] AddedShapeIds { get; set; }
            public Dictionary<string, string> InsertedFingerprints { get; set; }
        }
        public sealed class RecoveryState
        {
            public string BatchId { get; set; }
            public string Status { get; set; }
            public string PresentationId { get; set; }
            public int[] BeforeOrder { get; set; }
            public int[] AfterOrder { get; set; }
            public int[] LastOrder { get; set; }
            public bool StructuralStarted { get; set; }
            public RecoveryItem[] Items { get; set; }
            public Dictionary<string, string> BeforeFingerprints { get; set; }
            public Dictionary<string, string> UnrelatedContent { get; set; }
        }
        internal string Snapshot()
        {
            return new JavaScriptSerializer { MaxJsonLength = 16000000 }.Serialize(new RecoveryState {
                BatchId = BatchId, Status = Status, PresentationId = PresentationInspection.IdentityFor(Presentation),
                BeforeFingerprints = _beforeFingerprints.ToDictionary(p => p.Key.ToString(System.Globalization.CultureInfo.InvariantCulture), p => p.Value), UnrelatedContent = _unrelatedContent.ToDictionary(p => p.Key.ToString(System.Globalization.CultureInfo.InvariantCulture), p => p.Value), BeforeOrder = _beforeOrder, AfterOrder = _afterOrder, LastOrder = _lastKnownOrder, StructuralStarted = _structuralStarted,
                Items = Items.Select(i => new RecoveryItem { SlideId = i.SlideId, LiveId = i.Deleted ? i.SlideId : (int)((dynamic)i.Original).SlideID,
                    BackupId = (int)((dynamic)i.Backup).SlideID, Index = i.Index, Before = i.Before, After = i.After,
                    LastKnown = i.LastKnownContent, BackupFingerprint = i.BackupFingerprint, Started = i.Started, Applied = i.Applied, Deleted = i.Deleted,
                    Operations = i.Operations, InsertedIds = i.InsertedIds.ToArray(), AddedShapeIds = i.AddedShapeIds.ToArray(), InsertedFingerprints = i.InsertedFingerprints.ToDictionary(p => p.Key.ToString(System.Globalization.CultureInfo.InvariantCulture), p => p.Value) }).ToArray()
            });
        }
        internal static PresentationRevision Recover(object application, object presentation, string snapshot)
        {
            var state = new JavaScriptSerializer { MaxJsonLength = 16000000 }.Deserialize<RecoveryState>(snapshot);
            if (state.PresentationId != PresentationInspection.IdentityFor(presentation)) throw new InvalidOperationException("REVISION_RECOVERY_WRONG_DECK");
            dynamic app = application; var matches = new List<object>();
            for (var i = 1; i <= (int)app.Presentations.Count; i++)
            {
                dynamic candidate = app.Presentations[i];
                if (SamsungSlideDesign.SameOwner(Convert.ToString(candidate.Tags["ScribbleRevisionRecovery"]), state.BatchId)) matches.Add((object)candidate);
            }
            if (matches.Count != 1) throw new InvalidOperationException("REVISION_RECOVERY_SESSION_ENDED: The unique unsaved recovery presentation is unavailable. No edit was repeated.");
            var result = new PresentationRevision(presentation) { BatchId = state.BatchId, Status = state.Status, Recovery = matches[0],
                _beforeFingerprints = state.BeforeFingerprints.ToDictionary(p => int.Parse(p.Key, System.Globalization.CultureInfo.InvariantCulture), p => p.Value), _unrelatedContent = state.UnrelatedContent.ToDictionary(p => int.Parse(p.Key, System.Globalization.CultureInfo.InvariantCulture), p => p.Value), _beforeOrder = state.BeforeOrder, _afterOrder = state.AfterOrder, _lastKnownOrder = state.LastOrder, _structuralStarted = state.StructuralStarted };
            foreach (var saved in state.Items)
            {
                var backup = PresentationInspection.FindSlide(result.Recovery, saved.BackupId);
                if (PresentationInspection.Fingerprint(backup) != saved.BackupFingerprint) throw new InvalidOperationException("REVISION_RECOVERY_ORIGINAL_CHANGED: The recovery original was edited. No content was overwritten.");
                var item = new Item { SlideId = saved.SlideId, Index = saved.Index, Original = saved.Deleted ? null : PresentationInspection.FindSlide(presentation, saved.LiveId),
                    Backup = backup, BackupFingerprint = saved.BackupFingerprint, Before = saved.Before, After = saved.After, LastKnownContent = saved.LastKnown,
                    Started = saved.Started, Applied = saved.Applied, Deleted = saved.Deleted, Operations = saved.Operations };
                item.InsertedIds.AddRange(saved.InsertedIds); item.AddedShapeIds.AddRange(saved.AddedShapeIds);
                foreach (var pair in saved.InsertedFingerprints) item.InsertedFingerprints.Add(int.Parse(pair.Key, System.Globalization.CultureInfo.InvariantCulture), pair.Value);
                result.Items.Add(item);
            }
            return result;
        }
        internal string Reconcile()
        {
            VerifyRecoveryOriginals();
            if (Status == "applied")
            {
                if (!Order().SequenceEqual(_afterOrder)) throw new InvalidOperationException("REVISION_RECOVERY_ORDER_CHANGED");
                foreach (var item in Items)
                {
                    if (!item.Deleted && PresentationInspection.Fingerprint(item.Original) != item.After) throw new InvalidOperationException("REVISION_RECOVERY_USER_EDIT");
                    foreach (var pair in item.InsertedFingerprints)
                        if (PresentationInspection.Fingerprint(PresentationInspection.FindSlide(Presentation, pair.Key)) != pair.Value) throw new InvalidOperationException("REVISION_RECOVERY_USER_EDIT");
                }
                VerifyUnrelated();
                Latest.Remove(Presentation); Latest.Add(Presentation, this);
                return Status; // A verified completed batch is never replayed.
            }
            if (Status == "rolled_back" || Status == "reverted") { VerifyRestored(); return Status; }
            try
            {
                if (!Order().SequenceEqual(_lastKnownOrder)) throw new InvalidOperationException("REVISION_RECOVERY_UNCERTAIN_ORDER");
                foreach (var item in Items.Where(i => !i.Deleted))
                    if (PresentationInspection.ContentFingerprint(item.Original) != item.LastKnownContent) throw new InvalidOperationException("REVISION_RECOVERY_UNCERTAIN_CONTENT");
                if (_structuralStarted) RestoreStructure();
                foreach (var item in Items.Where(i => i.Started).Reverse()) RestoreContent(item);
                VerifyRestored(); Status = "rolled_back"; return Status;
            }
            catch { Status = "recovery_required"; throw; }
        }
        internal static bool CanResume(DurableTaskState state, ChatToolCall call)
        {
            string input;
            return state.SamsungWorkflowVersion >= 2 && (call.function.name == PresentationToolCatalog.ReviseSlides || call.function.name == PresentationToolCatalog.RevertSlides) &&
                state.HostData.ContainsKey("powerpoint_revision_snapshot") && state.HostData.TryGetValue("powerpoint_revision_input", out input) &&
                input == SamsungGenerationJournal.InputHash(call) && state.Writes.Where(w => w.Status != "verified" && w.Id.StartsWith("tool:")).All(w =>
                    w.BeforeFingerprint == TaskCheckpointStore.Fingerprint(new JavaScriptSerializer().Serialize(call.function)));
        }
    }
}
