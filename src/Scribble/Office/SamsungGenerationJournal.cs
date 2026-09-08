using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;
using Scribble.Chat;

namespace Scribble.Office
{
    // Checkpoints contain identities and receipts, never a saved presentation.
    internal sealed class SamsungGenerationJournal
    {
        public sealed class Receipt
        {
            public int Page { get; set; }
            public string SourceId { get; set; }
            public int PageOrdinal { get; set; }
            public int SlideId { get; set; }
            public string Fingerprint { get; set; }
            public string RepairedContent { get; set; }
            public int Repairs { get; set; }
        }
        public sealed class State
        {
            public string Owner { get; set; }
            public string Input { get; set; }
            public string ToolCall { get; set; }
            public string FunctionFingerprint { get; set; }
            public string ToolName { get; set; }
            public string Arguments { get; set; }
            public List<string> AttemptCalls { get; set; } = new List<string>();
            public int[] OriginalIds { get; set; }
            public int[] LastOrder { get; set; }
            public int Pages { get; set; }
            public List<Receipt> Receipts { get; set; } = new List<Receipt>();
        }
        private readonly TaskContextManager _task;
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = 16000000 };
        internal State Data;
        private object _deck;
        internal SamsungGenerationJournal(TaskContextManager task, ChatToolCall call)
        {
            _task = task;
            string saved;
            if (task.State.HostData.TryGetValue("samsung_pending", out saved))
            {
                Data = _json.Deserialize<State>(saved);
                if (Data.Input != InputHash(call)) throw new InvalidOperationException("SLIDE_RECOVERY_INPUT_CHANGED: Resume with the original generation payload unchanged; inspect the surviving draft before requesting different edits.");
            }
            else Data = new State { Owner = Guid.NewGuid().ToString("N"), ToolCall = call.id, Input = InputHash(call), FunctionFingerprint = TaskCheckpointStore.Fingerprint(_json.Serialize(call.function)), ToolName = call.function.name, Arguments = call.function.arguments };
            if (!Data.AttemptCalls.Contains(call.id)) Data.AttemptCalls.Add(call.id);
            if (Data.OriginalIds != null) Persist();
        }
        internal static string InputHash(ChatToolCall call)
        {
            var json = new JavaScriptSerializer { MaxJsonLength = 16000000 };
            return TaskCheckpointStore.Fingerprint(call.function.name + json.Serialize(SamsungRepairPolicy.Canonical(json.DeserializeObject(call.function.arguments))));
        }
        internal static bool CanResume(DurableTaskState state, ChatToolCall call)
        {
            if (state.SamsungWorkflowVersion < 2 || !(call.function.name == PresentationToolCatalog.AddDraftSlides || call.function.name == CrossAppToolCatalog.SendToPowerPoint)) return false;
            string saved;
            if (!state.HostData.TryGetValue("samsung_pending", out saved)) return false;
            var record = new JavaScriptSerializer { MaxJsonLength = 16000000 }.Deserialize<State>(saved);
            return record.Input == InputHash(call) && state.Writes.Where(w => w.Status != "verified" && w.Id.StartsWith("tool:")).All(w =>
                record.AttemptCalls.Any(id => w.Id == "tool:" + id) || w.Id == "tool:" + record.ToolCall || w.BeforeFingerprint == TaskCheckpointStore.Fingerprint(new JavaScriptSerializer().Serialize(call.function)));
        }
        private static int[] Ids(object value)
        {
            dynamic deck = value;
            var ids = new List<int>();
            for (var i = 1; i <= (int)deck.Slides.Count; i++) ids.Add((int)deck.Slides[i].SlideID);
            return ids.ToArray();
        }
        internal void Bind(object value, int pages)
        {
            _deck = value; dynamic deck = value;
            if (Data.OriginalIds == null)
            {
                Data.OriginalIds = Ids(value); Data.LastOrder = Data.OriginalIds; Data.Pages = pages;
                deck.Tags.Add("ScribbleTask", _task.State.Id);
                _task.State.HostData["samsung_destination"] = _task.State.Id;
                _task.State.HostData["samsung_recovery_payload"] = _task.Store.PutEvidence(_task.State.Id, Data.Arguments);
                Persist(); // Before the first slide mutation.
            }
            else
            {
                if (!SamsungSlideDesign.SameOwner(Convert.ToString(deck.Tags["ScribbleTask"]), _task.State.Id)) throw new InvalidOperationException("SLIDE_RECOVERY_WRONG_DECK");
                ValidateReceipts(Data, pages, Ids(value), id => PresentationInspection.Fingerprint(PresentationInspection.FindSlide(value, id)));
            }
        }
        internal static void ValidateReceipts(State state, int pages, int[] order, Func<int, string> fingerprint)
        {
            if (state.Pages != pages || !order.SequenceEqual(state.LastOrder))
                throw new InvalidOperationException("SLIDE_RECOVERY_UNCERTAIN: The deck has unreceipted or reordered slides. Preserve them for inspection; no write was repeated.");
            foreach (var receipt in state.Receipts)
                if (fingerprint(receipt.SlideId) != receipt.Fingerprint) throw new InvalidOperationException("SLIDE_RECOVERY_USER_EDIT: A surviving slide changed; no user content was overwritten.");
        }
        internal PresentationDraftWriter.SamsungOutput Resume(PresentationDraftWriter.SamsungPage page, int index)
        {
            var receipt = Data.Receipts.SingleOrDefault(r => r.Page == index);
            if (receipt == null) return null;
            if (!string.IsNullOrEmpty(receipt.RepairedContent))
            {
                var parsed = PresentationDraftWriter.ParseSlides(new object[] { _json.DeserializeObject(receipt.RepairedContent) });
                parsed[0].ImageData.AddRange(page.Source.ImageData);
                var replacement = PresentationDraftWriter.ComposeSamsung(parsed);
                if (receipt.PageOrdinal >= replacement.Count) throw new InvalidOperationException("SLIDE_RECOVERY_LAYOUT_CHANGED");
                dynamic deck = _deck; var scale = (float)deck.PageSetup.SlideWidth / SamsungSlideDesign.Width;
                if (Math.Abs(scale - 1) > .001) PresentationDraftWriter.ScaleSamsungPage(replacement[receipt.PageOrdinal], scale);
                page = replacement[receipt.PageOrdinal];
            }
            var output = new PresentationDraftWriter.SamsungOutput { Slide = PresentationInspection.FindSlide(_deck, receipt.SlideId), Page = page, Owner = Data.Owner };
            dynamic slide = output.Slide;
            for (var i = 1; i <= (int)slide.Shapes.Count; i++) output.ShapeIds.Add((int)slide.Shapes[i].Id);
            output.Image = PresentationDraftWriter.ExportSamsung(output);
            return output;
        }
        internal void Record(PresentationDraftWriter.SamsungOutput output, int index, string content = null)
        {
            var receipt = Data.Receipts.SingleOrDefault(r => r.Page == index);
            if (receipt == null) { receipt = new Receipt { Page = index, SlideId = (int)((dynamic)output.Slide).SlideID, SourceId = output.Page.Source.Id, PageOrdinal = Data.Receipts.Count(r => r.SourceId == output.Page.Source.Id) }; Data.Receipts.Add(receipt); }
            receipt.Fingerprint = PresentationInspection.Fingerprint(output.Slide);
            if (content != null) receipt.RepairedContent = content;
            Data.LastOrder = Ids(_deck); Persist();
        }
        internal void BeforeRepair(int index)
        {
            var receipt = Data.Receipts.Single(r => r.Page == index);
            if (receipt.Repairs >= 3) throw new InvalidOperationException("SLIDE_REPAIR_LIMIT: Three repair cycles have already been attempted for this slide.");
            receipt.Repairs++; Persist();
        }
        internal void Complete()
        {
            foreach (var write in _task.State.Writes.Where(w => Data.AttemptCalls.Any(id => w.Id == "tool:" + id) || w.Id == "tool:" + Data.ToolCall || w.BeforeFingerprint == Data.FunctionFingerprint)) { write.Status = "verified"; write.AfterFingerprint = "native_generation_reconciled"; }
            _task.State.HostData.Remove("samsung_pending"); _task.Checkpoint();
        }
        private void Persist() { _task.State.HostData["samsung_pending"] = _json.Serialize(Data); _task.Checkpoint(); }
    }
}
