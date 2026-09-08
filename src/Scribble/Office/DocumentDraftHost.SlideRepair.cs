using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Scribble.Chat;
using Scribble.Configuration;

namespace Scribble.Office
{
    public sealed partial class DocumentDraftHost
    {
        private static object Canonical(object value)
        {
            var map = value as IDictionary<string, object>;
            if (map != null) return map.OrderBy(p => p.Key, StringComparer.Ordinal).ToDictionary(p => p.Key, p => Canonical(p.Value));
            var sequence = value as System.Collections.IEnumerable;
            if (sequence != null && !(value is string)) return sequence.Cast<object>().Select(Canonical).ToArray();
            return value;
        }
        private async Task<Dictionary<string, object>> RepairSlideContentAsync(PresentationDraftWriter.SamsungOutput output,
            Dictionary<string, object> original, string findings, string source, string prompt,
            OpenAiCompatibleClient client, AppSettings settings, CancellationToken token, IReadOnlyList<PresentationDraftWriter.SamsungOutput> related = null)
        {
            var response = await ReviewSamsungAsync(client, settings,
                SamsungAuthoringPolicy.Instructions + " Repair this single slide using the specific visual findings. Return JSON only: {\"slides\":[{...complete corrected slide...}]}. " +
                "Keep the ID, all required table rows, chart data, source images and evidence unchanged. You may choose a better Samsung layout and remove redundant wording. " +
                "Do not invent pixel coordinates or remove evidence to make it fit. Schema: " + _serializer.Serialize(PresentationToolCatalog.DraftDefinition().function.parameters),
                _serializer.Serialize(new { original, findings, instruction = prompt }), output.Image, token, Math.Min(32768, Math.Max(8192, _serializer.Serialize(original).Length / 2)));
            var wrapper = _serializer.Deserialize<Dictionary<string, object>>(response);
            var replacements = SamsungAuthoringPolicy.Array(wrapper, "slides");
            if (replacements.Length != 1) throw new InvalidOperationException("SLIDE_REPAIR_COUNT: Repair exactly one slide.");
            var replacement = SamsungAuthoringPolicy.ReadMap(replacements[0]);
            var testCall = new ChatToolCall { id = "repair", function = new ChatToolCallFunction { name = PresentationToolCatalog.AddDraftSlides, arguments = _serializer.Serialize(new { slides = replacements }) } };
            var errors = ToolContractValidator.Validate(testCall, PresentationToolCatalog.DraftDefinition());
            if (errors.Count > 0) throw new InvalidOperationException("SLIDE_REPAIR_SCHEMA: " + string.Join("; ", errors));
            if (SamsungAuthoringPolicy.Text(replacement, "id") != SamsungAuthoringPolicy.Text(original, "id")) throw new InvalidOperationException("SLIDE_REPAIR_ID_CHANGED");
            foreach (var field in new[] { "table", "secondary_table", "chart", "secondary_chart", "image_names", "source_spans", "evidence", "calculations", "content_kind" })
            {
                object before, after; original.TryGetValue(field, out before); replacement.TryGetValue(field, out after);
                if (_serializer.Serialize(Canonical(before)) != _serializer.Serialize(Canonical(after))) throw new InvalidOperationException("SLIDE_REPAIR_EVIDENCE_CHANGED: " + field);
            }
            if (SamsungRepairPolicy.Serialize(original) == SamsungRepairPolicy.Serialize(replacement)) throw new InvalidOperationException("SLIDE_REPAIR_STALLED: No meaningful content or layout change was proposed.");
            SamsungPresentationReview.ValidateEvidence(_serializer.Serialize(replacement), source);
            var review = await ReviewSamsungAsync(client, settings, SamsungAuthoringPolicy.FactReview +
                " Verify that every required point from the original remains represented. Reject omitted commitments, qualifications, owners or dates." + SamsungAuthoringPolicy.ReviewContract,
                _serializer.Serialize(new { original, replacement, source, instruction = prompt }), null, token);
            if (!ReviewApproved(review)) throw new InvalidOperationException("SLIDE_REPAIR_FACTS: " + review);
            var parsed = PresentationDraftWriter.ParseSlides(replacements);
            foreach (var image in output.Page.Source.ImageData) parsed[0].ImageData.Add(image);
            var pages = PresentationDraftWriter.ComposeSamsung(parsed);
            var targets = related ?? new[] { output };
            if (pages.Count != targets.Count) throw new InvalidOperationException("SLIDE_REPAIR_COUNT: The complete evidence requires " + pages.Count + " pages in this recipe, but " + targets.Count + " pages are allocated. Choose a fitting recipe or ask which count/content constraint may change.");
            dynamic deck = ((dynamic)output.Slide).Parent;
            var scale = (float)deck.PageSetup.SlideWidth / SamsungSlideDesign.Width;
            if (Math.Abs(scale - 1) > .001) foreach (var page in pages) PresentationDraftWriter.ScaleSamsungPage(page, scale);
            // Check every continuation before replacing any owned content.
            foreach (var target in targets)
                if (PresentationDraftWriter.ExportSamsung(target) != target.Image) throw new InvalidOperationException("SLIDE_CHANGED_DURING_REPAIR");
            for (var i = 0; i < pages.Count; i++) PresentationDraftWriter.ReplaceOwnedSamsung(targets[i], pages[i]);
            return replacement;
        }

        public sealed class OwnedPageReceipt
        {
            public int SlideId { get; set; }
            public int PageOrdinal { get; set; }
            public string SourceId { get; set; }
            public string Owner { get; set; }
            public string Fingerprint { get; set; }
            public string[] Images { get; set; }
        }
        private void ArchiveOwnedPages(IReadOnlyList<PresentationDraftWriter.SamsungOutput> outputs)
        {
            if (_taskContext == null) return;
            foreach (var group in outputs.GroupBy(o => o.Page.Source.Id))
            {
                var ordinal = 0;
                foreach (var output in group)
                {
                    var images = output.Page.Source.ImageData.Select(data =>
                    {
                        var key = "samsung_image:" + TaskCheckpointStore.Fingerprint(data); string id;
                        if (!_taskContext.State.HostData.TryGetValue(key, out id)) _taskContext.State.HostData[key] = id = _taskContext.Store.PutEvidence(_taskContext.State.Id, data);
                        return id;
                    }).ToArray();
                    var receipt = new OwnedPageReceipt { SlideId = (int)((dynamic)output.Slide).SlideID, SourceId = group.Key, PageOrdinal = ordinal++,
                        Owner = output.Owner, Fingerprint = PresentationInspection.Fingerprint(output.Slide), Images = images };
                    _taskContext.State.HostData["samsung_owned:" + receipt.SlideId] = _serializer.Serialize(receipt);
                }
            }
            _taskContext.Checkpoint();
        }
        private List<PresentationDraftWriter.SamsungOutput> LoadOwnedPages(object deck)
        {
            var result = new List<PresentationDraftWriter.SamsungOutput>();
            foreach (var pair in _taskContext.State.HostData.Where(p => p.Key.StartsWith("samsung_owned:")))
            {
                var receipt = _serializer.Deserialize<OwnedPageReceipt>(pair.Value);
                var slide = PresentationInspection.FindSlide(deck, receipt.SlideId);
                if (PresentationInspection.Fingerprint(slide) != receipt.Fingerprint) throw new InvalidOperationException("SLIDE_CHANGED_SINCE_REVIEW: Inspect the user-edited slide before continuing.");
                var raw = _serializer.DeserializeObject(_taskContext.State.HostData["samsung_content:" + receipt.SourceId]);
                var parsed = PresentationDraftWriter.ParseSlides(new object[] { raw });
                parsed[0].ImageData.AddRange(receipt.Images.Select(id => _taskContext.Store.ReadEvidence(_taskContext.State.Id, id)));
                var pages = PresentationDraftWriter.ComposeSamsung(parsed);
                if (receipt.PageOrdinal >= pages.Count) throw new InvalidOperationException("SLIDE_REVIEW_RECEIPT_LAYOUT");
                dynamic presentation = deck; var scale = (float)presentation.PageSetup.SlideWidth / SamsungSlideDesign.Width;
                var page = pages[receipt.PageOrdinal];
                if (Math.Abs(scale - 1) > .001) PresentationDraftWriter.ScaleSamsungPage(page, scale);
                var output = new PresentationDraftWriter.SamsungOutput { Slide = slide, Page = page, Owner = receipt.Owner };
                dynamic native = slide;
                for (var i = 1; i <= (int)native.Shapes.Count; i++) output.ShapeIds.Add((int)native.Shapes[i].Id);
                output.Image = PresentationDraftWriter.ExportSamsung(output); result.Add(output);
            }
            return result.OrderBy(o => (int)((dynamic)o.Slide).SlideIndex).ToList();
        }
        private async Task RepairOwnedGroupAsync(IReadOnlyList<PresentationDraftWriter.SamsungOutput> outputs,
            Dictionary<string, Dictionary<string, object>> content, string id, string findings, string source, string prompt,
            OpenAiCompatibleClient client, AppSettings settings, CancellationToken token, SamsungGenerationJournal journal)
        {
            var group = outputs.Where(o => o.Page.Source.Id == id).ToArray();
            foreach (var output in group)
            {
                var nativeId = (int)((dynamic)output.Slide).SlideID;
                if (_taskContext != null)
                {
                    var key = "samsung_repairs:" + nativeId; string value;
                    var attempts = _taskContext.State.HostData.TryGetValue(key, out value) ? int.Parse(value) : 0;
                    if (attempts >= 3) throw new InvalidOperationException("SLIDE_REPAIR_LIMIT: Three repair cycles already attempted for slide " + nativeId);
                    var defectKey = "samsung_defect:" + nativeId;
                    if (_taskContext.State.HostData.TryGetValue(defectKey, out value) && value == findings) throw new InvalidOperationException("SLIDE_REPAIR_STALLED: " + findings);
                    _taskContext.State.HostData[key] = (attempts + 1).ToString(); _taskContext.State.HostData[defectKey] = findings;
                    _taskContext.State.PresentationReviewReceipt = null; _taskContext.Checkpoint();
                }
            }
            var repaired = await RepairSlideContentAsync(group[0], content[id], findings, source, prompt, client, settings, token, group);
            content[id] = repaired;
            if (_taskContext != null) _taskContext.State.HostData["samsung_content:" + id] = _serializer.Serialize(repaired);
            foreach (var output in group)
            {
                var receipt = journal?.Data.Receipts.SingleOrDefault(r => r.SlideId == (int)((dynamic)output.Slide).SlideID);
                if (receipt != null) journal.Record(output, receipt.Page, _serializer.Serialize(repaired));
            }
            ArchiveOwnedPages(group);
        }
        private async Task ReviewOwnedPagesAsync(IReadOnlyList<PresentationDraftWriter.SamsungOutput> outputs,
            Dictionary<string, Dictionary<string, object>> content, string source, string prompt, OpenAiCompatibleClient client,
            AppSettings settings, CancellationToken token, SamsungGenerationJournal journal, Action<int, int> progress)
        {
            var attempts = new Dictionary<string, int>();
            for (var i = 0; i < outputs.Count; i++)
            {
                token.ThrowIfCancellationRequested(); var output = outputs[i];
                progress?.Invoke(i + 1, outputs.Count);
                var before = PresentationInspection.Fingerprint(output.Slide);
                var review = await ReviewSamsungAsync(client, settings,
                    "Review this rendered Samsung executive slide. Check every item assigned to this page, readable dense evidence, geometry, table/chart labels, clipping, collisions and emphasis. Logical source content may span continuation pages; do not require other pages' items here. Report the provided logical slide_id in findings." + SamsungAuthoringPolicy.ReviewContract,
                    _serializer.Serialize(new { slide_id = output.Page.Source.Id, native_slide_id = (int)((dynamic)output.Slide).SlideID,
                        logical_content = content[output.Page.Source.Id], expected_page = output.Page.Elements.Select(e => new { text = e.Text, table = e.Table == null ? null : new { e.Table.Headers, e.Table.Rows }, chart = e.Chart == null ? null : new { e.Chart.Categories, series = e.Chart.Series.Select(v => new { v.Name, v.Values }) } }), evidence = output.Page.Source.Evidence }), output.Image, token);
                if (PresentationInspection.Fingerprint(output.Slide) != before || PresentationDraftWriter.ExportSamsung(output) != output.Image) throw new InvalidOperationException("SLIDE_CHANGED_DURING_REVIEW");
                if (ReviewApproved(review)) continue;
                var id = output.Page.Source.Id; int count; attempts.TryGetValue(id, out count);
                if (count >= 3) throw new InvalidOperationException("SLIDE_REVIEW_INCOMPLETE: " + review);
                attempts[id] = count + 1;
                await RepairOwnedGroupAsync(outputs, content, id, review, source, prompt, client, settings, token, journal);
                // A repair may reflow all continuations: review every affected page again.
                i = outputs.ToList().FindIndex(o => o.Page.Source.Id == id) - 1;
            }
        }
        private async Task<string> ReviewGeneratedDeckAsync(object deck, string[] plan, string source, string prompt,
            OpenAiCompatibleClient client, AppSettings settings, CancellationToken token, SamsungGenerationJournal journal, Action<int, int> progress)
        {
            var outputs = LoadOwnedPages(deck);
            var content = plan.ToDictionary(id => id, id => _serializer.Deserialize<Dictionary<string, object>>(_taskContext.State.HostData["samsung_content:" + id]));
            for (var cycle = 0; cycle <= 3; cycle++)
            {
                string findings = null;
                for (var offset = 0; offset < outputs.Count; offset += 12)
                {
                    var subset = outputs.Skip(offset).Take(12).ToArray();
                    var visual = await ReviewSamsungAsync(client, settings,
                        "Review consecutive native slides for visual consistency, hierarchy, numbering and Samsung fidelity. Dense evidence is intentional. Report the provided logical slide IDs for affected slides." + SamsungAuthoringPolicy.ReviewContract,
                        _serializer.Serialize(new { prompt, slides = subset.Select(o => new { slide_id = o.Page.Source.Id, native_id = (int)((dynamic)o.Slide).SlideID }) }), SamsungDeckOverview.Montage(subset.Select(o => o.Image)), token);
                    if (!ReviewApproved(visual)) { findings = visual; break; }
                }
                var deckContent = _serializer.Serialize(new { instruction = prompt, plan,
                    briefs = _taskContext.State.HostData.ContainsKey("samsung_briefs") ? _taskContext.State.HostData["samsung_briefs"] : null,
                    slides = plan.Select(id => content[id]) });
                if (findings == null)
                {
                    var verdict = await ReviewSamsungAsync(client, settings, SamsungAuthoringPolicy.DeckReview + SamsungAuthoringPolicy.ReviewContract, deckContent, null, token);
                    if (!ReviewApproved(verdict)) findings = verdict;
                }
                // Prevent a late user edit from inheriting the finished-deck receipt.
                foreach (var output in outputs)
                    if (PresentationInspection.Fingerprint(output.Slide) != _serializer.Deserialize<OwnedPageReceipt>(_taskContext.State.HostData["samsung_owned:" + (int)((dynamic)output.Slide).SlideID]).Fingerprint) throw new InvalidOperationException("SLIDE_CHANGED_DURING_DECK_REVIEW");
                if (findings == null) return SamsungAuthoringPolicy.CacheKey(settings.Model, settings.BaseUrl, deckContent, source);
                if (cycle == 3) throw new InvalidOperationException("SLIDE_DECK_REVIEW: " + findings);
                var report = _serializer.Deserialize<Dictionary<string, object>>(findings);
                var affected = SamsungAuthoringPolicy.Array(report, "findings").Select(SamsungAuthoringPolicy.ReadMap)
                    .Where(f => SamsungAuthoringPolicy.Text(f, "severity") == "blocker").Select(f => SamsungAuthoringPolicy.Text(f, "slide_id")).Distinct().ToArray();
                if (affected.Length == 0 || affected.Any(id => !content.ContainsKey(id))) throw new InvalidOperationException("SLIDE_DECK_REVIEW_TARGET: Review must identify affected logical slide IDs. " + findings);
                foreach (var id in affected) await RepairOwnedGroupAsync(outputs, content, id, findings, source, prompt, client, settings, token, journal);
                await ReviewOwnedPagesAsync(outputs.Where(o => affected.Contains(o.Page.Source.Id)).ToArray(), content, source, prompt, client, settings, token, journal, progress);
                ArchiveOwnedPages(outputs);
            }
            throw new InvalidOperationException("SLIDE_DECK_REVIEW_INCOMPLETE");
        }
    }
}
