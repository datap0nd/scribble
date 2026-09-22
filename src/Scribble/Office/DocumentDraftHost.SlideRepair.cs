using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Scribble.Chat;
using Scribble.Configuration;

namespace Scribble.Office
{
    public sealed partial class DocumentDraftHost
    {
        private static object RepairEvidenceValue(string field, object value)
        {
            var chart = value as IDictionary<string, object>;
            if ((field == "chart" || field == "secondary_chart") && chart != null)
            {
                // A visual repair may make the chart title clearer (for example,
                // add the requested currency unit). The native data contract is
                // still immutable: type, categories, series names and values must
                // remain byte-for-byte equivalent after canonical serialization.
                return chart.Where(pair => pair.Key != "title")
                    .ToDictionary(pair => pair.Key, pair => pair.Value);
            }
            return value;
        }
        internal static void ValidateSlideRepairEvidence(
            IDictionary<string, object> original,
            IDictionary<string, object> replacement)
        {
            foreach (var field in new[] { "table", "secondary_table", "chart", "secondary_chart", "image_names", "source_spans", "evidence", "sources", "calculations", "content_kind" })
            {
                object before, after;
                original.TryGetValue(field, out before);
                replacement.TryGetValue(field, out after);
                if (SamsungRepairPolicy.Serialize(RepairEvidenceValue(field, before)) !=
                    SamsungRepairPolicy.Serialize(RepairEvidenceValue(field, after)))
                    throw new InvalidOperationException("SLIDE_REPAIR_EVIDENCE_CHANGED: " + field);
            }
        }
        internal static void RetainSlideRepairSources(
            IDictionary<string, object> original,
            IDictionary<string, object> replacement)
        {
            foreach (var hostOwned in new[] { "evidence", "source_spans", "sources" })
            {
                object retained;
                if (original.TryGetValue(hostOwned, out retained)) replacement[hostOwned] = retained;
                else replacement.Remove(hostOwned);
            }
        }
        internal static object[] SelectSlideRepair(IDictionary<string, object> wrapper, string expectedId)
        {
            var slides = SamsungAuthoringPolicy.Array(wrapper, "slides");
            if (slides.Length == 1) return slides;
            // Some models return the next planned slide alongside the requested
            // repair. Never apply it here: only the uniquely identified target
            // can pass the schema, evidence and native-fingerprint checks below.
            var matching = slides.Where(value =>
            {
                var map = value as IDictionary<string, object>;
                return map != null && SamsungAuthoringPolicy.Text(map, "id") == expectedId;
            }).ToArray();
            if (matching.Length == 1) return matching;
            throw new InvalidOperationException("SLIDE_REPAIR_COUNT: Repair exactly one identifiable slide.");
        }
        internal static bool CanRetrySlideRepairShape(Exception error, int proposal)
        {
            if (proposal != 0 || !(error is InvalidOperationException)) return false;
            return new[] { "SLIDE_REPAIR_COUNT:", "SLIDE_REPAIR_SCHEMA:", "SLIDE_REPAIR_ID_CHANGED" }
                .Any(code => error.Message.StartsWith(code, StringComparison.Ordinal));
        }
        private async Task<Dictionary<string, object>> RepairSlideContentAsync(PresentationDraftWriter.SamsungOutput output,
            Dictionary<string, object> original, string findings, string source, string prompt,
            OpenAiCompatibleClient client, AppSettings settings, CancellationToken token, IReadOnlyList<PresentationDraftWriter.SamsungOutput> related = null)
        {
            var repairTokens = Math.Min(32768, Math.Max(8192, _serializer.Serialize(original).Length / 2));
            var response = await ReviewSamsungAsync(client, settings,
                SamsungAuthoringPolicy.Instructions + " Repair this single slide using the specific visual findings. Return JSON only: {\"slides\":[{...complete corrected slide...}]}. " +
                "Return only the slide whose id is '" + SamsungAuthoringPolicy.Text(original, "id") + "'; do not return any other planned slide. " +
                "Keep every user-required chart-title unit token (such as EUR); shorten surrounding wording if needed, never remove the unit. " +
                "Keep the ID, all required table rows, chart type, categories, series names and values, calculations and source images unchanged. Do not add or remove a primary or secondary chart or table: this is a visual repair, not new evidence. For a sparse table slide, enlarge the existing table and use semantic highlight_rows and a coherent subtitle/takeaway; do not invent a chart. You may correct an existing chart title when the finding requires it. Omit evidence, source_spans and sources from your answer: the host carries them over unchanged. You may choose a better Samsung layout and remove redundant wording. " +
                "Do not invent pixel coordinates or remove evidence to make it fit. Schema: " + _serializer.Serialize(PresentationToolCatalog.DraftDefinition().function.parameters),
                _serializer.Serialize(new { original, findings, instruction = prompt }), output.Image, token, repairTokens);
            object[] replacements = null;
            Dictionary<string, object> replacement = null;
            for (var proposal = 0; proposal < 2; proposal++)
            {
                var wrapper = await ReadSlideRepairJsonAsync(client, settings, response, token, repairTokens);
                try
                {
                    replacements = SelectSlideRepair(wrapper, SamsungAuthoringPolicy.Text(original, "id"));
                    replacement = SamsungAuthoringPolicy.ReadMap(replacements[0]);
                    var testCall = new ChatToolCall { id = "repair", function = new ChatToolCallFunction { name = PresentationToolCatalog.AddDraftSlides, arguments = _serializer.Serialize(new { slides = replacements }) } };
                    var errors = ToolContractValidator.Validate(testCall, PresentationToolCatalog.DraftDefinition());
                    if (errors.Count > 0) throw new InvalidOperationException("SLIDE_REPAIR_SCHEMA: " + string.Join("; ", errors));
                    if (SamsungAuthoringPolicy.Text(replacement, "id") != SamsungAuthoringPolicy.Text(original, "id")) throw new InvalidOperationException("SLIDE_REPAIR_ID_CHANGED");
                }
                catch (InvalidOperationException ex) when (CanRetrySlideRepairShape(ex, proposal))
                {
                    // Nothing native was changed. Correct a missing, ambiguous
                    // or malformed target inside this receipted repair call,
                    // instead of asking chat to replay the written deck.
                    response = await ReviewSamsungAsync(client, settings,
                        SamsungAuthoringPolicy.Instructions + " Your repair was rejected (" + ex.Message + "). " +
                        "Return JSON only: a slides array containing exactly one complete slide with id '" +
                        SamsungAuthoringPolicy.Text(original, "id") + "'. Do not include another planned slide. " +
                        "Preserve the original evidence, chart/table data, required facts and source image names. " +
                        "Omit evidence, source_spans and sources because the host retains them. Schema: " +
                        _serializer.Serialize(PresentationToolCatalog.DraftDefinition().function.parameters),
                        _serializer.Serialize(new { original, findings, previous_invalid_response = response, instruction = prompt }),
                        output.Image, token, repairTokens);
                    continue;
                }
                // Evidence, span IDs and the visible citation belong to the host.
                RetainSlideRepairSources(original, replacement);
                try { ValidateSlideRepairEvidence(original, replacement); }
                catch (InvalidOperationException ex) when (ex.Message.StartsWith("SLIDE_REPAIR_EVIDENCE_CHANGED:", StringComparison.Ordinal) && proposal == 0)
                {
                    // The candidate has not touched the native slide. Give the
                    // reviewer one bounded correction inside this same write
                    // receipt instead of stranding a generated deck when it
                    // invents a chart for a methodology card, for example.
                    response = await ReviewSamsungAsync(client, settings,
                        SamsungAuthoringPolicy.Instructions + " Your visual repair changed immutable source evidence (" + ex.Message + "). " +
                        "Return exactly one corrected slide with the original ID. Preserve table, secondary_table, chart, secondary_chart, image_names, calculations and content_kind exactly as in original; if a field is absent there, omit it here. " +
                        "Change only layout or explanatory wording to address the visual findings. Omit evidence, source_spans and sources; the host retains them. Return JSON only. Schema: " +
                        _serializer.Serialize(PresentationToolCatalog.DraftDefinition().function.parameters),
                        _serializer.Serialize(new { original, findings, previous_invalid = replacement, instruction = prompt }),
                        output.Image, token, repairTokens);
                    continue;
                }
                if (SamsungRepairPolicy.Serialize(original) != SamsungRepairPolicy.Serialize(replacement)) break;
                if (proposal == 1) throw new InvalidOperationException("SLIDE_REPAIR_STALLED: Two proposals made no meaningful content or layout change.");
                // An unchanged proposal has not touched PowerPoint. Give the
                // reviewer one precise opportunity to fix its own no-op inside
                // this receipted host call, rather than stranding the written
                // slide and relying on the chat model to replay a write.
                response = await ReviewSamsungAsync(client, settings,
                    SamsungAuthoringPolicy.Instructions + " The previous visual repair returned the original slide unchanged. Make one concrete content or Samsung-layout change that addresses the findings, while preserving its ID, source evidence, required facts, table/chart data and image names. Return JSON only with exactly this one complete slide. Omit evidence, source_spans and sources because the host retains them. Schema: " + _serializer.Serialize(PresentationToolCatalog.DraftDefinition().function.parameters),
                    _serializer.Serialize(new { original, findings, previous_no_op = replacement, instruction = prompt }),
                    output.Image, token, repairTokens);
            }
            SamsungPresentationReview.ValidateEvidence(_serializer.Serialize(replacement), source);
            var review = await ReviewSamsungAsync(client, settings, SamsungAuthoringPolicy.FactReview +
                " Verify that every required point from the original ONE slide remains represented. Reject omitted commitments, qualifications, owners or dates on that slide. Do not assess whether another planned slide has been added yet; this is a single-slide repair, not a deck review. Report only the original slide ID in findings." + SamsungAuthoringPolicy.ReviewContract,
                _serializer.Serialize(new { original, replacement, source, instruction = prompt }), null, token);
            if (!ReviewApproved(review) &&
                !SamsungAuthoringPolicy.OnlyOtherSlideCoverageBlockers(review,
                    SamsungAuthoringPolicy.Text(original, "id")))
                throw new InvalidOperationException("SLIDE_REPAIR_FACTS: " + review);
            var parsed = PresentationDraftWriter.ParseSlides(replacements);
            ValidatePromptChartConstraints(prompt, parsed);
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

        private async Task<Dictionary<string, object>> ReadSlideRepairJsonAsync(
            OpenAiCompatibleClient client,
            AppSettings settings,
            string response,
            CancellationToken token,
            int maxTokens)
        {
            Exception lastError = null;
            var candidate = response;
            for (var attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    var parsed = _serializer.Deserialize<Dictionary<string, object>>(candidate);
                    if (parsed == null) throw new InvalidOperationException("The slide repair JSON was null.");
                    return parsed;
                }
                catch (Exception exception) when (exception is ArgumentException || exception is InvalidOperationException)
                {
                    lastError = exception;
                    if (attempt == 2) break;
                }

                // A malformed JSON repair response is not a new Office action.
                // Correct its transport syntax inside this host call so the chat
                // loop does not replay the already-written slide or spend one of
                // its repeated-action recovery attempts. Schema, immutable data,
                // source evidence and native fingerprints are all checked below
                // before any repaired content can replace the owned slide.
                candidate = await ReviewSamsungAsync(client, settings,
                    "Repair JSON syntax only. Return exactly one valid JSON object with a top-level slides array containing the same one complete slide. Preserve every value, array item and object field from the attempted JSON. Do not summarize, explain, add facts or remove content. Escape quotes inside strings and output JSON only. Expected schema: " +
                    _serializer.Serialize(PresentationToolCatalog.DraftDefinition().function.parameters),
                    _serializer.Serialize(new
                    {
                        parsing_error = Scribble.Security.TextBoundary.SingleLine(lastError.Message, 400),
                        attempted_json = candidate
                    }), null, token, maxTokens);
            }
            throw new InvalidOperationException("SLIDE_REPAIR_JSON_INVALID: The reviewer returned malformed slide JSON after two syntax-only retries.", lastError);
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
                // Each logical slide is composed independently, so its design
                // model initially says "- 1 -" even when the owned native slide
                // is later in the deck. Keep the review model aligned with the
                // number already written to PowerPoint; otherwise the deck
                // reviewer invents a numbering defect and needlessly rewrites
                // an otherwise approved slide.
                page.PageNumber.Text = "- " + (int)((dynamic)slide).SlideIndex + " -";
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
        private static bool NativePageNumberMatches(PresentationDraftWriter.SamsungOutput output)
        {
            try
            {
                dynamic slide = output.Slide;
                var expected = "- " + (int)slide.SlideIndex + " -";
                if (output.Page.PageNumber.Text != expected) return false;
                for (var index = 1; index <= (int)slide.Shapes.Count; index++)
                {
                    dynamic shape = slide.Shapes[index];
                    if ((int)shape.HasTextFrame != 0 &&
                        string.Equals((Convert.ToString(shape.TextFrame.TextRange.Text) ?? "").Trim(), expected,
                            StringComparison.Ordinal)) return true;
                }
            }
            catch { /* An unreadable native footer cannot bypass visual review. */ }
            return false;
        }
        internal static string FilterReviewFindings(string review,
            Func<IDictionary<string, object>, bool> refuted)
        {
            try
            {
                var json = new System.Web.Script.Serialization.JavaScriptSerializer { MaxJsonLength = int.MaxValue };
                var report = json.Deserialize<Dictionary<string, object>>(review);
                var findings = SamsungAuthoringPolicy.Array(report, "findings")
                    .Select(SamsungAuthoringPolicy.ReadMap).ToArray();
                if (findings.Length == 0) return review;
                var remaining = findings.Where(finding => !refuted(finding)).ToArray();
                if (remaining.Length == findings.Length) return review;
                report["findings"] = remaining;
                report["approved"] = !remaining.Any(finding =>
                    string.Equals(SamsungAuthoringPolicy.Text(finding, "severity"), "blocker",
                        StringComparison.OrdinalIgnoreCase));
                report["issues"] = string.Join("; ", remaining.Select(finding => SamsungAuthoringPolicy.Text(finding, "correction")));
                return json.Serialize(report);
            }
            catch (Exception)
            {
                // Malformed review output never bypasses inspection.
                return review;
            }
        }
        internal static string FilterBriefRefutedReview(string review, string brief, string slide)
        {
            if (!Regex.IsMatch(review ?? "", @"\bbrief\b.{0,80}\b(?:require|specif|demand)",
                RegexOptions.IgnoreCase | RegexOptions.Singleline)) return review;
            var slideNumbers = new HashSet<string>(Regex.Matches(slide ?? "", @"(?<!\d)\d[\d,.]*%?")
                .Cast<Match>().Select(match => match.Value), StringComparer.OrdinalIgnoreCase);
            var staleBriefNumbers = new HashSet<string>(Regex.Matches(brief ?? "", @"(?<!\d)\d[\d,.]*%?")
                .Cast<Match>().Select(match => match.Value).Where(value => !slideNumbers.Contains(value)),
                StringComparer.OrdinalIgnoreCase);
            if (staleBriefNumbers.Count == 0) return review;
            return FilterReviewFindings(review, finding =>
                string.Equals(SamsungAuthoringPolicy.Text(finding, "type"), "facts", StringComparison.OrdinalIgnoreCase) &&
                staleBriefNumbers.Any(value => SamsungAuthoringPolicy.Text(finding, "correction")
                    .IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0));
        }
        private static string FilterNativeRefutedReview(string review,
            IReadOnlyList<PresentationDraftWriter.SamsungOutput> outputs)
        {
            return FilterReviewFindings(review, finding =>
            {
                var type = SamsungAuthoringPolicy.Text(finding, "type");
                if (string.Equals(type, "facts", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(type, "coverage", StringComparison.OrdinalIgnoreCase)) return false;
                var id = SamsungAuthoringPolicy.Text(finding, "slide_id");
                var output = outputs.FirstOrDefault(item => item.Page.Source.Id == id) ??
                    (outputs.Count == 1 ? outputs[0] : null);
                if (output == null) return false;
                var correction = SamsungAuthoringPolicy.Text(finding, "correction");
                if (Regex.IsMatch(correction, @"\bspeaker\s+notes?\b", RegexOptions.IgnoreCase) &&
                    Regex.IsMatch(correction, @"\b(?:overlap|collision|move|resize)\b", RegexOptions.IgnoreCase))
                    return true; // Speaker notes are not slide-canvas objects.
                if (correction.IndexOf(PresentationDraftWriter.DraftMarker, StringComparison.OrdinalIgnoreCase) >= 0 &&
                    NativeTextExists(output, PresentationDraftWriter.DraftMarker)) return true;
                if (Regex.IsMatch(correction, @"\b(?:zero[- ]based|zero\s+tick|0\s+tick|starts?\s+(?:at|from)\s+0)\b", RegexOptions.IgnoreCase) &&
                    NativeZeroBaseline(output)) return true;
                if (Regex.IsMatch(correction, @"\b(?:source|citation|footer|footnote)\b.{0,85}\b(?:cut off|clipp\w*|truncat\w*)\b", RegexOptions.IgnoreCase) &&
                    NativeFooterFits(output)) return true;
                if (Regex.IsMatch(correction, @"\b(?:takeaway|banner)\b.{0,100}\b(?:missing|represented|present)\b", RegexOptions.IgnoreCase) &&
                    NativeTextExists(output, SamsungAuthoringPolicy.AudienceTakeaway(output.Page.Source.Takeaway))) return true;
                if (Regex.IsMatch(correction, @"\b(?:cards?|containers?|headings?)\b", RegexOptions.IgnoreCase) &&
                    Regex.IsMatch(correction, @"\b(?:single|generic|missing|absent|implement|add|replac\w*)\b", RegexOptions.IgnoreCase) &&
                    NativeCardHeadingsPresent(output)) return true;
                return false;
            });
        }
        private static bool NativeTextExists(PresentationDraftWriter.SamsungOutput output, string wanted)
        {
            if (string.IsNullOrWhiteSpace(wanted)) return false;
            try
            {
                dynamic slide = output.Slide;
                for (var index = 1; index <= (int)slide.Shapes.Count; index++)
                {
                    dynamic shape = slide.Shapes[index];
                    if ((int)shape.HasTextFrame == 0) continue;
                    if (string.Equals((Convert.ToString(shape.TextFrame.TextRange.Text) ?? "").Trim(), wanted.Trim(),
                        StringComparison.OrdinalIgnoreCase)) return true;
                }
            }
            catch { /* An unreadable native slide cannot refute a finding. */ }
            return false;
        }
        private static bool NativeCardHeadingsPresent(PresentationDraftWriter.SamsungOutput output)
        {
            return output.Page.Source.Layout == "cards" && output.Page.Source.Cards.Count >= 2 &&
                output.Page.Source.Cards.Select(card => card.Heading).Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count() == output.Page.Source.Cards.Count &&
                output.Page.Source.Cards.All(card => NativeTextExists(output, card.Heading));
        }
        private static bool NativeZeroBaseline(PresentationDraftWriter.SamsungOutput output)
        {
            try
            {
                dynamic slide = output.Slide; var count = 0;
                for (var index = 1; index <= (int)slide.Shapes.Count; index++)
                {
                    dynamic shape = slide.Shapes[index];
                    if ((int)shape.HasChart == 0) continue;
                    count++;
                    if (Math.Abs(Convert.ToDouble(shape.Chart.Axes(2).MinimumScale)) > .001) return false;
                }
                return count > 0;
            }
            catch { return false; }
        }
        private static bool NativeFooterFits(PresentationDraftWriter.SamsungOutput output)
        {
            try
            {
                dynamic slide = output.Slide; var found = false;
                var height = (double)slide.Parent.PageSetup.SlideHeight;
                var width = (double)slide.Parent.PageSetup.SlideWidth;
                for (var index = 1; index <= (int)slide.Shapes.Count; index++)
                {
                    dynamic shape = slide.Shapes[index];
                    if ((int)shape.HasTextFrame == 0 || (double)shape.Top < height * .90) continue;
                    dynamic range = shape.TextFrame.TextRange;
                    var text = (Convert.ToString(range.Text) ?? "").Trim();
                    if (text.Length < 20 || text == PresentationDraftWriter.DraftMarker) continue;
                    found = true;
                    if ((double)shape.Left < -1 || (double)shape.Left + (double)shape.Width > width + 1 ||
                        (double)shape.Top + (double)shape.Height > height + 1 ||
                        (double)range.BoundHeight > (double)shape.Height + 1 ||
                        (double)range.BoundWidth > (double)shape.Width + 1) return false;
                }
                return found;
            }
            catch { return false; }
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
                    "Review this rendered Samsung executive slide as an audience would see it at thumbnail size. Check every item assigned to this page, geometry, table/chart labels, visible clipping, collisions, focal hierarchy, balanced canvas use and meaningful visual storytelling. Reject a plain multiline data dump or Word-page composition. A takeaway may recap the topic when it adds a distinct sourced fact; only exact redundant full-text repetition is a blocker. White space framing a substantial native chart or table is intentional, not an accidental empty canvas. The native text-fit check has already verified each textbox; do not infer clipping from its length without a visible cut-off. The host sets the page number from the actual native slide index: flag numbering only if the visible number differs from the supplied expected_page element, not from a logical-slide ordinal. Logical source content may span continuation pages; do not require other pages' items or any other planned slide here. Report the provided logical slide_id in findings; a chart planned for another slide is not missing from this one." + SamsungAuthoringPolicy.ReviewContract,
                    _serializer.Serialize(new { slide_id = output.Page.Source.Id, native_slide_id = (int)((dynamic)output.Slide).SlideID,
                        logical_content = content[output.Page.Source.Id], expected_page = output.Page.Elements.Select(e => new { text = e.Text, table = e.Table == null ? null : new { e.Table.Headers, e.Table.Rows }, chart = e.Chart == null ? null : new { title = e.Chart.Title, type = e.Chart.TypeCode, e.Chart.Categories, series = e.Chart.Series.Select(v => new { v.Name, v.Values }) } }), evidence = output.Page.Source.Evidence }), output.Image, token);
                if (PresentationInspection.Fingerprint(output.Slide) != before || PresentationDraftWriter.ExportSamsung(output) != output.Image) throw new InvalidOperationException("SLIDE_CHANGED_DURING_REVIEW");
                review = FilterNativeRefutedReview(review, new[] { output });
                if (ReviewApproved(review) ||
                    (NativePageNumberMatches(output) && SamsungAuthoringPolicy.OnlyHostOwnedPageNumberBlockers(review)) ||
                    SamsungAuthoringPolicy.OnlyOtherSlideCoverageBlockers(review, output.Page.Source.Id))
                    continue;
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
            SamsungAuthoringPolicy.ValidateDeckVisualDesign(content.Values);
            for (var cycle = 0; cycle <= 3; cycle++)
            {
                string findings = null;
                for (var offset = 0; offset < outputs.Count; offset += 12)
                {
                    var subset = outputs.Skip(offset).Take(12).ToArray();
                    var visual = await ReviewSamsungAsync(client, settings,
                        "Review consecutive native slides as an executive audience would see them at thumbnail size. Check visual consistency, focal hierarchy, balanced use of the canvas, meaningful visual storytelling and Samsung fidelity. Reject slides that resemble a Word page pasted onto a canvas or rely on a plain multiline data dump. A takeaway that adds a distinct sourced fact is not a repeated-conclusion defect. White space framing a substantial native chart or table is intentional. The page number is host-owned and follows actual native slide order; flag it only when the visible number differs from its supplied expected_page element. Report the provided logical slide IDs for affected slides." + SamsungAuthoringPolicy.ReviewContract,
                        _serializer.Serialize(new { prompt, slides = subset.Select(o => new { slide_id = o.Page.Source.Id, native_id = (int)((dynamic)o.Slide).SlideID,
                            expected_page = o.Page.Elements.Select(e => new { text = e.Text, table = e.Table == null ? null : new { e.Table.Headers, e.Table.Rows }, chart = e.Chart == null ? null : new { title = e.Chart.Title, type = e.Chart.TypeCode, e.Chart.Categories, series = e.Chart.Series.Select(v => new { v.Name, v.Values }) } }) }) }), SamsungDeckOverview.Montage(subset.Select(o => o.Image)), token);
                    visual = FilterNativeRefutedReview(visual, subset);
                    if (!ReviewApproved(visual) &&
                        !(subset.All(NativePageNumberMatches) && SamsungAuthoringPolicy.OnlyHostOwnedPageNumberBlockers(visual)))
                    { findings = visual; break; }
                }
                var deckContent = _serializer.Serialize(new { instruction = prompt, plan,
                    briefs = _taskContext.State.HostData.ContainsKey("samsung_briefs") ? _taskContext.State.HostData["samsung_briefs"] : null,
                    slides = plan.Select(id => content[id]) });
                if (findings == null)
                {
                    var verdict = await ReviewSamsungAsync(client, settings, SamsungAuthoringPolicy.DeckReview + SamsungAuthoringPolicy.ReviewContract, deckContent, null, token);
                    verdict = FilterNativeRefutedReview(verdict, outputs);
                    if (!ReviewApproved(verdict) &&
                        !(outputs.All(NativePageNumberMatches) && SamsungAuthoringPolicy.OnlyHostOwnedPageNumberBlockers(verdict)))
                        findings = verdict;
                }
                // Prevent a late user edit from inheriting the finished-deck receipt.
                foreach (var output in outputs)
                    if (PresentationInspection.Fingerprint(output.Slide) != _serializer.Deserialize<OwnedPageReceipt>(_taskContext.State.HostData["samsung_owned:" + (int)((dynamic)output.Slide).SlideID]).Fingerprint) throw new InvalidOperationException("SLIDE_CHANGED_DURING_DECK_REVIEW");
                if (findings == null) return SamsungAuthoringPolicy.CacheKey(settings.Model, settings.BaseUrl, deckContent, source);
                if (cycle == 3) throw new InvalidOperationException("SLIDE_DECK_REVIEW: " + findings);
                var report = _serializer.Deserialize<Dictionary<string, object>>(findings);
                var affected = AffectedDeckReviewSlides(report);
                if (affected.Length == 0 || affected.Any(id => !content.ContainsKey(id))) throw new InvalidOperationException("SLIDE_DECK_REVIEW_TARGET: Review must identify affected logical slide IDs. " + findings);
                foreach (var id in affected) await RepairOwnedGroupAsync(outputs, content, id, findings, source, prompt, client, settings, token, journal);
                await ReviewOwnedPagesAsync(outputs.Where(o => affected.Contains(o.Page.Source.Id)).ToArray(), content, source, prompt, client, settings, token, journal, progress);
                ArchiveOwnedPages(outputs);
            }
            throw new InvalidOperationException("SLIDE_DECK_REVIEW_INCOMPLETE");
        }

        internal static string[] AffectedDeckReviewSlides(IDictionary<string, object> report)
        {
            var findings = SamsungAuthoringPolicy.Array(report, "findings")
                .Select(SamsungAuthoringPolicy.ReadMap).ToArray();
            var blockers = findings.Where(f => SamsungAuthoringPolicy.Text(f, "severity") == "blocker").ToArray();
            // An independent reviewer can reject a deck with actionable
            // warnings only. Repair those instead of reporting a missing
            // target and abandoning an otherwise complete native deck.
            var actionable = blockers.Length > 0 ? blockers : findings.Where(f =>
                SamsungAuthoringPolicy.Text(f, "severity") == "warning").ToArray();
            return actionable.Select(f => SamsungAuthoringPolicy.Text(f, "slide_id"))
                .Where(id => !string.IsNullOrWhiteSpace(id)).Distinct().ToArray();
        }
    }
}
