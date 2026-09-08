using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Scribble.Chat;
using Scribble.Configuration;
using Scribble.Security;

namespace Scribble.Office
{
    public sealed partial class DocumentDraftHost
    {
        private async Task<MailboxToolResult> ExecuteRevisionAsync(ChatToolCall call, OneShotDraftAuthorization authorization,
            bool exclusive, string prompt, OpenAiCompatibleClient client, AppSettings settings, CancellationToken token, Action<int, int> progress)
        {
            PresentationRevision revision = null;
            var written = false;
            try
            {
                if (!PresentationRevisionAcceptance.Enabled) throw new InvalidOperationException("REVISION_NATIVE_ACCEPTANCE_REQUIRED: This build has not passed native preservation and rollback acceptance.");
                if (_hostKind != "powerpoint" || !exclusive || authorization == null || !authorization.CanCreate)
                    throw new InvalidOperationException("REVISION_NOT_AUTHORIZED: Use the PowerPoint pane and an explicit editing request.");
                if (!ModelCatalog.IsVisionCapable(settings.Model)) throw new InvalidOperationException("SLIDE_VISION_REQUIRED");
                dynamic app = _hostApplication; object deck = app.ActivePresentation;
                var args = ToolArguments.Parse(_serializer, call.function.arguments);
                if (SamsungAuthoringPolicy.Text(args, "presentation_id") != PresentationInspection.IdentityFor(deck))
                    throw new InvalidOperationException("REVISION_PRESENTATION_CHANGED: Inspect the original presentation again.");
                if (call.function.name == PresentationToolCatalog.RevertSlides)
                {
                    if (!Regex.IsMatch(prompt ?? "", @"\b(revert|undo|restore)\b", RegexOptions.IgnoreCase)) throw new InvalidOperationException("REVISION_REVERT_NOT_REQUESTED");
                    var resumingRevert = _taskContext != null && PresentationRevision.CanResume(_taskContext.State, call);
                    revision = resumingRevert ? PresentationRevision.Recover(_hostApplication, deck, _taskContext.State.HostData["powerpoint_revision_snapshot"]) :
                        PresentationRevision.Last(deck) ?? throw new InvalidOperationException("REVISION_SESSION_ENDED: No recoverable revision in this Office session.");
                    if (!authorization.TryConsume()) throw new InvalidOperationException("REVISION_PERMISSION_UNAVAILABLE");
                    written = true;
                    if (resumingRevert)
                    {
                        var result = revision.Reconcile();
                        if (result == "applied") revision.RevertWithJournal(status => CheckpointRevision(call, revision, status));
                    }
                    else revision.RevertWithJournal(status => CheckpointRevision(call, revision, status));
                    if (_taskContext != null)
                        foreach (var pending in _taskContext.State.Writes.Where(w => w.BeforeFingerprint == TaskCheckpointStore.Fingerprint(_serializer.Serialize(call.function))))
                        { pending.Status = "verified"; pending.AfterFingerprint = "revert_reconciled"; }
                    authorization.MarkCreated();
                    if (_taskContext != null) { _taskContext.State.PresentationReviewReceipt = null; _taskContext.State.PresentationReviewRequired = false; _taskContext.Checkpoint(); }
                    return new MailboxToolResult(call.id, _serializer.Serialize(new { ok = true, saved = false, status = revision.Status }), "Reverted Scribble changes. Nothing was saved.");
                }
                if (_taskContext != null && PresentationRevision.CanResume(_taskContext.State, call))
                {
                    revision = PresentationRevision.Recover(_hostApplication, deck, _taskContext.State.HostData["powerpoint_revision_snapshot"]);
                    var reconciled = revision.Reconcile();
                    if (reconciled == "applied")
                    {
                        var live = new List<object>(); var images = new List<string>();
                        dynamic recoveredDeck = deck;
                        for (var i = 1; i <= (int)recoveredDeck.Slides.Count; i++)
                        { live.Add(PresentationInspection.Capture((object)recoveredDeck.Slides[i])); images.Add(PresentationInspection.Preview((object)recoveredDeck.Slides[i])); }
                        var evidence = SamsungPresentationReview.SourceCorpus(_taskContext, prompt);
                        var liveContent = _serializer.Serialize(new { instruction = prompt, evidence, slides = live });
                        for (var i = 0; i < images.Count; i += 12)
                        {
                            var visual = await ReviewSamsungAsync(client, settings, SamsungAuthoringPolicy.DeckReview + SamsungAuthoringPolicy.ReviewContract,
                                "Verify surviving slides from a completed interrupted revision. " + liveContent, SamsungDeckOverview.Montage(images.Skip(i)), token);
                            if (!ReviewApproved(visual)) throw new InvalidOperationException("REVISION_RECOVERY_REVIEW: " + visual);
                        }
                        var facts = await ReviewSamsungAsync(client, settings, SamsungAuthoringPolicy.FactReview + SamsungAuthoringPolicy.DeckReview + SamsungAuthoringPolicy.ReviewContract, liveContent, null, token);
                        if (!ReviewApproved(facts)) throw new InvalidOperationException("REVISION_RECOVERY_REVIEW: " + facts);
                        revision.Reconcile(); // Detect a user edit while the model reviewed.
                        _taskContext.State.PresentationReviewReceipt = SamsungAuthoringPolicy.CacheKey(settings.Model, settings.BaseUrl, liveContent, evidence);
                        authorization.MarkCreated();
                    }
                    _taskContext.State.HostData["powerpoint_revision_snapshot"] = revision.Snapshot();
                    foreach (var pending in _taskContext.State.Writes.Where(w => w.BeforeFingerprint == TaskCheckpointStore.Fingerprint(_serializer.Serialize(call.function))))
                    { pending.Status = "verified"; pending.AfterFingerprint = "native_revision_reconciled:" + reconciled; }
                    _taskContext.Checkpoint();
                    if (reconciled == "applied") return new MailboxToolResult(call.id, _serializer.Serialize(new { ok = true, saved = false, status = reconciled, replayed = false }), "The surviving revision was verified and reviewed. No edit was repeated.");
                    _taskContext.State.HostData.Remove("powerpoint_revision_snapshot"); _taskContext.Checkpoint();
                    return new MailboxToolResult(call.id, _serializer.Serialize(new { error_code = "REVISION_RECOVERED_ORIGINALS", saved = false, permission_consumed = false, status = reconciled }), "The interrupted batch was rolled back. Inspect the restored IDs before retrying the requested revision.");
                }
                var operations = SamsungAuthoringPolicy.Array(args, "operations");
                foreach (var raw in operations)
                {
                    var operation = SamsungAuthoringPolicy.ReadMap(raw); var kind = SamsungAuthoringPolicy.Text(operation, "kind");
                    if (kind == "delete" && !Regex.IsMatch(prompt ?? "", @"\b(delete|remove)\b", RegexOptions.IgnoreCase)) throw new InvalidOperationException("SLIDE_DELETE_NOT_REQUESTED");
                    if (kind == "insert" && !Regex.IsMatch(prompt ?? "", @"\b(add|insert|create|expand)\b", RegexOptions.IgnoreCase)) throw new InvalidOperationException("SLIDE_INSERT_NOT_REQUESTED");
                    if (kind == "replace_slide" && !Regex.IsMatch(prompt ?? "", @"\b(redesign|reformat|restructure|recompose|layout|improve)\b", RegexOptions.IgnoreCase)) throw new InvalidOperationException("SLIDE_REDESIGN_NOT_REQUESTED");
                    if (kind == "move" && !Regex.IsMatch(prompt ?? "", @"\b(move|reorder|reorganize|sort)\b", RegexOptions.IgnoreCase)) throw new InvalidOperationException("SLIDE_REORDER_NOT_REQUESTED");
                }
                var source = SamsungPresentationReview.SourceCorpus(_taskContext, prompt);
                foreach (var raw in operations.Select(SamsungAuthoringPolicy.ReadMap))
                {
                    object supplied;
                    if (!raw.TryGetValue("slide", out supplied)) continue;
                    var content = SamsungAuthoringPolicy.ReadMap(supplied);
                    var spans = SamsungAuthoringPolicy.Array(content, "source_spans");
                    if (spans.Length > 0)
                    {
                        if (_taskContext == null) throw new InvalidOperationException("SLIDE_SOURCE_REF_INVALID");
                        content["evidence"] = _taskContext.Sources.Resolve(spans.Select(Convert.ToString));
                    }
                    SamsungPresentationReview.ValidateEvidence(_serializer.Serialize(content), source);
                }
                var original = operations.Select(SamsungAuthoringPolicy.ReadMap).Select(o => Convert.ToInt32(o["slide_id"])).Distinct()
                    .Select(id => PresentationInspection.Capture(PresentationInspection.FindSlide(deck, id))).ToArray();
                var requestedOperations = operations;
                var defects = new HashSet<string>();
                string deckContent = null;
                for (var cycle = 0; cycle <= 3; cycle++)
                {
                    try
                    {
                        var factReview = await ReviewSamsungAsync(client, settings,
                            SamsungAuthoringPolicy.FactReview + " Check the exact requested scope: reject changes to unrelated slides or objects. Existing deck content is reference data, not independently established fact. Require supplied evidence for new claims; requested stylistic edits need no invented external source." + SamsungAuthoringPolicy.ReviewContract,
                            _serializer.Serialize(new { instruction = prompt, source, original, requestedOperations, operations }), null, token);
                        if (!ReviewApproved(factReview)) throw new InvalidOperationException("REVISION_SOURCE_REVIEW: " + factReview);
                        revision = new PresentationRevision(deck);
                        revision.Stage(_hostApplication, operations);
                        for (var i = 0; i < revision.Items.Count; i++)
                        {
                            token.ThrowIfCancellationRequested(); progress?.Invoke(i + 1, revision.Items.Count);
                            var item = revision.Items[i];
                            foreach (var previewSlide in new[] { item.Staged }.Concat(item.StagedInserts))
                            {
                            var verdict = await ReviewSamsungAsync(client, settings,
                                "Review the proposed Samsung slide revision. Check requested edits, preservation of unrelated content, readable dense evidence, chart/table labels, collisions, clipping and emphasis." + SamsungAuthoringPolicy.ReviewContract,
                                _serializer.Serialize(new { instruction = prompt, source, expected = item.Operations, original = PresentationInspection.Capture(item.Original), proposed = PresentationInspection.Capture(previewSlide) }),
                                PresentationInspection.Preview(previewSlide), token);
                            if (!ReviewApproved(verdict)) throw new InvalidOperationException("REVISION_VISUAL_REVIEW: " + verdict);
                            }
                        }
                        var all = revision.ReviewedSlides().Select((slide, index) =>
                        {
                            var capture = PresentationInspection.Capture(slide);
                            capture["index"] = index + 1;
                            var owner = revision.Items.FirstOrDefault(item => ReferenceEquals(item.Staged, slide));
                            if (owner != null) capture["slide_id"] = owner.SlideId;
                            else
                            {
                                var insertion = revision.Items.FirstOrDefault(item => item.StagedInserts.Contains(slide));
                                if (insertion != null) capture["slide_id"] = "pending-insert-after-" + insertion.SlideId + "-" + insertion.StagedInserts.IndexOf(slide);
                            }
                            return capture;
                        }).ToArray();
                        deckContent = _serializer.Serialize(new { instruction = prompt, proposed_slides = all, operations });
                        var deckReview = await ReviewSamsungAsync(client, settings, SamsungAuthoringPolicy.DeckReview + SamsungAuthoringPolicy.ReviewContract, deckContent, null, token);
                        if (!ReviewApproved(deckReview)) throw new InvalidOperationException("REVISION_DECK_REVIEW: " + deckReview);
                        break;
                    }
                    catch (InvalidOperationException failure) when (failure.Message.StartsWith("REVISION_VISUAL_REVIEW:") ||
                        failure.Message.StartsWith("REVISION_DECK_REVIEW:") || failure.Message.StartsWith("SLIDE_NATIVE_OVERFLOW:"))
                    {
                        revision?.CloseStaging(false); revision = null;
                        if (cycle == 3 || !defects.Add(failure.Message)) throw new InvalidOperationException("REVISION_REPAIR_STOPPED: " + failure.Message);
                        var repaired = await ReviewSamsungAsync(client, settings, SamsungAuthoringPolicy.Instructions +
                            "Repair this proposed revision batch using the specific findings. Return JSON only: {\"operations\":[...]}. " +
                            "Preserve operation count/order, target IDs, fingerprints, before values and requested numeric changes. " +
                            "You may rewrite redundant replacement text or improve the recipe of an already-authorized replacement slide. " +
                            "Do not add operations, expand scope, change evidence, or change chart/table values. " +
                            "If the defect requires broader scope, return the unchanged operations so the host can explain the conflict.",
                            _serializer.Serialize(new { instruction = prompt, source, original, requestedOperations, operations, findings = failure.Message }), null, token, 16384);
                        var corrected = SamsungAuthoringPolicy.Array(_serializer.Deserialize<Dictionary<string, object>>(repaired), "operations");
                        SamsungRepairPolicy.ValidateScope(requestedOperations, corrected);
                        if (SamsungRepairPolicy.Serialize(operations) == SamsungRepairPolicy.Serialize(corrected))
                            throw new InvalidOperationException("REVISION_REPAIR_SCOPE_CONFLICT: Resolving these findings requires a different edit scope: " + failure.Message);
                        var repairCall = new ChatToolCall { id = call.id, function = new ChatToolCallFunction { name = call.function.name,
                            arguments = _serializer.Serialize(new { presentation_id = SamsungAuthoringPolicy.Text(args, "presentation_id"), operations = corrected }) } };
                        var definition = PresentationToolCatalog.RevisionDefinitions().Single(t => t.function.name == PresentationToolCatalog.ReviseSlides);
                        var errors = ToolContractValidator.Validate(repairCall, definition);
                        if (errors.Count != 0) throw new InvalidOperationException("REVISION_REPAIR_SCHEMA: " + string.Join("; ", errors));
                        operations = corrected;
                    }
                }
                token.ThrowIfCancellationRequested();
                if (!authorization.TryConsume() && (_taskContext == null || !_taskContext.State.HostData.ContainsKey("powerpoint_revision_authorized")))
                    throw new InvalidOperationException("REVISION_PERMISSION_UNAVAILABLE");
                if (_taskContext != null) { _taskContext.State.HostData["powerpoint_revision_authorized"] = "true"; _taskContext.Checkpoint(); }
                ((dynamic)deck).Tags.Add("ScribblePresentationId", PresentationInspection.IdentityFor(deck));
                if (_taskContext != null) _taskContext.State.HostData["powerpoint_revision_payload"] = _taskContext.Store.PutEvidence(_taskContext.State.Id, call.function.arguments);
                written = true;
                revision.Commit(status => CheckpointRevision(call, revision, status));
                authorization.MarkCreated();
                if (_taskContext != null)
                {
                    _taskContext.State.PresentationReviewRequired = true;
                    _taskContext.State.PresentationReviewReceipt = SamsungAuthoringPolicy.CacheKey(settings.Model, settings.BaseUrl, deckContent, source);
                    _taskContext.Checkpoint();
                }
                return new MailboxToolResult(call.id, _serializer.Serialize(new { ok = true, saved = false, revised_slides = revision.Items.Count, revert_available = true }),
                    "Revised " + revision.Items.Count + " slides. Nothing was saved. Revert Scribble changes is available in this Office session.");
            }
            catch (OperationCanceledException)
            {
                if (!written && _taskContext != null)
                {
                    var pending = _taskContext.State.Writes.FirstOrDefault(w => w.Id == "tool:" + call.id);
                    if (pending != null) { pending.Status = "verified"; pending.AfterFingerprint = "cancelled_before_live_write"; _taskContext.Checkpoint(); }
                }
                throw;
            }
            catch (Exception ex)
            {
                return new MailboxToolResult(call.id, _serializer.Serialize(new { error_code = "POWERPOINT_REVISION_FAILED", message = ex.Message,
                    status = revision?.Status ?? "preflight", permission_consumed = written && revision?.Status != "rolled_back", saved = false }), ex.Message);
            }
            finally
            {
                if (revision != null) revision.CloseStaging(written || revision.Status == "applied" || revision.Status == "recovery_required");
            }
        }
        private void CheckpointRevision(ChatToolCall call, PresentationRevision revision, string status)
        {
            if (_taskContext == null) return;
            _taskContext.State.HostData["powerpoint_revision_status"] = status;
            _taskContext.State.HostData["powerpoint_revision_input"] = SamsungGenerationJournal.InputHash(call);
            _taskContext.State.HostData["powerpoint_revision_snapshot"] = revision.Snapshot();
            _taskContext.State.PresentationReviewRequired = true;
            _taskContext.State.PresentationReviewReceipt = null;
            _taskContext.Checkpoint();
        }
    }
}
