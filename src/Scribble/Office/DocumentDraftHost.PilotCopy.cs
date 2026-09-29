using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Configuration;
using Scribble.Security;

namespace Scribble.Office
{
    public sealed partial class DocumentDraftHost
    {
        private bool PilotCopyRequested(ChatToolCall call)
        {
            return call?.function?.name ==
                    PresentationToolCatalog.ReviseSlides &&
                _hostKind == "powerpoint" && _taskContext != null &&
                AnalysisDocumentPilot.Enabled &&
                ShouldDraftRepairedDeck(_hostKind,
                    string.Join("\n", _taskContext.State.OriginalDecisions),
                    _taskContext.State.RequiredPresentationSlides);
        }

        private async Task<MailboxToolResult> ExecutePilotCopyRevisionAsync(
            ChatToolCall call, OneShotDraftAuthorization authorization,
            bool exclusive, string prompt, OpenAiCompatibleClient client,
            AppSettings settings, CancellationToken token,
            Action<int, int> progress)
        {
            PresentationDraftCopy copy = null;
            var stage = "preflight";
            var statusKey = "pilot_copy_status";
            var permissionConsumed = false;
            try
            {
                if (!PresentationRevisionAcceptance.Enabled || !exclusive ||
                    authorization == null || !authorization.CanCreate)
                    throw new InvalidOperationException(
                        "PILOT_COPY_NOT_AUTHORIZED");
                if (client == null || settings == null ||
                    !ModelCatalog.IsVisionCapable(settings.Model))
                    throw new InvalidOperationException(
                        "PILOT_COPY_VISION_REQUIRED");
                if (_taskContext.State.HostData.ContainsKey(statusKey))
                    throw new InvalidOperationException(
                        "PILOT_COPY_NEEDS_INSPECTION: A prior copy attempt has a durable checkpoint; do not repeat it.");
                dynamic app = _hostApplication;
                object sourceDeck = app.ActivePresentation;
                dynamic source = sourceDeck;
                var args = ToolArguments.Parse(_serializer,
                    call.function.arguments);
                if (SamsungAuthoringPolicy.Text(args,
                        "presentation_id") != PresentationInspection
                            .IdentityFor(sourceDeck) ||
                    (int)source.Slides.Count !=
                        _taskContext.State.RequiredPresentationSlides ||
                    string.IsNullOrEmpty(Convert.ToString(source.Path)) ||
                    (int)source.Saved == 0)
                    throw new InvalidOperationException(
                        "PILOT_COPY_SOURCE_CHANGED: Inspect the saved source deck again.");
                var operations = SamsungAuthoringPolicy.Array(args,
                    "operations");
                var mapped = operations.Select(
                    SamsungAuthoringPolicy.ReadMap).ToArray();
                if (PresentationRevisionAcceptance
                    .ContainsChartOperation(operations))
                    throw new InvalidOperationException(
                        "PILOT_COPY_MODEL_CHART_UNSUPPORTED");
                if (!_taskContext.State.HostData.ContainsKey(
                        "recovery_input"))
                    throw new InvalidOperationException(
                        "PILOT_COPY_WORKBOOK_MISSING");
                var workbooks = TaskRecoveryInput.Read(
                    _taskContext.State).Documents.Where(document =>
                        new[] { ".xlsx", ".xlsm" }.Contains(
                            Path.GetExtension(document.SourcePath ?? ""),
                            StringComparer.OrdinalIgnoreCase)).ToArray();
                if (workbooks.Length != 1 ||
                    string.IsNullOrEmpty(workbooks[0].SourcePath) ||
                    !File.Exists(workbooks[0].SourcePath) ||
                    !string.Equals(workbooks[0].SourceFingerprint,
                        ExternalContextDocument.FingerprintFile(
                            workbooks[0].SourcePath),
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "PILOT_COPY_WORKBOOK_CHANGED: Reattach one saved workbook before repair.");
                var trustedRequest = string.Join("\n",
                    _taskContext.State.OriginalDecisions);
                var chartBindings = PresentationDraftCopy
                    .BindMonthlyCharts(sourceDeck,
                        workbooks[0].SourcePath, trustedRequest, token);
                var factCatalog = RevisionFactCatalog.FromBindings(
                    chartBindings);
                var measuredReplacements = PresentationDraftCopy
                    .MeasuredReplacementSlides(sourceDeck);
                var contract = PresentationToolCatalog
                    .PilotRevisionDefinition(measuredReplacements);
                var contractErrors = ToolContractValidator.Validate(
                    call, contract);
                if (contractErrors.Count != 0)
                    throw new InvalidOperationException(
                        "PILOT_COPY_SCHEMA: " +
                        string.Join("; ", contractErrors));
                ValidatePilotCopyOperations(mapped,
                    measuredReplacements, chartBindings);
                ValidatePilotSourceSpanIds(mapped,
                    ids => _taskContext.Sources.Resolve(ids));
                var sourceFact = chartBindings.SelectMany(binding =>
                        binding.Facts.Facts).Where(fact =>
                        fact.Dimensions.Count == 0)
                    .OrderByDescending(fact => fact.Period,
                        StringComparer.Ordinal).First();
                var sourceReference = "[[fact:" + sourceFact.FactId +
                    ":locator]]";
                foreach (var operation in mapped.Where(item =>
                    SamsungAuthoringPolicy.Text(item, "kind") ==
                        "replace_slide"))
                {
                    var slide = SamsungAuthoringPolicy.ReadMap(
                        operation["slide"]);
                    var sourceSlide = PresentationInspection.FindSlide(
                        sourceDeck, Convert.ToInt32(operation["slide_id"]));
                    slide["sources"] = sourceReference;
                    slide["footnote"] = sourceReference;
                    slide["evidence"] =
                        PresentationInspection.CitationTextFromCaptured(
                            PresentationInspection.Capture(sourceSlide));
                }
                var factBound = factCatalog.BindOperations(
                    mapped.Cast<object>().ToArray());
                token.ThrowIfCancellationRequested();
                if (!authorization.TryConsume())
                    throw new InvalidOperationException(
                        "PILOT_COPY_PERMISSION_UNAVAILABLE");
                permissionConsumed = true;
                _taskContext.State.HostData[statusKey] = "creating";
                _taskContext.Checkpoint();
                stage = "copy";
                copy = PresentationDraftCopy.Create(_hostApplication,
                    sourceDeck, _taskContext.State.Id);
                _taskContext.State.HostData["pilot_copy_snapshot"] =
                    copy.Snapshot();
                _taskContext.State.HostData[statusKey] = "copied";
                _taskContext.Checkpoint();
                var bound = copy.BindOperations(factBound);
                var nativeStyle = copy.MeasuredNativeStyleOperations(
                    measuredReplacements);
                var combined = bound.Concat(nativeStyle).ToArray();
                if (combined.Length > 24)
                    throw new InvalidOperationException(
                        "PILOT_COPY_OPERATIONS_INVALID");
                ((dynamic)copy.Draft).Windows.Item(1).Activate();
                _taskContext.State.HostData[statusKey] = "patching";
                _taskContext.Checkpoint();
                stage = "patch";
                var changed = PresentationRevision.ApplyOwnedDraft(
                    copy.Draft, combined);
                copy.AcceptDirectRevision(changed, combined);
                var patchReviewReceipt = SamsungAuthoringPolicy.CacheKey(
                    settings.Model, settings.BaseUrl,
                    SerializeChangedSlideFingerprints(changed),
                    copy.Snapshot() + "|" +
                    workbooks[0].SourceFingerprint);
                _taskContext.State.HostData["pilot_copy_snapshot"] =
                    copy.Snapshot();
                _taskContext.State.HostData[statusKey] = "patched";
                _taskContext.Checkpoint();
                stage = "chart";
                _taskContext.State.HostData[statusKey] = "charting";
                _taskContext.Checkpoint();
                foreach (var binding in chartBindings)
                {
                    if (!string.Equals(binding.Facts.SourceSha256,
                            workbooks[0].SourceFingerprint,
                            StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException(
                            "PILOT_COPY_CHART_FACTS_INVALID");
                    copy.RecreateBoundChart(binding);
                }
                copy.VerifySource();
                copy.VerifyDraft();
                _taskContext.State.HostData["pilot_copy_snapshot"] =
                    copy.Snapshot();
                _taskContext.State.HostData[statusKey] = "complete";
                for (var index = 1; index <=
                    (int)source.Slides.Count; index++)
                {
                    var pageId = "ppt:" + index;
                    if (!_taskContext.State.ExpectedSourceIds
                        .Contains(pageId))
                        _taskContext.State.ExpectedSourceIds.Add(pageId);
                    if (!_taskContext.State.Batches.Any(batch =>
                        batch.Id == pageId))
                        _taskContext.State.Batches.Add(
                            new TaskBatchResult
                            {
                                Id = pageId,
                                CoveredSourceIds = new List<string>
                                    { pageId },
                                Output = "Copied, repaired and verified"
                            });
                }
                _taskContext.State.PresentationReviewRequired = true;
                _taskContext.State.PresentationReviewReceipt =
                    SamsungAuthoringPolicy.CacheKey(settings.Model,
                        settings.BaseUrl, patchReviewReceipt,
                        workbooks[0].SourceFingerprint);
                _taskContext.State.HostData[
                    "pilot_copy_review_scope"] =
                        "model-reviewed chartless content; host-verified workbook chart; human visual approval pending";
                _taskContext.Checkpoint();
                Scribble.Testing.TestLab.RegisterOutput(copy.Draft,
                    "pptx");
                authorization.MarkCreated();
                return new MailboxToolResult(call.id,
                    _serializer.Serialize(new
                    {
                        ok = true, saved = false,
                        copied_slides = (int)source.Slides.Count,
                        revised_slides = changed.Count,
                        native_style_changes = nativeStyle.Length,
                        charts_recreated = chartBindings.Length,
                        visual_approval_required = true,
                        revert_available = false
                    }), "Opened an unsaved native repair draft. The source and workbook were preserved; visual approval is still required.");
            }
            catch (OperationCanceledException)
            {
                if (copy != null)
                {
                    try
                    {
                        copy.DiscardOwnedDraft();
                        _taskContext.State.HostData.Remove(statusKey);
                        _taskContext.State.HostData.Remove(
                            "pilot_copy_snapshot");
                        _taskContext.Checkpoint();
                    }
                    catch
                    {
                        // Retain an uncertain draft for inspection.
                    }
                }
                throw;
            }
            catch (Exception error)
            {
                var code = error.Message.Split(':')[0];
                var needsInspection = permissionConsumed ||
                    _taskContext.State.HostData.ContainsKey(statusKey);
                if (copy != null)
                {
                    try
                    {
                        copy.DiscardOwnedDraft();
                        _taskContext.State.HostData.Remove(statusKey);
                        _taskContext.State.HostData.Remove(
                            "pilot_copy_snapshot");
                        _taskContext.Checkpoint();
                        needsInspection = false;
                    }
                    catch
                    {
                        // An unexpected owner or saved path must be inspected;
                        // never close a deck whose ownership is uncertain.
                    }
                }
                return new MailboxToolResult(call.id,
                    _serializer.Serialize(new
                    {
                        error_code = code.StartsWith("PILOT_COPY_") ||
                            code.StartsWith("ANALYSIS_") ? code :
                            "PILOT_COPY_FAILED",
                        message = error.Message,
                        stage,
                        saved = false,
                        needs_inspection = needsInspection,
                        permission_consumed = permissionConsumed
                    }), error.Message);
            }
        }

        internal static string SerializeChangedSlideFingerprints(
            IDictionary<int, string> changed)
        {
            if (changed == null || changed.Count == 0)
                throw new InvalidOperationException(
                    "PILOT_COPY_RECEIPT_INVALID");
            return new JavaScriptSerializer { MaxJsonLength = 16000000 }
                .Serialize(changed.OrderBy(pair => pair.Key)
                    .Select(pair => new {
                        slide_id = pair.Key,
                        fingerprint = pair.Value
                    }).ToArray());
        }

        private static void ValidatePilotCopyOperations(
            Dictionary<string, object>[] mapped,
            int[] measuredReplacementIds,
            PresentationDraftCopy.MonthlyChartBinding[] charts)
        {
            if (mapped == null || mapped.Length == 0 ||
                mapped.Length > 24 ||
                measuredReplacementIds == null || charts == null ||
                charts.Length == 0)
                throw new InvalidOperationException(
                    "PILOT_COPY_OPERATIONS_INVALID");
            var replacements = mapped.Where(operation =>
                SamsungAuthoringPolicy.Text(operation,
                    "kind") == "replace_slide").Select(operation =>
                Convert.ToInt32(operation["slide_id"])).ToArray();
            if (replacements.Distinct().Count() !=
                    replacements.Length)
                throw new InvalidOperationException(
                    "PILOT_COPY_REPLACEMENT_AMBIGUOUS");
            if (replacements.Any(id => charts.Any(chart =>
                    chart.SourceSlideId == id)))
                throw new InvalidOperationException(
                    "PILOT_COPY_CHART_SLIDE_REPLACEMENT_UNSUPPORTED");
            foreach (var operation in mapped)
            {
                var kind = SamsungAuthoringPolicy.Text(operation,
                    "kind");
                if (kind == "shape_geometry")
                    throw new InvalidOperationException(
                        "ANALYSIS_GEOMETRY_UNSUPPORTED");
                if (kind.IndexOf("chart", StringComparison
                        .OrdinalIgnoreCase) >= 0)
                    throw new InvalidOperationException(
                        "ANALYSIS_CHART_REFLOW_UNSUPPORTED");
                if (!new[] { "replace_text", "table_cell",
                        "replace_slide", "annotate", "notes_append" }
                    .Contains(kind))
                    throw new InvalidOperationException(
                        "PILOT_COPY_OPERATION_UNSUPPORTED");
                object target;
                if (operation.TryGetValue("shape_id", out target) &&
                    charts.Any(chart =>
                        Convert.ToInt32(operation["slide_id"]) ==
                            chart.SourceSlideId &&
                        Convert.ToInt32(target) ==
                            chart.SourceShapeId))
                    throw new InvalidOperationException(
                        "ANALYSIS_CHART_REFLOW_UNSUPPORTED: The host recreates monthly charts from the bound workbook after patching.");
            }
            if (!replacements.OrderBy(id => id).SequenceEqual(
                    measuredReplacementIds.OrderBy(id => id)))
                throw new InvalidOperationException(
                    "PILOT_COPY_LAYOUT_SCOPE_REQUIRED: Replace exactly the source slides with measured overflow: " +
                    string.Join(", ", measuredReplacementIds));
        }

        internal static void ValidatePilotCopyTextEvidence(
            Dictionary<string, object>[] operations, string source)
        {
            if (operations == null)
                throw new InvalidOperationException("PILOT_COPY_OPERATIONS_INVALID");
            foreach (var operation in operations)
            {
                var kind = SamsungAuthoringPolicy.Text(operation, "kind");
                if (kind == "replace_text")
                {
                    var before = SamsungAuthoringPolicy.Text(operation,
                        "before");
                    var after = SamsungAuthoringPolicy.Text(operation,
                        "text");
                    var oldNumbers = Regex.Matches(before,
                        @"\d[\d,.]*%?")
                        .Cast<Match>().Select(match => match.Value);
                    var newNumbers = Regex.Matches(after,
                        @"\d[\d,.]*%?")
                        .Cast<Match>().Select(match => match.Value);
                    if (!oldNumbers.SequenceEqual(newNumbers))
                        throw new InvalidOperationException(
                            "PILOT_COPY_TEXT_NUMBER_CHANGED: Use a cited replacement slide for changed numeric claims.");
                }
                else if (kind == "notes_append")
                {
                    var notes = SamsungAuthoringPolicy.Text(operation,
                        "notes").Trim();
                    if (notes.Length > 0 &&
                        (source ?? "").IndexOf(notes,
                            StringComparison.OrdinalIgnoreCase) < 0)
                        throw new InvalidOperationException(
                            "PILOT_COPY_NOTES_UNVERIFIED");
                }
            }
        }

        internal static void ValidatePilotSourceSpanIds(
            Dictionary<string, object>[] operations,
            Func<IEnumerable<string>, string> resolve)
        {
            if (operations == null || resolve == null)
                throw new InvalidOperationException(
                    "PILOT_COPY_SOURCE_SPANS_INVALID");
            foreach (var operation in operations)
            {
                object supplied;
                if (!operation.TryGetValue("slide", out supplied)) continue;
                var content = SamsungAuthoringPolicy.ReadMap(supplied);
                var spans = SamsungAuthoringPolicy.Array(content,
                    "source_spans");
                if (spans.Length > 0)
                    resolve(spans.Select(Convert.ToString));
            }
        }

        internal static void ValidateRevisionSlideEvidence(
            Dictionary<string, object>[] operations, string source,
            Func<IEnumerable<string>, string> resolve,
            System.Web.Script.Serialization.JavaScriptSerializer serializer)
        {
            if (operations == null || resolve == null || serializer == null)
                throw new InvalidOperationException("PILOT_COPY_OPERATIONS_INVALID");
            foreach (var operation in operations)
            {
                object supplied;
                if (!operation.TryGetValue("slide", out supplied)) continue;
                var content = SamsungAuthoringPolicy.ReadMap(supplied);
                var spans = SamsungAuthoringPolicy.Array(content,
                    "source_spans");
                string resolvedEvidence = null;
                if (spans.Length > 0)
                {
                    resolvedEvidence = resolve(spans.Select(Convert.ToString));
                    content["evidence"] = resolvedEvidence;
                }
                SamsungPresentationReview.ValidateEvidence(
                    serializer.Serialize(content), source, resolvedEvidence);
            }
        }
    }
}
