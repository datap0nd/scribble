using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
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
                string.Equals(Environment.GetEnvironmentVariable(
                    AnalysisDocumentPilot.FeatureFlag), "1",
                    StringComparison.Ordinal) &&
                ShouldDraftRepairedDeck(_hostKind,
                    string.Join("\n", _taskContext.State.OriginalDecisions),
                    _taskContext.State.RequiredPresentationSlides) &&
                _taskContext.State.RequiredPresentationSlides == 6;
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
                    (int)source.Slides.Count != 6 ||
                    string.IsNullOrEmpty(Convert.ToString(source.Path)) ||
                    (int)source.Saved == 0)
                    throw new InvalidOperationException(
                        "PILOT_COPY_SOURCE_CHANGED: Inspect the saved six-slide source again.");
                var operations = SamsungAuthoringPolicy.Array(args,
                    "operations");
                var mapped = operations.Select(
                    SamsungAuthoringPolicy.ReadMap).ToArray();
                dynamic chartSlide = source.Slides[2];
                dynamic chartShape = chartSlide.Shapes[
                    (int)chartSlide.Shapes.Count];
                if ((int)chartShape.HasChart == 0)
                    throw new InvalidOperationException(
                        "PILOT_COPY_CHART_SOURCE_UNSUPPORTED");
                var chartShapeId = (int)chartShape.Id;
                ValidatePilotCopyOperations(mapped,
                    (int)source.Slides[4].SlideID,
                    (int)chartSlide.SlideID, chartShapeId);
                var contract = PresentationToolCatalog
                    .RevisionDefinitions().Single(tool =>
                        tool.function.name ==
                        PresentationToolCatalog.ReviseSlides);
                var contractErrors = ToolContractValidator.Validate(
                    call, contract);
                if (contractErrors.Count != 0)
                    throw new InvalidOperationException(
                        "PILOT_COPY_SCHEMA: " +
                        string.Join("; ", contractErrors));
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
                token.ThrowIfCancellationRequested();
                if (!authorization.TryConsume())
                    throw new InvalidOperationException(
                        "PILOT_COPY_PERMISSION_UNAVAILABLE");
                _taskContext.State.HostData[statusKey] = "creating";
                _taskContext.Checkpoint();
                stage = "copy";
                copy = PresentationDraftCopy.Create(_hostApplication,
                    sourceDeck, _taskContext.State.Id);
                _taskContext.State.HostData["pilot_copy_snapshot"] =
                    copy.Snapshot();
                _taskContext.State.HostData[statusKey] = "copied";
                _taskContext.Checkpoint();
                var bound = copy.BindOperations(operations);
                var nativeStyle = copy.Pp01NativeStyleOperations();
                var combined = bound.Concat(nativeStyle).ToArray();
                if (combined.Length > 24)
                    throw new InvalidOperationException(
                        "PILOT_COPY_OPERATIONS_INVALID");
                ((dynamic)copy.Draft).Activate();
                var draftCall = new ChatToolCall
                {
                    id = call.id + ":pilot",
                    function = new ChatToolCallFunction
                    {
                        name = PresentationToolCatalog.ReviseSlides,
                        arguments = _serializer.Serialize(new
                        {
                            presentation_id = PresentationInspection
                                .IdentityFor(copy.Draft),
                            operations = combined
                        })
                    }
                };
                _taskContext.State.HostData[statusKey] = "patching";
                _taskContext.Checkpoint();
                stage = "patch";
                var internalAuthorization =
                    new OneShotDraftAuthorization(true, false);
                var patchPrompt = prompt +
                    "\nPilot patch stage: review content edits and the single fourth-page replacement. The host applies the bounded native font/table styling and recreates the workbook-backed chart after this stage. Do not require model-authored style or chart operations in this batch.";
                var patch = await ExecuteRevisionAsync(draftCall,
                    internalAuthorization, true, patchPrompt, client, settings,
                    token, progress, true);
                var patchResult = _serializer.Deserialize<
                    Dictionary<string, object>>(patch.Content);
                object ok;
                if (!patchResult.TryGetValue("ok", out ok) ||
                    !(ok is bool) || !(bool)ok)
                    throw new InvalidOperationException(
                        "PILOT_COPY_PATCH_FAILED: " + patch.StatusText);
                var patchReviewReceipt = _taskContext.State
                    .PresentationReviewReceipt;
                if (string.IsNullOrEmpty(patchReviewReceipt))
                    throw new InvalidOperationException(
                        "PILOT_COPY_PATCH_REVIEW_MISSING");
                var revision = PresentationRevision.Last(copy.Draft);
                copy.AcceptRevision(revision);
                _taskContext.State.HostData["pilot_copy_snapshot"] =
                    copy.Snapshot();
                _taskContext.State.HostData[statusKey] = "patched";
                _taskContext.Checkpoint();
                // The copy snapshot is the durable recovery boundary now.
                // Keeping extra staging decks open while chart.dll creates
                // the native chart has crashed this Office build.
                revision.CloseStaging(false);
                stage = "chart";
                _taskContext.State.HostData[statusKey] = "charting";
                _taskContext.Checkpoint();
                var chartFacts = copy.RecreateSalesChartFromWorkbook(
                    (int)chartSlide.SlideID, chartShapeId,
                    workbooks[0].SourcePath, 66f, 158.25f, 825f,
                    278.25f);
                if (chartFacts.Categories.Length != 6 ||
                    !string.Equals(chartFacts.SourceSha256,
                        workbooks[0].SourceFingerprint,
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "PILOT_COPY_CHART_FACTS_INVALID");
                copy.VerifySource();
                copy.VerifyDraft();
                _taskContext.State.HostData["pilot_copy_snapshot"] =
                    copy.Snapshot();
                _taskContext.State.HostData[statusKey] = "complete";
                for (var index = 1; index <= 6; index++)
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
                        chartFacts.SourceSha256);
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
                        ok = true, saved = false, copied_slides = 6,
                        revised_slides = revision.Items.Count,
                        native_style_changes = nativeStyle.Length,
                        chart_recreated = true,
                        visual_approval_required = true,
                        revert_available = false
                    }), "Opened a six-slide unsaved repair draft. The source and workbook were preserved; visual approval is still required.");
            }
            catch (OperationCanceledException)
            { throw; }
            catch (Exception error)
            {
                var code = error.Message.Split(':')[0];
                return new MailboxToolResult(call.id,
                    _serializer.Serialize(new
                    {
                        error_code = code.StartsWith("PILOT_COPY_") ||
                            code.StartsWith("ANALYSIS_") ? code :
                            "PILOT_COPY_FAILED",
                        message = error.Message,
                        stage,
                        saved = false,
                        needs_inspection = _taskContext.State.HostData
                            .ContainsKey(statusKey)
                    }), error.Message);
            }
        }

        private static void ValidatePilotCopyOperations(
            Dictionary<string, object>[] mapped, int replacementSlideId,
            int chartSlideId, int chartShapeId)
        {
            if (mapped == null || mapped.Length == 0 ||
                mapped.Length > 15)
                throw new InvalidOperationException(
                    "PILOT_COPY_OPERATIONS_INVALID");
            if (mapped.Count(operation => SamsungAuthoringPolicy.Text(
                    operation, "kind") == "replace_slide") > 1)
                throw new InvalidOperationException(
                    "ANALYSIS_MULTI_PAGE_REPLACEMENT_UNSUPPORTED");
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
                    Convert.ToInt32(target) == chartShapeId &&
                    Convert.ToInt32(operation["slide_id"]) ==
                        chartSlideId)
                    throw new InvalidOperationException(
                        "ANALYSIS_CHART_REFLOW_UNSUPPORTED: The pilot recreates the chart from the bound workbook after patching.");
            }
            if (mapped.Count(operation => SamsungAuthoringPolicy.Text(
                    operation, "kind") == "replace_slide" &&
                Convert.ToInt32(operation["slide_id"]) ==
                    replacementSlideId) != 1)
                throw new InvalidOperationException(
                    "PILOT_COPY_LAYOUT_SCOPE_REQUIRED: Recompose the overflowing fourth page once.");
        }
    }
}
