using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Scribble.Chat;
using Scribble.Configuration;
using Scribble.Security;

namespace Scribble.Office
{
    public sealed partial class DocumentDraftHost
    {
        private async Task<MailboxToolResult> ExecuteAnalysisDeckAsync(
            ChatToolCall call, OneShotDraftAuthorization authorization,
            bool exclusive, OpenAiCompatibleClient client,
            AppSettings settings, CancellationToken token)
        {
            if (!exclusive || authorization == null ||
                !authorization.CanCreate || _taskContext == null)
                return Error(call.id, authorization,
                    "DRAFT_PERMISSION_NOT_AVAILABLE",
                    "The deck needs the user's explicit draft instruction and an exclusive call.");
            AnalysisArtifact artifact;
            AnalysisDocumentPlan plan;
            IReadOnlyList<PresentationDraftWriter.DraftSlide> slides;
            try
            {
                artifact = _taskContext.LoadAnalysis();
                if (_taskContext.State.HostData.ContainsKey(
                        "analysis_deck_complete") &&
                    !_taskContext.State.HostData.ContainsKey(
                        "samsung_pending"))
                    throw new InvalidOperationException(
                        "ANALYSIS_DECK_ALREADY_COMPLETE");
                plan = AnalysisSlidePlanContract.Parse(artifact,
                    call.function.arguments);
                var compiled = AnalysisDocumentCompiler.Compile(artifact,
                    plan);
                slides = PresentationDraftWriter.ParseSlides(
                    compiled.Slides.Cast<object>().ToArray());
                var pages = PresentationDraftWriter.ComposeSamsung(slides);
                if (_taskContext.State.RequiredPresentationSlides > 0 &&
                    pages.Count != _taskContext.State.RequiredPresentationSlides)
                    throw new InvalidOperationException(
                        "ANALYSIS_DECK_PAGE_COUNT_INVALID");
                if (client == null || settings == null ||
                    !ModelCatalog.IsVisionCapable(settings.Model))
                    throw new InvalidOperationException(
                        "ANALYSIS_DECK_VISION_REQUIRED");
                OfficeTaskBinding.Validate(_taskContext.State, "excel",
                    _hostApplication);
                AnalysisWorkbookSourceGuard.Validate(_hostApplication,
                    artifact);
                token.ThrowIfCancellationRequested();
            }
            catch (Exception exception) when (!(exception is
                OperationCanceledException))
            {
                return Error(call.id, authorization,
                    "ANALYSIS_DECK_PREFLIGHT_FAILED", exception.Message);
            }

            var written = false;
            var outputs = new List<PresentationDraftWriter.SamsungOutput>();
            SamsungGenerationJournal journal = null;
            try
            {
                var app = GetSiblingApplication("PowerPoint.Application");
                if (_taskContext.State.HostData.ContainsKey(
                    "samsung_destination"))
                {
                    dynamic application = app;
                    var matches = new List<object>();
                    for (var p = 1; p <=
                        (int)application.Presentations.Count; p++)
                    {
                        dynamic candidate = application.Presentations[p];
                        if (string.Equals((string)candidate.Tags[
                                "ScribbleTask"], _taskContext.State.Id,
                            StringComparison.OrdinalIgnoreCase))
                            matches.Add((object)candidate);
                    }
                    if (matches.Count != 1)
                        throw new InvalidOperationException(
                            "ANALYSIS_DECK_DESTINATION_MISSING");
                    _samsungPresentation = matches[0];
                }
                var resuming = _taskContext.State.HostData.ContainsKey(
                    "samsung_pending");
                journal = new SamsungGenerationJournal(_taskContext, call);
                Action beforeNativeWrite = () =>
                {
                    if (written) return;
                    token.ThrowIfCancellationRequested();
                    if (!resuming && !authorization.TryConsume())
                        throw new InvalidOperationException(
                            "DRAFT_PERMISSION_NOT_AVAILABLE");
                    _taskContext.State.PresentationReviewRequired = true;
                    _taskContext.State.PresentationReviewReceipt = null;
                    foreach (var id in plan.Slides.Select(slide =>
                        "ppt:" + slide.Id))
                        if (!_taskContext.State.ExpectedSourceIds.Contains(id))
                            _taskContext.State.ExpectedSourceIds.Add(id);
                    _taskContext.Checkpoint();
                    written = true;
                };
                var status = PresentationDraftWriter.AddDraftSlides(app,
                    slides, null, true, output =>
                    {
                        outputs.Add(output);
                        _samsungPresentation = (object)((dynamic)
                            output.Slide).Parent;
                    }, _samsungPresentation, journal, beforeNativeWrite,
                    true);
                dynamic deck = _samsungPresentation;
                if (deck == null ||
                    !string.Equals((string)deck.Tags["ScribbleTask"],
                        _taskContext.State.Id,
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "ANALYSIS_DECK_OWNERSHIP_CHANGED");

                // Native geometry belongs to the renderer. A measured repair
                // consumes the shared task budget before touching its target.
                for (var repairs = 0; repairs <
                    AnalysisRepairBudget.MaxCorrectivePatches; repairs++)
                {
                    var nativePages = AnalysisDocumentPilot
                        .CapturePresentationPages((object)deck, artifact,
                            plan);
                    var defects = AnalysisDocumentPilot
                        .CaptureNativeMeasurements((object)deck,
                            nativePages);
                    if (defects.Count == 0) break;
                    var defect = defects[0];
                    var page = nativePages.Single(value =>
                        value.NativeSlideId == defect.NativeSlideId);
                    var reservation = _taskContext.ReserveAnalysisPatch(
                        page, defect, true);
                    AnalysisDocumentPilot.RepairNativeMeasurement(
                        (object)deck, nativePages, defect,
                        reservation.BudgetReceipt, reservation);
                    var repairedPages = AnalysisDocumentPilot
                        .CapturePresentationPages((object)deck, artifact,
                            plan);
                    var remaining = AnalysisDocumentPilot
                        .CaptureNativeMeasurements((object)deck,
                            repairedPages);
                    journal.Record(outputs[page.ExpectedPageNumber - 1],
                        page.ExpectedPageNumber - 1);
                    _taskContext.ReconcileAnalysisPatch(reservation,
                        repairedPages.Single(value =>
                            value.NativeSlideId == defect.NativeSlideId),
                        remaining);
                }
                var review = AnalysisDocumentPilot.ReserveNativeReview(
                    _taskContext, (object)deck, artifact, plan, true);
                if (review.Context.Measurements.Count != 0)
                    throw new InvalidOperationException(
                        "ANALYSIS_DECK_GEOMETRY_UNRESOLVED");
                var parts = new List<object>
                {
                    new ChatMultimodalTextPart { type = "text",
                        text = review.Request.Content }
                };
                foreach (var image in review.PageImages)
                    parts.Add(new ChatMultimodalImagePart
                    {
                        type = "image_url",
                        image_url = new ChatMultimodalImageUrl { url = image }
                    });
                var response = await client.CompleteAsync(settings,
                    new ChatCompletionRequest
                    {
                        Diagnostics = _taskContext.Diagnostics,
                        model = settings.Model,
                        max_tokens = review.Request.MaxResponseTokens,
                        messages = new List<object>
                        {
                            new ChatCompletionInputMessage
                            {
                                role = "system",
                                content = review.Request.Instructions
                            },
                            new ChatCompletionInputMessage
                            {
                                role = "user", content = parts.ToArray()
                            }
                        }
                    }, token);
                var verdict = AnalysisDocumentPilot.CompleteNativeReview(
                    (object)deck, review,
                    (response.RawContent ?? response.content ?? "").Trim());
                if (!verdict.Approved)
                    throw new InvalidOperationException(
                        "ANALYSIS_DECK_REVIEW_REJECTED: " +
                        _serializer.Serialize(verdict.Findings));
                _taskContext.State.PresentationReviewReceipt =
                    review.Context.ContextId;
                _taskContext.State.HostData["analysis_deck_complete"] =
                    review.Context.ContextId;
                foreach (var id in plan.Slides.Select(slide => slide.Id))
                    if (!_taskContext.State.Batches.Any(batch =>
                            batch.Id == "ppt:" + id))
                        _taskContext.State.Batches.Add(new TaskBatchResult
                        {
                            Id = "ppt:" + id,
                            CoveredSourceIds = new List<string> { "ppt:" + id },
                            Output = "Typed native review passed"
                        });
                _taskContext.Checkpoint();
                journal.Complete();
                authorization.MarkCreated();
                return new MailboxToolResult(call.id,
                    _serializer.Serialize(new
                    {
                        ok = true, saved = false,
                        analysis_id = artifact.AnalysisId,
                        native_pages = review.Context.Pages.Count,
                        review_context_id = review.Context.ContextId,
                        status
                    }), status);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception exception)
            {
                if (exception is COMException &&
                    (unchecked((uint)exception.HResult) == 0x800706BA ||
                     unchecked((uint)exception.HResult) == 0x800706BE ||
                     unchecked((uint)exception.HResult) == 0x80010108))
                    return Error(call.id, authorization,
                        "POWERPOINT_EXITED",
                        "PowerPoint exited during native deck creation. The draft remains pending; inspect the task before retrying.");
                if (exception.Message.StartsWith(
                        "ANALYSIS_VISUAL_REVIEW_UNAVAILABLE:",
                        StringComparison.Ordinal))
                    return Error(call.id, authorization,
                        "ANALYSIS_VISUAL_REVIEW_UNAVAILABLE",
                        exception.Message);
                if (exception.Message.StartsWith(
                        "ANALYSIS_PILOT_GEOMETRY_UNSUPPORTED:",
                        StringComparison.Ordinal))
                    return Error(call.id, authorization,
                        "ANALYSIS_DECK_GEOMETRY_UNSUPPORTED",
                        exception.Message);
                return Error(call.id, authorization,
                    "ANALYSIS_DECK_FAILED", exception.Message);
            }
            finally
            {
                foreach (var output in outputs)
                    if (System.Runtime.InteropServices.Marshal.IsComObject(
                            output.Slide))
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(
                            output.Slide);
                var retained = _samsungPresentation;
                _samsungPresentation = null;
                if (retained != null &&
                    System.Runtime.InteropServices.Marshal.IsComObject(
                        retained))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(
                        retained);
            }
        }

        private MailboxToolResult ExecuteAnalysisWorkbookDraft(
            string callId, IDictionary<string, object> arguments,
            OneShotDraftAuthorization authorization)
        {
            if (_hostKind != "excel" ||
                !string.Equals(Environment.GetEnvironmentVariable(
                    AnalysisDocumentPilot.FeatureFlag), "1",
                    StringComparison.Ordinal))
                return Error(callId, authorization,
                    "ANALYSIS_PILOT_DISABLED",
                    "This typed report route is available only in the development pilot.");
            if (_taskContext == null || authorization == null ||
                !authorization.CanCreate)
                return Error(callId, authorization,
                    "DRAFT_PERMISSION_NOT_AVAILABLE",
                    "The task needs the user's explicit draft instruction.");
            if (arguments.Keys.Except(new[] { "analysis_id", "title" },
                    StringComparer.Ordinal).Any())
                return Error(callId, authorization,
                    "ANALYSIS_DRAFT_ARGUMENTS_INVALID",
                    "Supply only analysis_id and an optional title; the host builds rows and formulas.");
            AnalysisArtifact artifact;
            AnalysisDocumentPlan plan;
            int formulaCount;
            try
            {
                artifact = _taskContext.LoadAnalysis();
                if (artifact == null ||
                    !string.Equals(ToolArguments.GetString(arguments,
                            "analysis_id", string.Empty),
                        artifact.AnalysisId, StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "ANALYSIS_PLAN_BINDING_INVALID");
                var title = ToolArguments.GetString(arguments, "title",
                    "Verified analysis").Trim();
                if (title.Length == 0 || title.Length > 120 ||
                    title.StartsWith("=", StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "ANALYSIS_WORKBOOK_TITLE_INVALID");
                plan = new AnalysisDocumentPlan
                {
                    AnalysisId = artifact.AnalysisId,
                    WorkbookTitle = title,
                    WorkbookRows = AnalysisWorkbookPlanBuilder.Build(
                        artifact)
                };
                formulaCount = AnalysisDocumentCompiler.Compile(artifact,
                    plan, false).ExpectedFormulaFacts.Count;
                OfficeTaskBinding.Validate(_taskContext.State, "excel",
                    _hostApplication);
                AnalysisWorkbookSourceGuard.Validate(_hostApplication,
                    artifact);
            }
            catch (Exception exception)
            {
                return Error(callId, authorization,
                    "ANALYSIS_DRAFT_PREFLIGHT_FAILED", exception.Message);
            }
            if (!authorization.TryConsume())
                return Error(callId, authorization,
                    "DRAFT_PERMISSION_NOT_AVAILABLE",
                    "The task's draft call budget is exhausted.");
            try
            {
                var status = AnalysisDocumentPilot.WriteWorkbook(
                    _hostApplication, artifact, plan);
                authorization.MarkCreated();
                return new MailboxToolResult(callId,
                    _serializer.Serialize(new
                    {
                        ok = true, saved = false,
                        analysis_id = artifact.AnalysisId,
                        verified_live_formulas = formulaCount,
                        status
                    }), status);
            }
            catch (Exception exception)
            {
                return Error(callId, authorization,
                    "ANALYSIS_DRAFT_FAILED", exception.Message);
            }
        }
    }
}
