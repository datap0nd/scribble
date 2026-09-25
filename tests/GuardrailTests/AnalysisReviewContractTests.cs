using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Office;

namespace GuardrailTests
{
    internal static class AnalysisReviewContractTests
    {
        public static void PdfExportOnlyPermitsMetadataAndTableRoundoff()
        {
            var root = Path.Combine(Path.GetTempPath(),
                "scribble-pdf-boundary-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                var before = Path.Combine(root, "before.pptx");
                var harmless = Path.Combine(root, "harmless.pptx");
                var changedData = Path.Combine(root, "changed-data.pptx");
                var changedText = Path.Combine(root, "changed-text.pptx");
                var changedGeometry = Path.Combine(root, "changed-geometry.pptx");
                WritePdfBoundaryPackage(before, 100, 4, "Safe", "42");
                WritePdfBoundaryPackage(harmless, 102, 5, "Safe", "42");
                WritePdfBoundaryPackage(changedData, 101, 5, "Safe", "43");
                WritePdfBoundaryPackage(changedText, 101, 5, "Altered", "42");
                WritePdfBoundaryPackage(changedGeometry, 103, 5, "Safe", "42");
                var method = typeof(PresentationInspection).GetMethod(
                    "PdfExportPackageEquivalent", BindingFlags.NonPublic |
                    BindingFlags.Static);
                Check(method != null, "The PDF package boundary is missing.");
                Func<string, bool> equivalent = candidate =>
                {
                    var arguments = new object[] { before, candidate, null };
                    return (bool)method.Invoke(null, arguments);
                };
                Check(equivalent(harmless),
                    "Office metadata and two EMU table rounding were rejected.");
                Check(!equivalent(changedData) && !equivalent(changedText) &&
                    !equivalent(changedGeometry),
                    "A chart value, slide text or material geometry change crossed the PDF boundary.");
            }
            finally { Directory.Delete(root, true); }
        }

        public static void ChartFingerprintPreservesEmbeddedWorkbookData()
        {
            var method = typeof(PresentationInspection).GetMethod(
                "EmbeddedWorkbookFingerprint", BindingFlags.NonPublic |
                BindingFlags.Static);
            Check(method != null,
                "The chart workbook fingerprint boundary is missing.");
            Func<int, string, string, string, string> fingerprint =
                (revision, modified, title, value) => (string)method.Invoke(
                    null, new object[] { ChartWorkbookPackage(revision,
                        modified, title, value) });
            var original = fingerprint(1, "2026-09-24T12:00:00Z",
                "Ledger", "42");
            Check(original == fingerprint(2, "2026-09-24T12:01:00Z",
                    "Ledger", "42"),
                "PowerPoint's embedded workbook revision metadata changed the chart fingerprint.");
            Check(original != fingerprint(1, "2026-09-24T12:00:00Z",
                    "Ledger", "43") &&
                original != fingerprint(1, "2026-09-24T12:00:00Z",
                    "Changed title", "42"),
                "A workbook value or substantive metadata edit escaped the chart fingerprint.");
        }

        public static void SavedDeckCannotUseChartPackageFingerprint()
        {
            var method = typeof(PresentationInspection).GetMethod(
                "OwnedUnsavedDraft", BindingFlags.NonPublic |
                BindingFlags.Static);
            Check(method != null,
                "The chart package ownership boundary is missing.");
            Func<string, string, string, string, string, bool> allowed =
                (path, slideOwner, deckOwner, journalOwner,
                    revisionOwner) =>
                    (bool)method.Invoke(null, new object[] {
                        path, slideOwner, deckOwner, journalOwner,
                        revisionOwner });
            Check(!allowed(@"C:\\user\\saved.pptx", "journal",
                    "task", "journal", "revision") &&
                !allowed("", "", "", "", "") &&
                !allowed("", "journal-a", "task", "journal-b", ""),
                "A saved or unowned deck can reach SaveCopyAs.");
            Check(allowed("", "journal", "task", "journal", "") &&
                allowed("", "", "", "", "revision"),
                "An owned unsaved draft cannot reach the package fingerprint.");
        }

        private static byte[] ChartWorkbookPackage(int revision,
            string modified, string title, string value)
        {
            using (var stream = new MemoryStream())
            {
                using (var archive = new ZipArchive(stream,
                    ZipArchiveMode.Create, true))
                {
                    Action<string, string> add = (name, content) =>
                    {
                        using (var writer = new StreamWriter(
                            archive.CreateEntry(name).Open(),
                            new UTF8Encoding(false)))
                            writer.Write(content);
                    };
                    add("docProps/core.xml",
                        "<cp:coreProperties xmlns:cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' xmlns:dcterms='http://purl.org/dc/terms/' xmlns:dc='http://purl.org/dc/elements/1.1/'><cp:revision>" +
                        revision + "</cp:revision><dcterms:modified>" +
                        modified + "</dcterms:modified><dc:title>" +
                        title + "</dc:title></cp:coreProperties>");
                    add("xl/worksheets/sheet1.xml", "<sheet><value>" +
                        value + "</value></sheet>");
                }
                return stream.ToArray();
            }
        }

        private static void WritePdfBoundaryPackage(string path, int width,
            int revision, string title, string chartValue)
        {
            using (var archive = new ZipArchive(File.Create(path),
                ZipArchiveMode.Create))
            {
                Action<string, string> add = (name, value) =>
                {
                    using (var writer = new StreamWriter(
                        archive.CreateEntry(name).Open(), new UTF8Encoding(false)))
                        writer.Write(value);
                };
                add("docProps/core.xml",
                    "<cp:coreProperties xmlns:cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' xmlns:dcterms='http://purl.org/dc/terms/'><cp:revision>" +
                    revision + "</cp:revision><dcterms:modified>2026-09-24</dcterms:modified></cp:coreProperties>");
                add("ppt/slides/slide2.xml",
                    "<p:sld xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'><p:graphicFrame><p:xfrm><a:off x='0' y='0'/><a:ext cx='" +
                    width + "' cy='200'/></p:xfrm><a:graphic><a:graphicData uri='http://schemas.openxmlformats.org/drawingml/2006/table'/></a:graphic></p:graphicFrame><p:sp><a:t>" +
                    title + "</a:t></p:sp></p:sld>");
                add("ppt/charts/chart1.xml", "<chart><value>" +
                    chartValue + "</value></chart>");
            }
        }

        public static void SharedBudgetSurvivesEveryStage()
        {
            var receipt = new AnalysisRepairBudget().Serialize();
            for (var call = 0; call < AnalysisRepairBudget.MaxModelCalls; call++)
                receipt = AnalysisRepairBudget.Read(receipt).ConsumeModelCall(
                    call % 2 == 0 ? 400 : 3200, 2048);
            Reject(() => AnalysisRepairBudget.Read(receipt).ConsumeModelCall(100, 512),
                "REPAIR_MODEL_CALL_LIMIT");
            var budget = AnalysisRepairBudget.Read(receipt);
            receipt = budget.ConsumePatch("june", "subtitle");
            Reject(() => AnalysisRepairBudget.Read(receipt).ConsumePatch("june", "subtitle"),
                "REPAIR_TARGET_ALREADY_PATCHED");
            for (var patch = 1; patch < AnalysisRepairBudget.MaxCorrectivePatches; patch++)
                receipt = AnalysisRepairBudget.Read(receipt).ConsumePatch("june", "block-" + patch);
            Reject(() => AnalysisRepairBudget.Read(receipt).ConsumePatch("june", "extra"),
                "REPAIR_PATCH_LIMIT");
            Reject(() => new AnalysisRepairBudget().ConsumeModelCall(
                AnalysisRepairBudget.MaxPromptCharacters + 1, 100),
                "REVIEW_CONTEXT_LIMIT");
            receipt = AnalysisRepairBudget.CrossApp().Serialize();
            for (var call = 0; call < 12; call++)
                receipt = AnalysisRepairBudget.Read(receipt).ConsumeModelCall(100, 512);
            Reject(() => AnalysisRepairBudget.Read(receipt).ConsumeModelCall(100, 512),
                "REPAIR_MODEL_CALL_LIMIT");
        }

        public static void FindingsCannotOverrideVerifiedFactsOrPages()
        {
            var locator = new SourceLocator
            {
                Kind = "excel_range", SourceInstanceId = "WB01",
                WorksheetIdentity = "Ledger", Range = "A1:C3"
            };
            var snapshot = AnalysisContract.CreateSnapshot("WB01", "excel_workbook",
                "revision-1", "complete_range", "recalculated",
                new[] { locator }, new TableDataset[0]);
            var fact = AnalysisContract.CreateObservedFact(snapshot.SnapshotId,
                "RevenueEUR", AnalysisContract.DecimalValue, "82992", "82,992",
                "currency", "EUR", "2026-06", new Dictionary<string, string>(),
                new[] { locator }, AnalysisContract.Verified);
            var artifact = AnalysisContract.CreateArtifact(new[] { snapshot },
                new[] { fact }, new AnalysisCalculation[0], new string[0],
                new string[0]);
            var plan = new AnalysisDocumentPlan
            {
                AnalysisId = artifact.AnalysisId, WorkbookTitle = "Audit",
                WorkbookRows = new List<AnalysisPlanRow>
                {
                    new AnalysisPlanRow { Cells = new List<AnalysisPlanCell>
                    { new AnalysisPlanCell { Text = "Metric" } } }
                },
                Slides = new List<AnalysisPlanSlide>
                {
                    new AnalysisPlanSlide
                    {
                        Id = "june", Layout = "scorecard", Title = "June revenue",
                        Cards = new List<AnalysisPlanCard>
                        {
                            new AnalysisPlanCard { Heading = "Revenue",
                                Points = new List<AnalysisPlanText>
                                { new AnalysisPlanText { FactId = fact.FactId } } }
                        }
                    }
                }
            };
            var page = new AnalysisReviewPage
            {
                LogicalSlideId = "june", NativeSlideId = 412,
                ExpectedPageNumber = 1, PageOrdinal = 0,
                RenderFingerprint = "sha256:rendered",
                NativeStateFingerprint = "sha256:native"
            };
            var context = AnalysisReviewContract.Context(artifact, plan,
                new[] { page });
            var request = AnalysisReviewContract.PrepareRequest(artifact, plan,
                context, new AnalysisRepairBudget().Serialize());
            Check(request.Content.Contains(context.ContextId) &&
                request.Content.Contains(fact.FactId) &&
                AnalysisRepairBudget.Read(request.BudgetReceipt).ModelCalls == 1,
                "The bounded review request lost its context or call receipt.");
            var checkpointRoot = Path.Combine(Path.GetTempPath(),
                "scribble-analysis-review-" + Guid.NewGuid().ToString("N"));
            try
            {
                var input = new ChatCompletionRequest
                {
                    model = "offline-test",
                    messages = new List<object>
                    {
                        new ChatCompletionInputMessage
                        { role = "user", content = "Review the analysis draft" }
                    }
                };
                var store = new TaskCheckpointStore(checkpointRoot);
                var task = new TaskContextManager(input, "excel",
                    "Review the analysis draft", store);
                task.PersistAnalysis(artifact);
                var first = task.ReserveAnalysisReview(artifact, plan,
                    context, true);
                var resumed = new TaskContextManager(input, "excel",
                    task.State.Objective, store, store.Load(task.State.Id));
                var second = resumed.ReserveAnalysisReview(artifact, plan,
                    context, true);
                Check(AnalysisRepairBudget.Read(first.BudgetReceipt).ModelCalls == 1 &&
                    AnalysisRepairBudget.Read(second.BudgetReceipt).ModelCalls == 2 &&
                    resumed.State.HostData["analysis_repair_budget"] ==
                        second.BudgetReceipt,
                    "A resumed task reset or failed to persist the review call budget.");
                var defect = new AnalysisReviewMeasurement
                {
                    MeasurementId = "folio-412", Code = "PAGE_NUMBER",
                    LogicalSlideId = "june", NativeSlideId = 412,
                    TargetId = "shape:10", Observed = "2", Expected = "1"
                };
                var reservedPatch = resumed.ReserveAnalysisPatch(page,
                    defect, true);
                resumed = new TaskContextManager(input, "excel",
                    task.State.Objective, store, store.Load(task.State.Id));
                Check(AnalysisRepairBudget.Read(reservedPatch.BudgetReceipt)
                    .PatchedTargets
                    .SequenceEqual(new[] { "june/shape:10" }) &&
                    resumed.State.HostData["analysis_repair_budget"] ==
                        reservedPatch.BudgetReceipt &&
                    resumed.State.HostData.ContainsKey(
                        "analysis_pending_patch"),
                    "A resumed task lost its write-ahead patch reservation.");
                Reject(() => resumed.ReserveAnalysisPatch(page,
                    defect, true), "REPAIR_PENDING_RECONCILIATION");
                Reject(() => resumed.ReserveAnalysisReview(artifact, plan,
                    context, true), "REPAIR_PENDING_RECONCILIATION");
                var savedPage = new AnalysisReviewPage
                {
                    LogicalSlideId = "june", NativeSlideId = 412,
                    ExpectedPageNumber = 1, PageOrdinal = 0,
                    RenderFingerprint = "sha256:repaired",
                    NativeStateFingerprint = "sha256:repaired-native"
                };
                Reject(() => resumed.ReconcileAnalysisPatch(reservedPatch,
                    page, new[] { defect }), "REPAIR_PENDING_RECONCILIATION");
                Reject(() => resumed.ReconcileAnalysisPatch(reservedPatch,
                    savedPage, new[] { defect }),
                    "REPAIR_PENDING_RECONCILIATION");
                Reject(() => resumed.ReconcileAnalysisPatch(reservedPatch,
                    savedPage, new[] { new AnalysisReviewMeasurement
                    {
                        MeasurementId = "different-defect", Code =
                            "OUT_OF_BOUNDS", LogicalSlideId = "june",
                        NativeSlideId = 412, TargetId = "shape:10"
                    } }), "REPAIR_PENDING_RECONCILIATION");
                resumed.ReconcileAnalysisPatch(reservedPatch, savedPage,
                    new AnalysisReviewMeasurement[0]);
                resumed = new TaskContextManager(input, "excel",
                    task.State.Objective, store, store.Load(task.State.Id));
                Check(!resumed.State.HostData.ContainsKey(
                    "analysis_pending_patch"),
                    "A completed native repair remained pending after resume.");
                Reject(() => resumed.ReserveAnalysisPatch(page,
                    defect, true), "REPAIR_TARGET_ALREADY_PATCHED");
                Reject(() => resumed.ReserveAnalysisReview(artifact, plan,
                    context, false), "REPAIR_BUDGET_TASK_MODE_CHANGED");
                resumed.State.HostData["analysis_repair_budget"] = "";
                Reject(() => resumed.ReserveAnalysisReview(artifact, plan,
                    context, true), "REPAIR_BUDGET_RECEIPT_INVALID");
            }
            finally
            {
                if (Directory.Exists(checkpointRoot))
                    Directory.Delete(checkpointRoot, true);
            }
            var json = new JavaScriptSerializer();
            Func<bool, object[], string> verdict = (approved, findings) =>
                json.Serialize(new Dictionary<string, object>
                {
                    { "contract_version", AnalysisReviewContract.Version },
                    { "context_id", context.ContextId },
                    { "approved", approved }, { "findings", findings }
                });
            var clean = AnalysisReviewContract.Parse(verdict(true,
                new object[0]), context);
            Check(clean.Approved, "A clean review was rejected.");
            var oldApproval = verdict(true, new object[0]);
            page.NativeStateFingerprint = "sha256:changed-native-state";
            var changedNative = AnalysisReviewContract.Context(artifact, plan,
                new[] { page });
            Reject(() => AnalysisReviewContract.Parse(oldApproval, changedNative),
                "REVIEW_CONTEXT_CHANGED");
            page.NativeStateFingerprint = "sha256:native";
            page.ExpectedPageNumber = 2;
            Reject(() => AnalysisReviewContract.Context(artifact, plan,
                new[] { page }), "REVIEW_PAGE_METADATA_INVALID");
            page.ExpectedPageNumber = 1;
            var claim = Finding("UNSUPPORTED_CLAIM", "content", "june",
                412, "title", "", "", "blocker", "revise_text",
                "The title implies a wider audit than this ledger supports.");
            var claimReview = verdict(false, new object[] { claim });
            var claimDecision = AnalysisReviewContract.Parse(claimReview,
                context);
            var patchRequest = AnalysisDocumentRepair.PreparePatchRequest(
                artifact, plan, context, claimDecision,
                claimDecision.Findings.Single());
            Check(patchRequest.Content.Contains("June revenue") &&
                patchRequest.Content.Contains(context.ContextId) &&
                patchRequest.MaxResponseTokens <= 512,
                "The content patch prompt lost its bound literal or context.");
            var patchJson = json.Serialize(new Dictionary<string, object>
            {
                { "context_id", context.ContextId },
                { "logical_slide_id", "june" },
                { "target_id", "title" },
                { "segment_index", 0 },
                { "expected_text", "June revenue" },
                { "replacement_text", "Verified June revenue" }
            });
            var parsedPatch = AnalysisDocumentRepair.ParsePatch(patchJson,
                context, claimDecision.Findings.Single());
            Check(parsedPatch.ReplacementText == "Verified June revenue",
                "The exact content patch did not parse.");
            Reject(() => AnalysisDocumentRepair.ParsePatch(
                patchJson.Replace("\"title\"", "\"chart\""),
                context, claimDecision.Findings.Single()),
                "REPAIR_PATCH_SCHEMA_INVALID");
            var titlePatch = new AnalysisDocumentPatch
            {
                ContextId = context.ContextId, LogicalSlideId = "june",
                TargetId = "title", SegmentIndex = 0,
                ExpectedText = "June revenue",
                ReplacementText = "Verified June revenue"
            };
            var repaired = AnalysisDocumentRepair.Apply(artifact, plan,
                context, claimReview, titlePatch,
                new AnalysisRepairBudget().Serialize());
            Check(plan.Slides[0].Title == "June revenue" &&
                repaired.Plan.Slides[0].Title == "Verified June revenue" &&
                repaired.Plan.Slides[0].Cards[0].Points[0].FactId == fact.FactId,
                "A one-field repair altered its source plan or fact reference.");
            var hierarchy = Finding("VISUAL_HIERARCHY", "content", "june",
                412, "page", "", "", "blocker", "revise_layout",
                "The KPI needs a clearer hierarchy.");
            var hierarchyReview = verdict(false, new object[] { hierarchy });
            var hierarchyDecision = AnalysisReviewContract.Parse(
                hierarchyReview, context);
            var layoutRequest = AnalysisDocumentRepair.PreparePatchRequest(
                artifact, plan, context, hierarchyDecision,
                hierarchyDecision.Findings.Single());
            Check(layoutRequest.Instructions ==
                    AnalysisDocumentRepair.LayoutPatchInstructions &&
                layoutRequest.Content.Contains("\"current_layout\":\"scorecard\"") &&
                layoutRequest.Content.Contains("\"target_id\":\"page\"") &&
                layoutRequest.Content.Contains(fact.FactId) &&
                layoutRequest.MaxResponseTokens <= 512,
                "The bounded layout prompt lost its current layout or evidence.");
            var layoutPatchJson = json.Serialize(new Dictionary<string, object>
            {
                { "context_id", context.ContextId },
                { "logical_slide_id", "june" },
                { "target_id", "page" },
                { "segment_index", 0 },
                { "expected_text", "scorecard" },
                { "replacement_text", "cards" }
            });
            var layoutPatch = AnalysisDocumentRepair.ParsePatch(layoutPatchJson,
                context, hierarchyDecision.Findings.Single());
            var reflowed = AnalysisDocumentRepair.Apply(artifact, plan,
                context, hierarchyReview, layoutPatch,
                new AnalysisRepairBudget().Serialize());
            Check(plan.Slides[0].Layout == "scorecard" &&
                reflowed.Plan.Slides[0].Layout == "cards" &&
                reflowed.Plan.Slides[0].Cards[0].Points[0].FactId == fact.FactId &&
                reflowed.Plan.Slides[0].Title == plan.Slides[0].Title,
                "A layout repair changed its source plan, text, or fact reference.");
            layoutPatch.ReplacementText = "invented_layout";
            Reject(() => AnalysisDocumentRepair.Apply(artifact, plan,
                context, hierarchyReview, layoutPatch,
                new AnalysisRepairBudget().Serialize()),
                "REPAIR_LAYOUT_INVALID");
            layoutPatch.ReplacementText = "cards";
            layoutPatch.ExpectedText = "stale_layout";
            Reject(() => AnalysisDocumentRepair.Apply(artifact, plan,
                context, hierarchyReview, layoutPatch,
                new AnalysisRepairBudget().Serialize()),
                "REPAIR_LAYOUT_INVALID");
            var cardPlan = json.Deserialize<AnalysisDocumentPlan>(
                json.Serialize(plan));
            cardPlan.Slides[0].Cards[0].Points.Add(
                new AnalysisPlanText { Text =
                    "The regional mix is complete" });
            var cardContext = AnalysisReviewContract.Context(artifact,
                cardPlan, new[] { page });
            var cardClaim = Finding("UNSUPPORTED_CLAIM", "content",
                "june", 412, "cards[0]", "", "", "blocker",
                "revise_text", "The card overstates the evidence.");
            var cardReview = verdict(false, new object[] { cardClaim })
                .Replace(context.ContextId, cardContext.ContextId);
            var cardDecision = AnalysisReviewContract.Parse(cardReview,
                cardContext);
            var cardRequest = AnalysisDocumentRepair.PreparePatchRequest(
                artifact, cardPlan, cardContext, cardDecision,
                cardDecision.Findings.Single());
            Check(cardRequest.Content.Contains(
                    "The regional mix is complete"),
                "The card literal was not available for bounded repair.");
            var cardPatch = new AnalysisDocumentPatch
            {
                ContextId = cardContext.ContextId,
                LogicalSlideId = "june", TargetId = "cards[0]",
                SegmentIndex = 1,
                ExpectedText = "The regional mix is complete",
                ReplacementText = "The shown mix is limited to June"
            };
            var repairedCard = AnalysisDocumentRepair.Apply(artifact,
                cardPlan, cardContext, cardReview, cardPatch,
                new AnalysisRepairBudget().Serialize());
            Check(repairedCard.Plan.Slides[0].Cards[0].Points[0].FactId ==
                    fact.FactId &&
                repairedCard.Plan.Slides[0].Cards[0].Points[1].Text ==
                    "The shown mix is limited to June" &&
                cardPlan.Slides[0].Cards[0].Points[1].Text ==
                    "The regional mix is complete",
                "A card-body repair changed a verified fact or its source plan.");
            var priorPilot = Environment.GetEnvironmentVariable(
                AnalysisDocumentPilot.FeatureFlag);
            try
            {
                Environment.SetEnvironmentVariable(
                    AnalysisDocumentPilot.FeatureFlag, "1");
                Check(AnalysisDocumentPilot.CompiledNativeText(artifact,
                        cardPlan, "june", "cards[0]") ==
                        "The regional mix is complete" &&
                    AnalysisDocumentPilot.CompiledNativeText(artifact,
                        repairedCard.Plan, "june", "cards[0]") ==
                        "The shown mix is limited to June",
                    "The exact native card text could not be bound.");
                cardPlan.Slides[0].Cards[0].Points.Add(
                    new AnalysisPlanText { Text = "A second literal" });
                Reject(() => AnalysisDocumentPilot.CompiledNativeText(
                    artifact, cardPlan, "june", "cards[0]"),
                    "REPAIR_NATIVE_TEXT_TARGET_UNSUPPORTED");
            }
            finally
            {
                Environment.SetEnvironmentVariable(
                    AnalysisDocumentPilot.FeatureFlag, priorPilot);
            }
            var contentCheckpoint = Path.Combine(Path.GetTempPath(),
                "scribble-analysis-content-" + Guid.NewGuid().ToString("N"));
            try
            {
                var contentInput = new ChatCompletionRequest
                {
                    model = "offline-test",
                    messages = new List<object>
                    {
                        new ChatCompletionInputMessage
                        { role = "user", content = "Correct the draft" }
                    }
                };
                var contentStore = new TaskCheckpointStore(
                    contentCheckpoint);
                var contentTask = new TaskContextManager(contentInput,
                    "excel", "Correct the draft", contentStore);
                contentTask.PersistAnalysis(artifact);
                var reservedCall = contentTask
                    .ReserveAnalysisContentPatchRequest(artifact, plan,
                        context, claimDecision,
                        claimDecision.Findings.Single(), true);
                var desired = AnalysisDocumentRepair.Apply(artifact, plan,
                    context, claimReview, titlePatch,
                    reservedCall.BudgetReceipt);
                var contentReservation = contentTask
                    .ReserveAnalysisContentPatch(page, titlePatch,
                        desired.Plan, "June revenue",
                        "Verified June revenue", "input-fingerprint",
                        desired.BudgetReceipt, true);
                var resumedContent = new TaskContextManager(contentInput,
                    "excel", contentTask.State.Objective, contentStore,
                    contentStore.Load(contentTask.State.Id));
                Reject(() => resumedContent.ReserveAnalysisReview(
                    artifact, plan, context, true),
                    "REPAIR_PENDING_RECONCILIATION");
                var changedPage = new AnalysisReviewPage
                {
                    LogicalSlideId = page.LogicalSlideId,
                    NativeSlideId = page.NativeSlideId,
                    ExpectedPageNumber = page.ExpectedPageNumber,
                    PageOrdinal = page.PageOrdinal,
                    RenderFingerprint = "sha256:corrected",
                    NativeStateFingerprint = "sha256:corrected-native"
                };
                Reject(() => resumedContent.ReconcileAnalysisContentPatch(
                    contentReservation, changedPage, "wrong text"),
                    "REPAIR_PENDING_RECONCILIATION");
                resumedContent.ReconcileAnalysisContentPatch(
                    contentReservation, changedPage,
                    "Verified June revenue");
                var resumedPlan = json.Deserialize<AnalysisDocumentPlan>(
                    resumedContent.State.HostData["analysis_deck_plan"]);
                Check(resumedPlan.Slides[0].Title ==
                        "Verified June revenue" &&
                    !resumedContent.State.HostData.ContainsKey(
                        "analysis_pending_content_patch") &&
                    AnalysisRepairBudget.Read(resumedContent.State.HostData[
                        "analysis_repair_budget"]).PatchedTargets.Contains(
                            "june/title"),
                    "Content repair lost its durable plan or patch receipt.");
            }
            finally
            {
                if (Directory.Exists(contentCheckpoint))
                    Directory.Delete(contentCheckpoint, true);
            }
            Reject(() => AnalysisDocumentRepair.Apply(artifact, plan,
                context, claimReview, titlePatch, repaired.BudgetReceipt),
                "REPAIR_TARGET_ALREADY_PATCHED");
            titlePatch.ReplacementText = "Revenue 9";
            Reject(() => AnalysisDocumentRepair.Apply(artifact, plan,
                context, claimReview, titlePatch,
                new AnalysisRepairBudget().Serialize()),
                "ANALYSIS_NUMERIC_LITERAL_UNVERIFIED");
            titlePatch.ReplacementText = "Verified June revenue";
            plan.Slides[0].Title = "Changed after review";
            Reject(() => AnalysisDocumentRepair.Apply(artifact, plan,
                context, claimReview, titlePatch,
                new AnalysisRepairBudget().Serialize()),
                "REPAIR_CONTEXT_CHANGED");
            plan.Slides[0].Title = "June revenue";
            var binding = Finding("BINDING_CHALLENGE", "analysis", "june",
                412, "cards[0]", fact.FactId, "", "blocker",
                "inspect_binding", "Check the Revenue label against Ledger column I.");
            var challenged = AnalysisReviewContract.Parse(verdict(false,
                new object[] { binding }), context);
            Check(!challenged.Approved && challenged.Findings.Single().FactId ==
                fact.FactId, "A binding challenge was not routed to inspection.");
            binding["evidence"] = "The reviewer incorrectly claims 82,992 should be 85,519.";
            var falseArithmetic = AnalysisReviewContract.Parse(verdict(false,
                new object[] { binding }), context);
            Check(falseArithmetic.Findings.Single().Action ==
                    "inspect_binding" && fact.Value == "82992",
                "A false arithmetic objection overrode the verified fact.");
            binding["evidence"] = "Check the Revenue label against Ledger column I.";
            Reject(() => AnalysisReviewContract.Parse(verdict(true,
                new object[] { binding }), context), "REVIEW_VERDICT_CONTRADICTORY");
            Reject(() => AnalysisReviewContract.Parse(verdict(false,
                new object[0]), context), "REVIEW_VERDICT_CONTRADICTORY");
            binding["replacement_value"] = "85519";
            Reject(() => AnalysisReviewContract.Parse(verdict(false,
                new object[] { binding }), context), "REVIEW_FINDING_SCHEMA_INVALID");
            binding.Remove("replacement_value");
            binding["fact_id"] = "model-invented-fact";
            Reject(() => AnalysisReviewContract.Parse(verdict(false,
                new object[] { binding }), context), "REVIEW_FACT_BINDING_INVALID");
            binding["fact_id"] = fact.FactId;
            binding["native_slide_id"] = 1;
            Reject(() => AnalysisReviewContract.Parse(verdict(false,
                new object[] { binding }), context), "REVIEW_FINDING_PAGE_INVALID");

            var collision = new AnalysisReviewMeasurement
            {
                MeasurementId = "geometry-1", Code = "COLLISION",
                LogicalSlideId = "june", NativeSlideId = 412,
                TargetId = "shape:10", OtherTargetId = "shape:9",
                Observed = "chart and callout overlap by 12 px",
                Expected = "no overlap"
            };
            context = AnalysisReviewContract.Context(artifact, plan,
                new[] { page }, new[] { collision });
            Reject(() => AnalysisReviewContract.Parse(oldApproval, context),
                "REVIEW_CONTEXT_CHANGED");
            Reject(() => AnalysisReviewContract.Parse(verdict(true,
                new object[0]), context), "REVIEW_VERDICT_CONTRADICTORY");
            var hostOwned = AnalysisReviewContract.Parse(verdict(false,
                new object[0]), context);
            Check(hostOwned.Findings.Count == 1 &&
                hostOwned.Findings[0].Owner == "renderer" &&
                hostOwned.Findings[0].MeasurementId == "geometry-1",
                "The host geometry measurement was lost when the model omitted it.");
            var geometry = Finding("COLLISION", "renderer", "june", 412,
                "shape:10", "", "geometry-1", "blocker", "adjust_layout",
                "Chart overlaps the callout.");
            Check(!AnalysisReviewContract.Parse(verdict(false,
                new object[] { geometry }), context).Approved,
                "Measured native collision was silently approved.");
            geometry["measurement_id"] = "invented-measurement";
            Reject(() => AnalysisReviewContract.Parse(verdict(false,
                new object[] { geometry }), context),
                "REVIEW_RENDERER_MEASUREMENT_REQUIRED");
        }

        private static Dictionary<string, object> Finding(string code, string owner,
            string slide, int native, string target, string fact, string measurement,
            string severity, string action, string evidence)
        {
            return new Dictionary<string, object>
            {
                { "code", code }, { "owner", owner },
                { "logical_slide_id", slide }, { "native_slide_id", native },
                { "target_id", target }, { "fact_id", fact },
                { "measurement_id", measurement }, { "severity", severity },
                { "action", action }, { "evidence", evidence }
            };
        }

        private static void Reject(Action action, string code)
        {
            try { action(); }
            catch (InvalidOperationException error)
            { if (error.Message.Contains(code)) return; }
            throw new Exception("Expected rejection: " + code);
        }

        private static void Check(bool success, string message)
        { if (!success) throw new Exception(message); }
    }
}
