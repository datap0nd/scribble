using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;
using Scribble.Office;

namespace GuardrailTests
{
    internal static class AnalysisReviewContractTests
    {
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
            var claim = Finding("UNSUPPORTED_CLAIM", "content", "june",
                412, "title", "", "", "blocker", "revise_text",
                "The title implies a wider audit than this ledger supports.");
            var claimReview = verdict(false, new object[] { claim });
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
                TargetId = "callout",
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
                "callout", "", "geometry-1", "blocker", "adjust_layout",
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
