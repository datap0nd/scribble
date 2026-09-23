using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace Scribble.Office
{
    // A write-ahead reservation identifies the exact native state that may be
    // changed. A task with a pending reservation must reconcile it on resume.
    public sealed class AnalysisPatchReservation
    {
        public string LogicalSlideId { get; set; }
        public int NativeSlideId { get; set; }
        public string TargetId { get; set; }
        public string MeasurementId { get; set; }
        public string NativeStateFingerprint { get; set; }
        public string BudgetReceipt { get; set; }

        public void Validate(AnalysisReviewPage page,
            AnalysisReviewMeasurement measurement)
        {
            if (page == null || measurement == null ||
                LogicalSlideId != page.LogicalSlideId ||
                NativeSlideId != page.NativeSlideId ||
                TargetId != measurement.TargetId ||
                MeasurementId != measurement.MeasurementId ||
                NativeStateFingerprint != page.NativeStateFingerprint ||
                measurement.LogicalSlideId != page.LogicalSlideId ||
                measurement.NativeSlideId != page.NativeSlideId ||
                !AnalysisRepairBudget.Read(BudgetReceipt).PatchedTargets.Contains(
                    LogicalSlideId + "/" + TargetId, StringComparer.Ordinal))
                throw new InvalidOperationException("REPAIR_RESERVATION_CHANGED");
        }
    }

    // One serialized task-level counter survives every review and repair stage.
    // Nested loops may consume it but cannot reset it.
    public sealed class AnalysisRepairBudget
    {
        public const int Version = 1;
        public const int MaxModelCalls = 18;
        // A six-slide repair can need one patch on every slide, with two
        // spare targets for a workbook or follow-up content defect.
        public const int MaxCorrectivePatches = 8;
        public const int MaxPromptCharacters = 36000;
        public const int MaxResponseTokens = 8192;

        public int ContractVersion { get; set; } = Version;
        public int CallLimit { get; set; } = MaxModelCalls;
        public int ModelCalls { get; set; }
        public List<string> PatchedTargets { get; set; } = new List<string>();

        public static AnalysisRepairBudget Read(string receipt)
        {
            if (string.IsNullOrWhiteSpace(receipt)) return new AnalysisRepairBudget();
            AnalysisRepairBudget budget;
            try { budget = new JavaScriptSerializer().Deserialize<AnalysisRepairBudget>(receipt); }
            catch (Exception error) when (error is ArgumentException ||
                error is InvalidOperationException)
            { throw new InvalidOperationException("REPAIR_BUDGET_RECEIPT_INVALID", error); }
            if (budget == null || budget.ContractVersion != Version ||
                (budget.CallLimit != 12 && budget.CallLimit != MaxModelCalls) ||
                budget.ModelCalls < 0 || budget.ModelCalls > budget.CallLimit ||
                budget.PatchedTargets == null ||
                budget.PatchedTargets.Count > MaxCorrectivePatches ||
                budget.PatchedTargets.Any(string.IsNullOrWhiteSpace) ||
                budget.PatchedTargets.Distinct(StringComparer.Ordinal).Count() !=
                    budget.PatchedTargets.Count)
                throw new InvalidOperationException("REPAIR_BUDGET_RECEIPT_INVALID");
            return budget;
        }

        public static AnalysisRepairBudget CrossApp()
        { return new AnalysisRepairBudget { CallLimit = 12 }; }

        public string ConsumeModelCall(int promptCharacters, int responseTokens)
        {
            if (promptCharacters < 0 || promptCharacters > MaxPromptCharacters ||
                responseTokens < 1 || responseTokens > MaxResponseTokens)
                throw new InvalidOperationException("REVIEW_CONTEXT_LIMIT: Shorten or split the request before inference.");
            if (ModelCalls >= CallLimit)
                throw new InvalidOperationException("REPAIR_MODEL_CALL_LIMIT: The task-level call budget is exhausted.");
            ModelCalls++;
            return Serialize();
        }

        public string ConsumePatch(string logicalSlideId, string targetId)
        {
            if (string.IsNullOrWhiteSpace(logicalSlideId) ||
                string.IsNullOrWhiteSpace(targetId))
                throw new InvalidOperationException("REPAIR_TARGET_REQUIRED");
            var key = logicalSlideId + "/" + targetId;
            if (PatchedTargets.Contains(key, StringComparer.Ordinal))
                throw new InvalidOperationException("REPAIR_TARGET_ALREADY_PATCHED: " + key);
            if (PatchedTargets.Count >= MaxCorrectivePatches)
                throw new InvalidOperationException("REPAIR_PATCH_LIMIT: The task-level patch budget is exhausted.");
            PatchedTargets.Add(key);
            return Serialize();
        }

        public string Serialize()
        {
            return new JavaScriptSerializer().Serialize(this);
        }
    }
}
