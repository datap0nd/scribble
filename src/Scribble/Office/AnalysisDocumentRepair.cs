using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace Scribble.Office
{
    public sealed class AnalysisDocumentPatch
    {
        public string ContextId { get; set; }
        public string LogicalSlideId { get; set; }
        public string TargetId { get; set; }
        // A literal segment index; fact references cannot be selected here.
        public int SegmentIndex { get; set; }
        public string ExpectedText { get; set; }
        public string ReplacementText { get; set; }
    }

    public sealed class AnalysisDocumentRepairResult
    {
        public AnalysisDocumentPlan Plan { get; set; }
        public string BudgetReceipt { get; set; }
    }

    public static class AnalysisDocumentRepair
    {
        public static AnalysisDocumentRepairResult Apply(
            AnalysisArtifact artifact, AnalysisDocumentPlan original,
            AnalysisReviewContext context, string reviewerJson,
            AnalysisDocumentPatch patch, string budgetReceipt)
        {
            if (original == null || context == null ||
                AnalysisReviewContract.Context(artifact, original,
                    context.Pages, context.Measurements).ContextId !=
                    context.ContextId)
                throw new InvalidOperationException("REPAIR_CONTEXT_CHANGED");
            var verdict = AnalysisReviewContract.Parse(reviewerJson, context);
            if (verdict.Approved || patch == null ||
                patch.ContextId != context.ContextId ||
                patch.LogicalSlideId == null || patch.TargetId == null)
                throw new InvalidOperationException("REPAIR_FINDING_REQUIRED");
            var finding = verdict.Findings.FirstOrDefault(item =>
                item.Severity == "blocker" && item.Owner == "content" &&
                item.LogicalSlideId == patch.LogicalSlideId &&
                item.TargetId == patch.TargetId);
            if (finding == null)
                throw new InvalidOperationException("REPAIR_FINDING_REQUIRED");
            if (patch.ExpectedText == null || patch.ReplacementText == null ||
                patch.ExpectedText == patch.ReplacementText ||
                patch.ReplacementText.Length > 240)
                throw new InvalidOperationException("REPAIR_TEXT_INVALID");
            var json = new JavaScriptSerializer { MaxJsonLength = 16000000 };
            var plan = json.Deserialize<AnalysisDocumentPlan>(json.Serialize(original));
            var slide = plan.Slides.SingleOrDefault(item => item.Id ==
                patch.LogicalSlideId);
            if (slide == null)
                throw new InvalidOperationException("REPAIR_SLIDE_CHANGED");
            if (finding.Code == "VISUAL_HIERARCHY" &&
                finding.Action == "revise_layout" && patch.TargetId == "page")
            {
                if (patch.SegmentIndex != 0 || slide.Layout != patch.ExpectedText ||
                    !SamsungSlideDesign.Layouts.Contains(patch.ReplacementText))
                    throw new InvalidOperationException("REPAIR_LAYOUT_INVALID");
                slide.Layout = patch.ReplacementText;
            }
            else if (finding.Code == "UNSUPPORTED_CLAIM" &&
                finding.Action == "revise_text")
                ReplaceLiteral(slide, patch);
            else throw new InvalidOperationException("REPAIR_ACTION_UNSUPPORTED");
            // Recompile before consuming a task-level patch. This validates
            // the immutable fact references, native recipe and numeric prose.
            var compiled = AnalysisDocumentCompiler.Compile(artifact, plan);
            var composed = PresentationDraftWriter.ComposeSamsung(
                PresentationDraftWriter.ParseSlides(
                    compiled.Slides.Cast<object>().ToArray()));
            if (!composed.Select(page => page.Source.Id).SequenceEqual(
                context.Pages.OrderBy(page => page.ExpectedPageNumber)
                    .Select(page => page.LogicalSlideId)))
                throw new InvalidOperationException(
                    "REPAIR_PAGE_COUNT_CHANGED: The patch cannot replace the allocated native pages.");
            var budget = AnalysisRepairBudget.Read(budgetReceipt);
            var next = budget.ConsumePatch(patch.LogicalSlideId,
                patch.TargetId);
            return new AnalysisDocumentRepairResult
            {
                Plan = plan, BudgetReceipt = next
            };
        }

        private static void ReplaceLiteral(AnalysisPlanSlide slide,
            AnalysisDocumentPatch patch)
        {
            if (patch.TargetId == "title")
            {
                if (patch.SegmentIndex != 0 ||
                    slide.Title != patch.ExpectedText)
                    throw new InvalidOperationException("REPAIR_TEXT_CHANGED");
                slide.Title = patch.ReplacementText;
                return;
            }
            List<AnalysisPlanText> parts = null;
            if (patch.TargetId == "subtitle") parts = slide.Subtitle;
            else if (patch.TargetId == "takeaway") parts = slide.Takeaway;
            else if (patch.TargetId.StartsWith("cards[", StringComparison.Ordinal) &&
                patch.TargetId.EndsWith("]", StringComparison.Ordinal))
            {
                int card;
                if (!int.TryParse(patch.TargetId.Substring(6,
                    patch.TargetId.Length - 7), out card) || card < 0 ||
                    slide.Cards == null || card >= slide.Cards.Count)
                    throw new InvalidOperationException("REPAIR_TARGET_INVALID");
                parts = slide.Cards[card].Points;
            }
            if (parts == null || patch.SegmentIndex < 0 ||
                patch.SegmentIndex >= parts.Count ||
                parts[patch.SegmentIndex] == null ||
                parts[patch.SegmentIndex].FactId != null ||
                parts[patch.SegmentIndex].Text != patch.ExpectedText)
                throw new InvalidOperationException("REPAIR_TEXT_CHANGED");
            parts[patch.SegmentIndex].Text = patch.ReplacementText;
        }
    }
}
