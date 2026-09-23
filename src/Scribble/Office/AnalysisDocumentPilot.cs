using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace Scribble.Office
{
    // Development-only bridge to the existing Office writers. A caller must
    // provide a new, disposable destination; no model-facing tool uses this
    // bridge until the native, review and recovery gates have passed.
    public static class AnalysisDocumentPilot
    {
        public const string FeatureFlag = "SCRIBBLE_ANALYSIS_PILOT";

        public static string WriteWorkbook(object excelApplication,
            AnalysisArtifact artifact, AnalysisDocumentPlan plan)
        {
            RequireEnabled();
            var compiled = AnalysisDocumentCompiler.Compile(artifact, plan);
            dynamic application = excelApplication;
            dynamic workbook = application.ActiveWorkbook;
            if (workbook == null)
                throw new InvalidOperationException(
                    "ANALYSIS_PILOT_DESTINATION_MISSING: Open a disposable source workbook first.");
            string[] before = SheetNames((object)workbook);
            var rows = compiled.WorkbookRows.Select(row =>
                (IReadOnlyList<string>)row).ToList();
            var status = WorkbookDraftWriter.WriteDraftSheet(excelApplication,
                compiled.WorkbookTitle, rows, null, false, (object)workbook);
            string[] after = SheetNames((object)workbook);
            var added = after.Except(before,
                StringComparer.OrdinalIgnoreCase).ToArray();
            if (added.Length != 1 || !added[0].StartsWith(
                WorkbookDraftWriter.DraftSheetName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "ANALYSIS_PILOT_DRAFT_IDENTITY_INVALID: Expected one new marked draft sheet.");
            dynamic sheet = workbook.Worksheets[added[0]];
            var facts = artifact.Facts.ToDictionary(fact => fact.FactId,
                StringComparer.Ordinal);
            foreach (var expected in compiled.ExpectedFormulaFacts)
            {
                dynamic cell = sheet.Range(expected.Key);
                var formula = Convert.ToString(cell.Formula,
                    CultureInfo.InvariantCulture) ?? string.Empty;
                object native = cell.Value2;
                decimal actual = 0m;
                if (!formula.StartsWith("=", StringComparison.Ordinal) ||
                    native == null ||
                    !decimal.TryParse(Convert.ToString(native,
                        CultureInfo.InvariantCulture), NumberStyles.Float,
                        CultureInfo.InvariantCulture, out actual))
                    throw new InvalidOperationException(
                        "ANALYSIS_PILOT_FORMULA_UNRESOLVED: " + expected.Key);
                var wanted = AnalysisContract.Decimal(facts[expected.Value]);
                var tolerance = facts[expected.Value].Unit == "currency"
                    ? 0.005m : 0.000001m;
                if (Math.Abs(actual - wanted) > tolerance)
                    throw new InvalidOperationException(
                        "ANALYSIS_PILOT_FORMULA_MISMATCH: " + expected.Key +
                        " expected " + wanted.ToString(CultureInfo.InvariantCulture) +
                        " but Excel returned " + actual.ToString(CultureInfo.InvariantCulture));
                // Excel's generic optional-decimal format can print a bare
                // trailing separator in PDF exports. Use a definite native
                // format only after the fact has passed readback.
                var scale = (decimal.GetBits(wanted)[3] >> 16) & 0xff;
                cell.NumberFormat = scale == 0 ? "#,##0" :
                    scale <= 2 ? "#,##0.00" : "General";
            }
            return status + " Verified " + compiled.ExpectedFormulaFacts.Count +
                " live formula result(s) against the analysis.";
        }

        public static string WritePresentation(object powerPointApplication,
            AnalysisArtifact artifact, AnalysisDocumentPlan plan)
        {
            RequireEnabled();
            var compiled = AnalysisDocumentCompiler.Compile(artifact, plan);
            var slides = PresentationDraftWriter.ParseSlides(
                compiled.Slides.Cast<object>().ToArray());
            return PresentationDraftWriter.AddDraftSlides(powerPointApplication,
                slides, null, true);
        }

        // After native write, the caller supplies actual slide identities and
        // render fingerprints. The typed verdict is bound to exactly that
        // analysis and page set; this method performs no model inference.
        public static AnalysisReviewDecision ReviewPresentation(
            AnalysisArtifact artifact, AnalysisDocumentPlan plan,
            IEnumerable<AnalysisReviewPage> pages,
            IEnumerable<AnalysisReviewMeasurement> measurements,
            string reviewerJson)
        {
            RequireEnabled();
            var context = AnalysisReviewContract.Context(artifact, plan,
                pages, measurements);
            return AnalysisReviewContract.Parse(reviewerJson, context);
        }

        public static IReadOnlyList<AnalysisReviewPage> CapturePresentationPages(
            object presentation, AnalysisArtifact artifact,
            AnalysisDocumentPlan plan)
        {
            RequireEnabled();
            var compiled = AnalysisDocumentCompiler.Compile(artifact, plan);
            var slides = PresentationDraftWriter.ParseSlides(
                compiled.Slides.Cast<object>().ToArray());
            var composed = PresentationDraftWriter.ComposeSamsung(slides);
            dynamic deck = presentation;
            if (deck == null || (int)deck.Slides.Count != composed.Count)
                throw new InvalidOperationException(
                    "ANALYSIS_PILOT_NATIVE_PAGE_COUNT_CHANGED");
            var pages = new List<AnalysisReviewPage>();
            var ordinals = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var index = 1; index <= composed.Count; index++)
            {
                dynamic native = deck.Slides[index];
                var logicalId = composed[index - 1].Source.Id;
                int ordinal;
                ordinals.TryGetValue(logicalId, out ordinal);
                ordinals[logicalId] = ordinal + 1;
                pages.Add(new AnalysisReviewPage
                {
                    LogicalSlideId = logicalId,
                    NativeSlideId = (int)native.SlideID,
                    ExpectedPageNumber = index,
                    PageOrdinal = ordinal,
                    RenderFingerprint = RenderFingerprint(native)
                });
            }
            return pages;
        }

        public static IReadOnlyList<AnalysisReviewMeasurement> CaptureNativeMeasurements(
            object presentation, IEnumerable<AnalysisReviewPage> reviewPages)
        {
            RequireEnabled();
            dynamic deck = presentation;
            var pages = (reviewPages ?? Enumerable.Empty<AnalysisReviewPage>())
                .OrderBy(page => page.ExpectedPageNumber).ToArray();
            if (deck == null || (int)deck.Slides.Count != pages.Length ||
                pages.Select(page => page.ExpectedPageNumber)
                    .Where((number, index) => number != index + 1).Any())
                throw new InvalidOperationException(
                    "ANALYSIS_PILOT_NATIVE_PAGE_COUNT_CHANGED");
            var findings = new List<AnalysisReviewMeasurement>();
            for (var index = 1; index <= pages.Length; index++)
            {
                dynamic slide = deck.Slides[index];
                var page = pages[index - 1];
                if ((int)slide.SlideID != page.NativeSlideId)
                    throw new InvalidOperationException(
                        "ANALYSIS_PILOT_NATIVE_PAGE_CHANGED");
                var width = (double)deck.PageSetup.SlideWidth;
                var height = (double)deck.PageSetup.SlideHeight;
                var pageNumbers = new List<string>();
                for (var shapeIndex = 1; shapeIndex <= (int)slide.Shapes.Count;
                    shapeIndex++)
                {
                    dynamic shape = slide.Shapes[shapeIndex];
                    var target = "shape:" + (int)shape.Id;
                    if ((double)shape.Left < -.5 || (double)shape.Top < -.5 ||
                        (double)shape.Left + (double)shape.Width > width + .5 ||
                        (double)shape.Top + (double)shape.Height > height + .5)
                        findings.Add(Measure("OUT_OF_BOUNDS", page, target,
                            "shape exceeds slide canvas", "inside slide canvas"));
                    if ((int)shape.HasTextFrame == 0 ||
                        (int)shape.HasTable != 0 || (int)shape.HasChart != 0)
                        continue;
                    dynamic range = shape.TextFrame.TextRange;
                    var text = Convert.ToString(range.Text) ?? string.Empty;
                    if (Regex.IsMatch(text.Trim(), @"^-\s*\d+\s*-$"))
                        pageNumbers.Add(text.Trim());
                    if (PresentationRevision.NativeTextOverflows(text,
                        (float)range.BoundHeight, (float)range.BoundWidth,
                        (float)shape.Height, (float)shape.Width))
                        findings.Add(Measure("TEXT_OVERFLOW", page, target,
                            "text bounds exceed native shape", "text fits native shape"));
                }
                var expected = "- " + page.ExpectedPageNumber + " -";
                if (pageNumbers.Count != 1 || pageNumbers[0] != expected)
                    findings.Add(Measure("PAGE_NUMBER", page, "page",
                        string.Join("; ", pageNumbers), expected));
            }
            return findings;
        }

        // Renderer-owned fixes are finite native operations. They neither
        // ask the model to rewrite the slide nor modify a bound fact.
        public static void RepairNativeMeasurement(object presentation,
            IReadOnlyList<AnalysisReviewPage> pages,
            AnalysisReviewMeasurement measurement)
        {
            RequireEnabled();
            if (measurement == null || pages == null)
                throw new InvalidOperationException("RENDERER_REPAIR_TARGET_INVALID");
            var page = pages.SingleOrDefault(item =>
                item.NativeSlideId == measurement.NativeSlideId &&
                item.LogicalSlideId == measurement.LogicalSlideId);
            if (page == null || page.ExpectedPageNumber < 1)
                throw new InvalidOperationException("RENDERER_REPAIR_TARGET_INVALID");
            dynamic deck = presentation;
            dynamic slide = deck.Slides[page.ExpectedPageNumber];
            if ((int)slide.SlideID != page.NativeSlideId ||
                RenderFingerprint(slide) != page.RenderFingerprint)
                throw new InvalidOperationException("RENDERER_REPAIR_PAGE_CHANGED");
            var current = CaptureNativeMeasurements(presentation, pages)
                .SingleOrDefault(item => item.MeasurementId ==
                    measurement.MeasurementId);
            if (current == null || current.Code != measurement.Code ||
                current.TargetId != measurement.TargetId ||
                current.Observed != measurement.Observed ||
                current.Expected != measurement.Expected)
                throw new InvalidOperationException("RENDERER_REPAIR_MEASUREMENT_CHANGED");
            if (measurement.Code == "PAGE_NUMBER")
            {
                var candidates = new List<object>();
                for (var index = 1; index <= (int)slide.Shapes.Count; index++)
                {
                    dynamic shape = slide.Shapes[index];
                    if ((int)shape.HasTextFrame == 0) continue;
                    var text = (Convert.ToString(shape.TextFrame.TextRange.Text) ??
                        string.Empty).Trim();
                    if (Regex.IsMatch(text, @"^-\s*\d+\s*-$"))
                        candidates.Add((object)shape);
                }
                if (candidates.Count != 1)
                    throw new InvalidOperationException(
                        "RENDERER_PAGE_NUMBER_UNSUPPORTED: Expected one native folio shape.");
                dynamic target = candidates[0];
                var before = Convert.ToString(target.TextFrame.TextRange.Text);
                try
                {
                    target.TextFrame.TextRange.Text = current.Expected;
                    if ((Convert.ToString(target.TextFrame.TextRange.Text) ??
                        string.Empty).Trim() != current.Expected)
                        throw new InvalidOperationException(
                            "RENDERER_PAGE_NUMBER_READBACK_FAILED");
                }
                catch
                {
                    try { target.TextFrame.TextRange.Text = before; }
                    catch { throw new InvalidOperationException(
                        "RENDERER_REPAIR_RECOVERY_REQUIRED"); }
                    throw;
                }
                return;
            }
            if (measurement.Code == "TEXT_OVERFLOW")
            {
                if (!measurement.TargetId.StartsWith("shape:",
                    StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "RENDERER_REPAIR_TARGET_INVALID");
                int shapeId;
                if (!int.TryParse(measurement.TargetId.Substring(6),
                    out shapeId))
                    throw new InvalidOperationException(
                        "RENDERER_REPAIR_TARGET_INVALID");
                dynamic target = null;
                for (var index = 1; index <= (int)slide.Shapes.Count; index++)
                    if ((int)slide.Shapes[index].Id == shapeId)
                        target = slide.Shapes[index];
                if (target == null || (int)target.HasTextFrame == 0)
                    throw new InvalidOperationException(
                        "RENDERER_REPAIR_TARGET_INVALID");
                dynamic range = target.TextFrame.TextRange;
                var before = (float)range.Font.Size;
                var text = Convert.ToString(range.Text);
                var minimum = (double)target.Top >
                    (double)deck.PageSetup.SlideHeight * .9 ? 9f :
                    SamsungSlideDesign.BodyMinimum;
                try
                {
                    var fitted = false;
                    for (var size = before - .5f; size >= minimum;
                        size -= .5f)
                    {
                        range.Font.Size = size;
                        if (!PresentationRevision.NativeTextOverflows(text,
                            (float)range.BoundHeight, (float)range.BoundWidth,
                            (float)target.Height, (float)target.Width))
                        { fitted = true; break; }
                    }
                    if (!fitted || Convert.ToString(range.Text) != text)
                        throw new InvalidOperationException(
                            "RENDERER_TEXT_FIT_UNSUPPORTED: The text cannot fit above its minimum readable size.");
                }
                catch
                {
                    try { range.Font.Size = before; }
                    catch { throw new InvalidOperationException(
                        "RENDERER_REPAIR_RECOVERY_REQUIRED"); }
                    throw;
                }
                return;
            }
            throw new InvalidOperationException(
                "RENDERER_REPAIR_UNSUPPORTED: " + measurement.Code);
        }

        private static string RenderFingerprint(dynamic slide)
        {
            var temporary = Path.Combine(Path.GetTempPath(),
                "scribble-analysis-review-" + Guid.NewGuid().ToString("N") +
                ".png");
            try
            {
                slide.Export(temporary, "PNG", 1600, 900);
                using (var digest = SHA256.Create())
                    return BitConverter.ToString(digest.ComputeHash(
                        File.ReadAllBytes(temporary))).Replace("-", "")
                        .ToLowerInvariant();
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        private static AnalysisReviewMeasurement Measure(string code,
            AnalysisReviewPage page, string target, string observed,
            string expected)
        {
            return new AnalysisReviewMeasurement
            {
                MeasurementId = "native:" + page.NativeSlideId + ":" +
                    code + ":" + target,
                Code = code, LogicalSlideId = page.LogicalSlideId,
                NativeSlideId = page.NativeSlideId, TargetId = target,
                Observed = observed, Expected = expected
            };
        }

        private static void RequireEnabled()
        {
            if (!string.Equals(Environment.GetEnvironmentVariable(FeatureFlag),
                "1", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "ANALYSIS_PILOT_DISABLED: The analysis writer is restricted to the development pilot.");
        }

        private static string[] SheetNames(dynamic workbook)
        {
            var names = new List<string>();
            foreach (dynamic sheet in workbook.Worksheets)
                names.Add(Convert.ToString(sheet.Name,
                    CultureInfo.InvariantCulture));
            return names.ToArray();
        }
    }
}
