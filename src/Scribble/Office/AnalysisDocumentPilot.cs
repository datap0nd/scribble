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
                var temporary = Path.Combine(Path.GetTempPath(),
                    "scribble-analysis-review-" + Guid.NewGuid().ToString("N") +
                    ".png");
                string fingerprint;
                try
                {
                    native.Export(temporary, "PNG", 1600, 900);
                    using (var digest = SHA256.Create())
                        fingerprint = BitConverter.ToString(digest.ComputeHash(
                            File.ReadAllBytes(temporary))).Replace("-", "")
                            .ToLowerInvariant();
                }
                finally { if (File.Exists(temporary)) File.Delete(temporary); }
                pages.Add(new AnalysisReviewPage
                {
                    LogicalSlideId = logicalId,
                    NativeSlideId = (int)native.SlideID,
                    ExpectedPageNumber = index,
                    PageOrdinal = ordinal,
                    RenderFingerprint = fingerprint
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
