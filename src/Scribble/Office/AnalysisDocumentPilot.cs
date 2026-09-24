using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Scribble.Chat;

namespace Scribble.Office
{
    public sealed class AnalysisNativeReviewSession
    {
        public AnalysisReviewContext Context { get; set; }
        public AnalysisReviewRequest Request { get; set; }
        // Each page image is bound to its rendered fingerprint.
        public List<string> PageImages { get; set; } = new List<string>();
    }

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
            var compiled = AnalysisDocumentCompiler.Compile(artifact, plan,
                false);
            AnalysisWorkbookSourceGuard.Validate(excelApplication, artifact);
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
                slides, null, true, null, null, null, null, true);
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

        // Capture once: the exact rendered/native page state in the request is
        // the state against which the reviewer response must be parsed. The
        // task checkpoints the call budget before any model request is sent.
        public static AnalysisNativeReviewSession ReserveNativeReview(
            TaskContextManager task, object presentation,
            AnalysisArtifact artifact, AnalysisDocumentPlan plan,
            bool crossApp, int maxResponseTokens = 2048)
        {
            RequireEnabled();
            if (task == null)
                throw new InvalidOperationException("REVIEW_TASK_REQUIRED");
            var images = new List<string>();
            var pages = CapturePresentationPages(presentation, artifact,
                plan, images);
            if (images.Any(image => string.IsNullOrEmpty(image)))
                throw new InvalidOperationException(
                    "ANALYSIS_VISUAL_REVIEW_UNAVAILABLE: A native chart slide cannot be safely exported on this Office build. The draft remains pending for visual inspection; no model approval or review receipt was issued.");
            var measurements = CaptureNativeMeasurements(presentation,
                pages);
            if (measurements.Count != 0)
                throw new InvalidOperationException(
                    "ANALYSIS_DECK_GEOMETRY_UNRESOLVED: Native measurements must be repaired or explicitly declined before model review.");
            var context = AnalysisReviewContract.Context(artifact, plan,
                pages, measurements);
            var request = task.ReserveAnalysisReview(artifact, plan,
                context, crossApp, maxResponseTokens);
            return new AnalysisNativeReviewSession
            {
                Context = context,
                Request = request,
                PageImages = images
            };
        }

        public static AnalysisReviewDecision CompleteNativeReview(
            object presentation, AnalysisNativeReviewSession session,
            string reviewerJson)
        {
            RequireEnabled();
            if (session?.Context == null || session.Request == null)
                throw new InvalidOperationException("REVIEW_SESSION_INVALID");
            var pages = session.Context.Pages.OrderBy(page =>
                page.ExpectedPageNumber).ToArray();
            dynamic deck = presentation;
            if (deck == null || (int)deck.Slides.Count != pages.Length ||
                pages.Where((page, index) =>
                    page.ExpectedPageNumber != index + 1).Any())
                throw new InvalidOperationException("REVIEW_NATIVE_STATE_CHANGED");
            for (var index = 1; index <= pages.Length; index++)
            {
                dynamic slide = deck.Slides[index];
                if ((int)slide.SlideID != pages[index - 1].NativeSlideId ||
                    NativeStateFingerprint(slide) !=
                        pages[index - 1].NativeStateFingerprint)
                    throw new InvalidOperationException(
                        "REVIEW_NATIVE_STATE_CHANGED");
            }
            return AnalysisReviewContract.Parse(reviewerJson,
                session.Context);
        }

        public static string CompiledNativeText(AnalysisArtifact artifact,
            AnalysisDocumentPlan plan, string logicalSlideId,
            string targetId)
        {
            RequireEnabled();
            if (targetId != "title" && targetId != "subtitle" &&
                targetId != "takeaway")
                throw new InvalidOperationException(
                    "REPAIR_NATIVE_TEXT_TARGET_UNSUPPORTED");
            var compiled = AnalysisDocumentCompiler.Compile(artifact, plan);
            var index = plan.Slides.FindIndex(slide =>
                slide.Id == logicalSlideId);
            if (index < 0)
                throw new InvalidOperationException("REPAIR_SLIDE_CHANGED");
            return Convert.ToString(compiled.Slides[index][targetId]) ??
                string.Empty;
        }

        public static string ReadNativePatchText(object presentation,
            AnalysisReviewPage page, string expectedText)
        {
            RequireEnabled();
            dynamic deck = presentation;
            if (page == null || page.PageOrdinal != 0 ||
                page.ExpectedPageNumber < 1 ||
                page.ExpectedPageNumber > (int)deck.Slides.Count)
                throw new InvalidOperationException(
                    "REPAIR_NATIVE_PAGE_UNSUPPORTED");
            dynamic slide = deck.Slides[page.ExpectedPageNumber];
            if ((int)slide.SlideID != page.NativeSlideId)
                throw new InvalidOperationException("REPAIR_NATIVE_PAGE_CHANGED");
            dynamic match = null;
            for (var index = 1; index <= (int)slide.Shapes.Count;
                index++)
            {
                dynamic shape = slide.Shapes[index];
                if ((int)shape.HasTextFrame == 0 ||
                    (int)shape.HasChart != 0 ||
                    (int)shape.HasTable != 0 ||
                    Convert.ToString(shape.TextFrame.TextRange.Text) !=
                        expectedText) continue;
                if (match != null)
                    throw new InvalidOperationException(
                        "REPAIR_NATIVE_TEXT_AMBIGUOUS");
                match = shape;
            }
            if (match == null)
                throw new InvalidOperationException(
                    "REPAIR_NATIVE_TEXT_CHANGED");
            return Convert.ToString(match.TextFrame.TextRange.Text) ??
                string.Empty;
        }

        public static void ApplyNativeContentPatch(object presentation,
            AnalysisReviewPage page,
            AnalysisContentPatchReservation reservation)
        {
            RequireEnabled();
            dynamic deck = presentation;
            if (page == null || reservation == null ||
                page.LogicalSlideId != reservation.LogicalSlideId ||
                page.NativeSlideId != reservation.NativeSlideId ||
                page.NativeStateFingerprint !=
                    reservation.NativeStateFingerprint ||
                page.PageOrdinal != 0)
                throw new InvalidOperationException(
                    "REPAIR_RESERVATION_CHANGED");
            dynamic slide = deck.Slides[page.ExpectedPageNumber];
            if ((int)slide.SlideID != page.NativeSlideId ||
                NativeStateFingerprint(slide) !=
                    page.NativeStateFingerprint)
                throw new InvalidOperationException(
                    "REPAIR_NATIVE_PAGE_CHANGED");
            // Uniqueness is checked before mutation; the same exact match is
            // located again to keep this operation limited to one text shape.
            ReadNativePatchText(presentation, page,
                reservation.NativeBeforeText);
            for (var index = 1; index <= (int)slide.Shapes.Count;
                index++)
            {
                dynamic shape = slide.Shapes[index];
                if ((int)shape.HasTextFrame != 0 &&
                    (int)shape.HasChart == 0 &&
                    (int)shape.HasTable == 0 &&
                    Convert.ToString(shape.TextFrame.TextRange.Text) ==
                        reservation.NativeAfterText)
                    throw new InvalidOperationException(
                        "REPAIR_NATIVE_TEXT_AMBIGUOUS");
            }
            dynamic target = null;
            for (var index = 1; index <= (int)slide.Shapes.Count;
                index++)
            {
                dynamic shape = slide.Shapes[index];
                if ((int)shape.HasTextFrame != 0 &&
                    (int)shape.HasChart == 0 &&
                    (int)shape.HasTable == 0 &&
                    Convert.ToString(shape.TextFrame.TextRange.Text) ==
                        reservation.NativeBeforeText)
                    target = shape;
            }
            try
            {
                target.TextFrame.TextRange.Text =
                    reservation.NativeAfterText;
                if (Convert.ToString(target.TextFrame.TextRange.Text) !=
                    reservation.NativeAfterText)
                    throw new InvalidOperationException(
                        "REPAIR_NATIVE_TEXT_READBACK_FAILED");
            }
            catch
            {
                try { target.TextFrame.TextRange.Text =
                    reservation.NativeBeforeText; }
                catch { throw new InvalidOperationException(
                    "REPAIR_NATIVE_RECOVERY_REQUIRED"); }
                throw;
            }
        }

        public static IReadOnlyList<AnalysisReviewPage> CapturePresentationPages(
            object presentation, AnalysisArtifact artifact,
            AnalysisDocumentPlan plan, List<string> pageImages = null)
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
            byte[][] pdfImages = null;
            if (pageImages != null && Enumerable.Range(1, composed.Count)
                .Any(index => PresentationInspection.ContainsNativeChart(
                    (object)deck.Slides[index])))
                pdfImages = RenderPresentationPdf(presentation,
                    composed.Count);
            var pages = new List<AnalysisReviewPage>();
            var ordinals = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var index = 1; index <= composed.Count; index++)
            {
                dynamic native = deck.Slides[index];
                var logicalId = composed[index - 1].Source.Id;
                int ordinal;
                ordinals.TryGetValue(logicalId, out ordinal);
                ordinals[logicalId] = ordinal + 1;
                string pageImage;
                var nativeState = NativeStateFingerprint(native);
                string rendered;
                if (pdfImages != null)
                {
                    var bytes = pdfImages[index - 1];
                    pageImage = "data:image/png;base64," +
                        Convert.ToBase64String(bytes);
                    using (var digest = SHA256.Create())
                        rendered = BitConverter.ToString(digest.ComputeHash(
                            bytes)).Replace("-", "").ToLowerInvariant();
                }
                else rendered = RenderFingerprint(native,
                    nativeState, pageImages != null, out pageImage);
                pageImages?.Add(pageImage);
                pages.Add(new AnalysisReviewPage
                {
                    LogicalSlideId = logicalId,
                    NativeSlideId = (int)native.SlideID,
                    ExpectedPageNumber = index,
                    PageOrdinal = ordinal,
                    RenderFingerprint = rendered,
                    NativeStateFingerprint = nativeState
                });
            }
            return pages;
        }

        private static byte[][] RenderPresentationPdf(object presentation,
            int expectedPages)
        {
            if (expectedPages < 1 || expectedPages > 16)
                throw new InvalidOperationException(
                    "ANALYSIS_VISUAL_REVIEW_UNAVAILABLE: Page count exceeds the bounded native review limit.");
            dynamic deck = presentation;
            if (!string.IsNullOrEmpty((string)deck.Path))
                throw new InvalidOperationException(
                    "ANALYSIS_VISUAL_REVIEW_UNAVAILABLE: Only an unsaved pilot draft can be exported for review.");
            var nameBefore = (string)deck.FullName;
            var savedBefore = (int)deck.Saved;
            var nativeBefore = Enumerable.Range(1, expectedPages)
                .Select(index => NativeStateFingerprint(deck.Slides[index]))
                .ToArray();
            var editableBefore = Enumerable.Range(1, expectedPages)
                .Select(index => NativeStateFingerprint(deck.Slides[index],
                    false)).ToArray();
            var pdf = Path.Combine(Path.GetTempPath(),
                "scribble-analysis-review-" + Guid.NewGuid().ToString("N") +
                ".pdf");
            try
            {
                // PDF uses PowerPoint's page renderer without Slide.Export,
                // which terminates chart.dll on some Office builds. SaveAs
                // format 32 leaves this unsaved native draft in place.
                deck.SaveAs(pdf, 32);
                var nameChanged = (string)deck.FullName != nameBefore;
                var savedChanged = (int)deck.Saved != savedBefore;
                var changedPages = Enumerable.Range(1, expectedPages)
                    .Where(index => NativeStateFingerprint(deck.Slides[index]) !=
                        nativeBefore[index - 1]).ToArray();
                var changedEditablePages = Enumerable.Range(1, expectedPages)
                    .Where(index => NativeStateFingerprint(deck.Slides[index],
                        false) != editableBefore[index - 1]).ToArray();
                var pdfMissing = !File.Exists(pdf);
                var pdfTooLarge = !pdfMissing && new FileInfo(pdf).Length >
                    30 * 1024 * 1024;
                if (nameChanged || savedChanged || changedPages.Length > 0 ||
                    pdfMissing || pdfTooLarge)
                    throw new InvalidOperationException(
                        "PDF export changed or exceeded the native draft boundary: " +
                        "name=" + nameChanged + ", saved=" + savedChanged +
                        ", pages=" + string.Join(",", changedPages) +
                        ", editable_pages=" + string.Join(",",
                            changedEditablePages) +
                        ", missing=" + pdfMissing + ", too_large=" +
                        pdfTooLarge + ".");
                using (var stream = File.OpenRead(pdf))
                {
                    var sizes = PDFtoImage.Conversion.GetPageSizes(stream,
                        leaveOpen: true);
                    if (sizes.Count != expectedPages)
                        throw new InvalidOperationException(
                            "PDF page count differs from the native draft.");
                    stream.Position = 0;
                    var images = new List<byte[]>(expectedPages);
                    var totalBytes = 0;
                    foreach (var bitmap in PDFtoImage.Conversion.ToImages(
                        stream, Enumerable.Range(0, expectedPages),
                        options: new PDFtoImage.RenderOptions(
                            Width: 1600, Height: 900)))
                    {
                        using (bitmap)
                        using (var encoded = bitmap.Encode(
                            SkiaSharp.SKEncodedImageFormat.Png, 100))
                        {
                            var bytes = encoded.ToArray();
                            if (bytes.Length < 1000 || bytes.Length >
                                8 * 1024 * 1024)
                                throw new InvalidOperationException(
                                    "PDF review page is empty or too large.");
                            totalBytes += bytes.Length;
                            if (totalBytes > 24 * 1024 * 1024)
                                throw new InvalidOperationException(
                                    "PDF review images exceed the task limit.");
                            images.Add(bytes);
                        }
                    }
                    if (images.Count != expectedPages)
                        throw new InvalidOperationException(
                            "PDF renderer omitted a native page.");
                    return images.ToArray();
                }
            }
            catch (COMException) { throw; }
            catch (Exception error)
            {
                throw new InvalidOperationException(
                    "ANALYSIS_VISUAL_REVIEW_UNAVAILABLE: The PowerPoint PDF review export failed; the draft remains pending. " +
                    error.Message, error);
            }
            finally { if (File.Exists(pdf)) File.Delete(pdf); }
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
                var content = ContentBounds((object)slide, height);
                // Header/footer shapes are excluded from the bounded central
                // move operation. Reject their intersections explicitly;
                // otherwise moving a chart into the title can evade review.
                RejectUnsupportedChromeCollision((object)slide, content,
                    height, page.NativeSlideId);
                for (var first = 0; first < content.Count; first++)
                    for (var second = first + 1; second < content.Count;
                        second++)
                    {
                        double horizontal, vertical;
                        if (!Collides(content[first], content[second],
                            out horizontal, out vertical)) continue;
                        var target = content[first].Id > content[second].Id
                            ? content[first] : content[second];
                        var other = target == content[first]
                            ? content[second] : content[first];
                        findings.Add(Measure("COLLISION", page,
                            "shape:" + target.Id,
                            "shape:" + target.Id + " intersects shape:" +
                                other.Id + " by " +
                                horizontal.ToString("0.0", CultureInfo.InvariantCulture) +
                                "x" + vertical.ToString("0.0",
                                    CultureInfo.InvariantCulture) + " pt",
                            "at least 6 pt separation", "shape:" + other.Id));
                    }
            }
            return findings;
        }

        // Renderer-owned fixes are finite native operations. They neither
        // ask the model to rewrite the slide nor modify a bound fact.
        public static string RepairNativeMeasurement(object presentation,
            IReadOnlyList<AnalysisReviewPage> pages,
            AnalysisReviewMeasurement measurement, string budgetReceipt,
            AnalysisPatchReservation reservation = null)
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
                NativeStateFingerprint(slide) != page.NativeStateFingerprint)
                throw new InvalidOperationException("RENDERER_REPAIR_PAGE_CHANGED");
            var current = CaptureNativeMeasurements(presentation, pages)
                .SingleOrDefault(item => item.MeasurementId ==
                    measurement.MeasurementId);
            if (current == null || current.Code != measurement.Code ||
                current.TargetId != measurement.TargetId ||
                current.OtherTargetId != measurement.OtherTargetId ||
                current.Observed != measurement.Observed ||
                current.Expected != measurement.Expected)
                throw new InvalidOperationException("RENDERER_REPAIR_MEASUREMENT_CHANGED");
            var budget = AnalysisRepairBudget.Read(budgetReceipt);
            if (reservation != null)
            {
                reservation.Validate(page, measurement);
                if (reservation.BudgetReceipt != budgetReceipt)
                    throw new InvalidOperationException(
                        "REPAIR_RESERVATION_CHANGED");
            }
            var nextReceipt = reservation != null ? budgetReceipt :
                budget.ConsumePatch(page.LogicalSlideId,
                    measurement.TargetId);
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
                return nextReceipt;
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
                return nextReceipt;
            }
            if (measurement.Code == "OUT_OF_BOUNDS")
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
                if (target == null)
                    throw new InvalidOperationException(
                        "RENDERER_REPAIR_TARGET_INVALID");
                var left = (float)target.Left;
                var top = (float)target.Top;
                var width = (float)target.Width;
                var height = (float)target.Height;
                var slideWidth = (float)deck.PageSetup.SlideWidth;
                var slideHeight = (float)deck.PageSetup.SlideHeight;
                if (width > slideWidth || height > slideHeight ||
                    width <= 0 || height <= 0)
                    throw new InvalidOperationException(
                        "RENDERER_BOUNDS_UNSUPPORTED: Shape is larger than the canvas.");
                try
                {
                    target.Left = Math.Max(0f, Math.Min(left, slideWidth - width));
                    target.Top = Math.Max(0f, Math.Min(top, slideHeight - height));
                    if ((float)target.Left < -.5f || (float)target.Top < -.5f ||
                        (float)target.Left + (float)target.Width > slideWidth + .5f ||
                        (float)target.Top + (float)target.Height > slideHeight + .5f)
                        throw new InvalidOperationException(
                            "RENDERER_BOUNDS_READBACK_FAILED");
                }
                catch
                {
                    try { target.Left = left; target.Top = top; }
                    catch { throw new InvalidOperationException(
                        "RENDERER_REPAIR_RECOVERY_REQUIRED"); }
                    throw;
                }
                return nextReceipt;
            }
            if (measurement.Code == "COLLISION")
            {
                var content = ContentBounds((object)slide,
                    (double)deck.PageSetup.SlideHeight);
                var target = content.SingleOrDefault(item =>
                    "shape:" + item.Id == measurement.TargetId);
                var other = content.SingleOrDefault(item =>
                    "shape:" + item.Id == measurement.OtherTargetId);
                if (target == null || other == null)
                    throw new InvalidOperationException(
                        "RENDERER_REPAIR_TARGET_INVALID");
                double overlapX, overlapY;
                if (!Collides(target, other, out overlapX, out overlapY))
                    throw new InvalidOperationException(
                        "RENDERER_REPAIR_MEASUREMENT_CHANGED");
                var slideWidth = (double)deck.PageSetup.SlideWidth;
                var slideHeight = (double)deck.PageSetup.SlideHeight;
                const double gap = 6d;
                var positions = new[]
                {
                    new[] { other.Left - target.Width - gap, target.Top },
                    new[] { other.Right + gap, target.Top },
                    new[] { target.Left, other.Top - target.Height - gap },
                    new[] { target.Left, other.Bottom + gap }
                };
                var viable = positions.Select(position => new
                {
                    Left = position[0], Top = position[1],
                    Distance = Math.Abs(position[0] - target.Left) +
                        Math.Abs(position[1] - target.Top)
                }).Where(position => position.Left >= 0 &&
                    position.Top >= 120 &&
                    position.Left + target.Width <= slideWidth &&
                    position.Top + target.Height <= slideHeight * .9 &&
                    content.Where(item => item.Id != target.Id).All(item =>
                    {
                        double x, y;
                        return !Collides(target.At(position.Left, position.Top),
                            item, out x, out y);
                    })).OrderBy(position => position.Distance).FirstOrDefault();
                if (viable == null)
                    throw new InvalidOperationException(
                        "RENDERER_COLLISION_UNSUPPORTED: No bounded translation clears the overlap.");
                dynamic native = null;
                for (var index = 1; index <= (int)slide.Shapes.Count; index++)
                    if ((int)slide.Shapes[index].Id == target.Id)
                        native = slide.Shapes[index];
                if (native == null)
                    throw new InvalidOperationException(
                        "RENDERER_REPAIR_TARGET_INVALID");
                try
                {
                    native.Left = viable.Left;
                    native.Top = viable.Top;
                    var updated = ContentBounds((object)slide, slideHeight);
                    var moved = updated.SingleOrDefault(item => item.Id == target.Id);
                    if (moved == null || updated.Where(item => item.Id != target.Id)
                        .Any(item =>
                        {
                            double x, y;
                            return Collides(moved, item, out x, out y);
                        }))
                        throw new InvalidOperationException(
                            "RENDERER_COLLISION_READBACK_FAILED");
                }
                catch
                {
                    try { native.Left = target.Left; native.Top = target.Top; }
                    catch { throw new InvalidOperationException(
                        "RENDERER_REPAIR_RECOVERY_REQUIRED"); }
                    throw;
                }
                return nextReceipt;
            }
            throw new InvalidOperationException(
                "RENDERER_REPAIR_UNSUPPORTED: " + measurement.Code);
        }

        private static string RenderFingerprint(dynamic slide,
            string nativeState, bool includeImage, out string dataUrl)
        {
            // PowerPoint chart.dll can terminate the host during Slide.Export.
            // Keep native identity for geometry/repair, but never present it as
            // visual evidence or send a model a partial set of page images.
            if (PresentationInspection.ContainsNativeChart((object)slide))
            {
                dataUrl = null;
                using (var digest = SHA256.Create())
                    return BitConverter.ToString(digest.ComputeHash(
                        Encoding.UTF8.GetBytes("preview-unavailable:" +
                            nativeState))).Replace("-", "").ToLowerInvariant();
            }
            var temporary = Path.Combine(Path.GetTempPath(),
                "scribble-analysis-review-" + Guid.NewGuid().ToString("N") +
                ".png");
            try
            {
                slide.Export(temporary, "PNG", 1600, 900);
                var bytes = File.ReadAllBytes(temporary);
                dataUrl = includeImage ? "data:image/png;base64," +
                    Convert.ToBase64String(bytes) : null;
                using (var digest = SHA256.Create())
                    return BitConverter.ToString(digest.ComputeHash(bytes))
                        .Replace("-", "").ToLowerInvariant();
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        // Native repair identity uses stable editable state. PowerPoint can
        // export different PNG bytes for the same live slide; rendered bytes
        // remain review evidence but cannot safely authorize a COM mutation.
        private static string NativeStateFingerprint(dynamic slide,
            bool includeChartPackage = true)
        {
            var state = new StringBuilder();
            AppendState(state, (int)slide.SlideID);
            AppendState(state, (int)slide.Shapes.Count);
            for (var index = 1; index <= (int)slide.Shapes.Count; index++)
            {
                dynamic shape = slide.Shapes[index];
                AppendState(state, (int)shape.Id);
                AppendState(state, (int)shape.Type);
                AppendState(state, Convert.ToDouble(shape.Left,
                    CultureInfo.InvariantCulture));
                AppendState(state, Convert.ToDouble(shape.Top,
                    CultureInfo.InvariantCulture));
                AppendState(state, Convert.ToDouble(shape.Width,
                    CultureInfo.InvariantCulture));
                AppendState(state, Convert.ToDouble(shape.Height,
                    CultureInfo.InvariantCulture));
                AppendState(state, Convert.ToDouble(shape.Rotation,
                    CultureInfo.InvariantCulture));
                var hasText = (int)shape.HasTextFrame != 0;
                AppendState(state, hasText ? 1 : 0);
                if (hasText)
                {
                    dynamic range = shape.TextFrame.TextRange;
                    AppendState(state, Convert.ToString(range.Text) ?? string.Empty);
                    AppendState(state, Convert.ToDouble(range.Font.Size,
                        CultureInfo.InvariantCulture));
                }
                var hasTable = (int)shape.HasTable != 0;
                AppendState(state, hasTable ? 1 : 0);
                if (hasTable)
                {
                    dynamic table = shape.Table;
                    AppendState(state, (int)table.Rows.Count);
                    AppendState(state, (int)table.Columns.Count);
                    for (var row = 1; row <= (int)table.Rows.Count; row++)
                        for (var column = 1;
                            column <= (int)table.Columns.Count; column++)
                        {
                            dynamic cell = table.Cell(row, column).Shape;
                            dynamic range = cell.TextFrame.TextRange;
                            AppendState(state, Convert.ToString(range.Text) ??
                                string.Empty);
                            AppendState(state, Convert.ToDouble(range.Font.Size,
                                CultureInfo.InvariantCulture));
                        }
                }
                var hasChart = (int)shape.HasChart != 0;
                AppendState(state, hasChart ? 1 : 0);
            }
            if (includeChartPackage &&
                PresentationInspection.ContainsNativeChart((object)slide))
                AppendState(state,
                    PresentationInspection.PackageSlideFingerprint(
                        (object)slide));
            using (var digest = SHA256.Create())
                return BitConverter.ToString(digest.ComputeHash(
                    Encoding.UTF8.GetBytes(state.ToString()))).Replace("-", "")
                    .ToLowerInvariant();
        }

        private static void AppendState(StringBuilder state, object value)
        {
            var text = value is double
                ? ((double)value).ToString("R", CultureInfo.InvariantCulture)
                : Convert.ToString(value, CultureInfo.InvariantCulture) ??
                    string.Empty;
            state.Append(text.Length).Append(':').Append(text);
        }

        private sealed class NativeBounds
        {
            public int Id;
            public double Left, Top, Width, Height;
            public double Right { get { return Left + Width; } }
            public double Bottom { get { return Top + Height; } }
            public NativeBounds At(double left, double top)
            { return new NativeBounds { Id = Id, Left = left, Top = top,
                Width = Width, Height = Height }; }
        }

        private static List<NativeBounds> ContentBounds(object slideObject,
            double slideHeight)
        {
            dynamic slide = slideObject;
            var result = new List<NativeBounds>();
            for (var index = 1; index <= (int)slide.Shapes.Count; index++)
            {
                dynamic shape = slide.Shapes[index];
                var bounds = new NativeBounds
                {
                    Id = (int)shape.Id, Left = (double)shape.Left,
                    Top = (double)shape.Top, Width = (double)shape.Width,
                    Height = (double)shape.Height
                };
                // Header and footer chrome has intentional ink and textbox
                // overlaps. The pilot measures the central content canvas.
                if (bounds.Top < 120 || bounds.Bottom > slideHeight * .9 ||
                    bounds.Width <= 0 || bounds.Height <= 0) continue;
                if ((int)shape.HasTextFrame != 0 &&
                    (int)shape.HasTable == 0 && (int)shape.HasChart == 0 &&
                    string.IsNullOrWhiteSpace(Convert.ToString(
                        shape.TextFrame.TextRange.Text))) continue;
                result.Add(bounds);
            }
            return result;
        }

        private static void RejectUnsupportedChromeCollision(object slideObject,
            IReadOnlyList<NativeBounds> central, double slideHeight,
            int nativeSlideId)
        {
            dynamic slide = slideObject;
            var semantic = new List<NativeBounds>();
            var chromeIds = new HashSet<int>();
            for (var index = 1; index <= (int)slide.Shapes.Count; index++)
            {
                dynamic shape = slide.Shapes[index];
                var table = (int)shape.HasTable != 0;
                var chart = (int)shape.HasChart != 0;
                var text = (int)shape.HasTextFrame != 0 &&
                    !string.IsNullOrWhiteSpace(Convert.ToString(
                        shape.TextFrame.TextRange.Text));
                if (!table && !chart && !text) continue;
                var bounds = new NativeBounds
                {
                    Id = (int)shape.Id, Left = (double)shape.Left,
                    Top = (double)shape.Top, Width = (double)shape.Width,
                    Height = (double)shape.Height
                };
                if (bounds.Width > 0 && bounds.Height > 0)
                {
                    semantic.Add(bounds);
                    // The writer creates the title/action first and the
                    // source/folio/draft marker last. Keep that identity even
                    // if a later edit moves a body shape into their band.
                    if (text && !table && !chart &&
                        ((index <= 2 && bounds.Top < 120) ||
                         (index > (int)slide.Shapes.Count - 3 &&
                          bounds.Top >= slideHeight * .9)))
                        chromeIds.Add(bounds.Id);
                }
            }
            var centralIds = new HashSet<int>(central.Select(item => item.Id));
            for (var first = 0; first < semantic.Count; first++)
                for (var second = first + 1; second < semantic.Count;
                    second++)
                {
                    var left = semantic[first];
                    var right = semantic[second];
                    if (centralIds.Contains(left.Id) &&
                        centralIds.Contains(right.Id)) continue;
                    // Chrome-on-chrome layering is deliberate (for example,
                    // the folio can sit over the draft footer marker).
                    if (chromeIds.Contains(left.Id) &&
                        chromeIds.Contains(right.Id)) continue;
                    var overlapX = Math.Min(left.Right, right.Right) -
                        Math.Max(left.Left, right.Left);
                    var overlapY = Math.Min(left.Bottom, right.Bottom) -
                        Math.Max(left.Top, right.Top);
                    if (overlapX <= 4 || overlapY <= 4 ||
                        overlapX * overlapY /
                        Math.Min(left.Width * left.Height,
                            right.Width * right.Height) < .10) continue;
                    throw new InvalidOperationException(
                        "ANALYSIS_PILOT_GEOMETRY_UNSUPPORTED: Slide " +
                        nativeSlideId + " has a collision outside the bounded " +
                        "content canvas between shape:" + left.Id +
                        " and shape:" + right.Id + ".");
                }
        }

        private static bool Collides(NativeBounds first, NativeBounds second,
            out double horizontal, out double vertical)
        {
            horizontal = Math.Min(first.Right, second.Right) -
                Math.Max(first.Left, second.Left);
            vertical = Math.Min(first.Bottom, second.Bottom) -
                Math.Max(first.Top, second.Top);
            if (horizontal <= 4 || vertical <= 4) return false;
            if (Contains(first, second) || Contains(second, first))
                return false;
            return horizontal * vertical /
                Math.Min(first.Width * first.Height,
                    second.Width * second.Height) >= .10;
        }

        private static bool Contains(NativeBounds outer, NativeBounds inner)
        {
            return inner.Left >= outer.Left - 2 &&
                inner.Top >= outer.Top - 2 &&
                inner.Right <= outer.Right + 2 &&
                inner.Bottom <= outer.Bottom + 2;
        }

        private static AnalysisReviewMeasurement Measure(string code,
            AnalysisReviewPage page, string target, string observed,
            string expected, string otherTarget = null)
        {
            return new AnalysisReviewMeasurement
            {
                MeasurementId = "native:" + page.NativeSlideId + ":" +
                    code + ":" + target,
                Code = code, LogicalSlideId = page.LogicalSlideId,
                NativeSlideId = page.NativeSlideId, TargetId = target,
                OtherTargetId = otherTarget,
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
