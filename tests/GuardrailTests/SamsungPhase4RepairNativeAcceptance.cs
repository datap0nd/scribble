using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Web.Script.Serialization;
using Scribble.Office;

namespace GuardrailTests
{
    // Explicit workstation run against a generated, disposable PP01 source.
    // The original deck and workbook are opened read-only or not opened at all.
    internal static class SamsungPhase4RepairNativeAcceptance
    {
        private static readonly Type CopyType =
            typeof(SamsungAuthoringPolicy).Assembly.GetType(
                "Scribble.Office.PresentationDraftCopy", true);
        private static readonly Type RevisionType =
            typeof(SamsungAuthoringPolicy).Assembly.GetType(
                "Scribble.Office.PresentationRevision", true);

        internal static int Run(string sourcePath, string workbookPath,
            string reportPath)
        {
            var output = Path.GetDirectoryName(
                Path.GetFullPath(reportPath));
            Directory.CreateDirectory(output);
            var sourceHash = Hash(sourcePath);
            var workbookHash = Hash(workbookPath);
            dynamic app = null;
            dynamic source = null;
            dynamic draft = null;
            object revision = null;
            var failure = string.Empty;
            var passed = false;
            var candidate = Path.Combine(output,
                "phase4-pp01-copy.pptx");
            var pdf = Path.Combine(output,
                "phase4-pp01-copy.pdf");
            var afterExport = Path.Combine(output,
                "phase4-pp01-copy-after-export.pptx");
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID(
                    "PowerPoint.Application", true));
                app.Visible = -1;
                source = app.Presentations.Open(sourcePath, -1, -1, 0);
                if ((int)source.Slides.Count != 6)
                    throw new InvalidOperationException(
                        "PP01_SOURCE_PAGE_COUNT_INVALID");
                var chart = OnlyShape((object)source.Slides[2],
                    shape => (int)shape.HasChart != 0);
                var table = OnlyShape((object)source.Slides[3],
                    shape => (int)shape.HasTable != 0);
                var commentary = OnlyShape((object)source.Slides[4],
                    shape => (int)shape.HasTextFrame != 0 &&
                        Convert.ToString(shape.TextFrame.TextRange.Text)
                            .Contains("The monthly comparison covers"));
                if ((float)chart.Left + (float)chart.Width <=
                        (float)source.PageSetup.SlideWidth ||
                    (int)table.Table.Cell(1, 1).Shape.Fill.ForeColor.RGB ==
                        MetoTheme.Rgb(SamsungSlideDesign.Blue) ||
                    (float)commentary.TextFrame.TextRange.BoundHeight <=
                        (float)commentary.Height + 1)
                    throw new InvalidOperationException(
                        "PP01_SOURCE_DEFECTS_NOT_PRESENT");
                var copy = InvokeStatic(CopyType, "Create",
                    (object)app, (object)source,
                    "phase4-pp01-native");
                draft = CopyType.GetField("Draft",
                    BindingFlags.Instance | BindingFlags.NonPublic)
                    .GetValue(copy);
                var untouched = new[] { 1, 5, 6 }.ToDictionary(index =>
                    index, index => CopyContent((object)
                        draft.Slides[index]));
                var operations = new List<object>();
                operations.Add(Geometry((object)source.Slides[2],
                    chart, 66f, 158.25f, 825f, 278.25f));
                for (var column = 1; column <= 3; column++)
                    operations.Add(new Dictionary<string, object>
                    {
                        { "kind", "table_cell_fill" },
                        { "slide_id", (int)source.Slides[3].SlideID },
                        { "fingerprint", PresentationInspection
                            .Fingerprint((object)source.Slides[3]) },
                        { "shape_id", (int)table.Id },
                        { "row", 1 }, { "column", column },
                        { "before_color", (int)table.Table.Cell(1,
                            column).Shape.Fill.ForeColor.RGB },
                        { "color", MetoTheme.Rgb(
                            SamsungSlideDesign.Blue) }
                    });
                operations.Add(Geometry((object)source.Slides[4],
                    commentary, 63f, 153f, 832.5f, 292.5f));
                var bound = (object[])Invoke(copy, CopyType,
                    "BindOperations", (object)operations.ToArray());
                revision = Activator.CreateInstance(RevisionType,
                    BindingFlags.Instance | BindingFlags.NonPublic,
                    null, new[] { (object)draft }, null);
                Invoke(revision, RevisionType, "Stage", (object)app,
                    bound);
                Invoke(copy, CopyType, "VerifySource");
                Invoke(revision, RevisionType, "Commit",
                    (Action<string>)(status => { }));
                Invoke(copy, CopyType, "VerifySource");
                foreach (var page in untouched)
                    if (CopyContent((object)draft.Slides[page.Key]) !=
                            page.Value)
                        throw new InvalidOperationException(
                            "PP01_UNRELATED_DRAFT_PAGE_CHANGED");
                dynamic repairedChart = OnlyShape(
                    (object)draft.Slides[2],
                    shape => (int)shape.HasChart != 0);
                dynamic repairedTable = OnlyShape(
                    (object)draft.Slides[3],
                    shape => (int)shape.HasTable != 0);
                dynamic repairedCommentary = OnlyShape(
                    (object)draft.Slides[4],
                    shape => (int)shape.HasTextFrame != 0 &&
                        Convert.ToString(shape.TextFrame.TextRange.Text)
                            .Contains("The monthly comparison covers"));
                if ((int)draft.Slides.Count != 6 ||
                    (float)repairedChart.Left < 0 ||
                    (float)repairedChart.Left +
                        (float)repairedChart.Width >
                            (float)draft.PageSetup.SlideWidth ||
                    (float)repairedCommentary.TextFrame.TextRange
                        .BoundHeight >
                            (float)repairedCommentary.Height + 1 ||
                    Enumerable.Range(1, 3).Any(column =>
                        (int)repairedTable.Table.Cell(1, column)
                            .Shape.Fill.ForeColor.RGB != MetoTheme.Rgb(
                                SamsungSlideDesign.Blue)))
                    throw new InvalidOperationException(
                        "PP01_DRAFT_REPAIR_READBACK_FAILED");
                for (var index = 1; index <= 6; index++)
                    if (PresentationInspection.Notes(
                            (object)source.Slides[index]) !=
                        PresentationInspection.Notes(
                            (object)draft.Slides[index]))
                        throw new InvalidOperationException(
                            "PP01_DRAFT_NOTES_CHANGED");
                draft.SaveCopyAs(candidate);
                draft.SaveAs(pdf, 32);
                draft.SaveCopyAs(afterExport);
                var packageCheck = typeof(PresentationInspection)
                    .GetMethod("PdfExportPackageEquivalent",
                        BindingFlags.Static | BindingFlags.NonPublic);
                if (packageCheck == null)
                    throw new InvalidOperationException(
                        "PP01_PACKAGE_BOUNDARY_MISSING");
                var packageArguments = new object[] { candidate,
                    afterExport, null };
                if (!(bool)packageCheck.Invoke(null, packageArguments))
                    throw new InvalidOperationException(
                        "PP01_PDF_EXPORT_CHANGED_NATIVE_PACKAGE: " +
                        Convert.ToString(packageArguments[2]));
                Invoke(copy, CopyType, "VerifySource");
                passed = true;
            }
            catch (Exception error) { failure = error.ToString(); }
            finally
            {
                if (revision != null) try
                {
                    Invoke(revision, RevisionType, "CloseStaging",
                        false);
                }
                catch { }
                if (draft != null) try { draft.Close(); } catch { }
                if (source != null) try { source.Close(); } catch { }
                if (app != null) try
                {
                    if ((int)app.Presentations.Count == 0) app.Quit();
                }
                catch { }
            }
            var sourceUnchanged = Hash(sourcePath) == sourceHash;
            var workbookUnchanged = Hash(workbookPath) == workbookHash;
            var report = new
            {
                execution_kind = "native_disposable_phase4_pp01_copy",
                assembly_sha256 = Hash(typeof(SamsungAuthoringPolicy)
                    .Assembly.Location),
                source_sha256 = sourceHash,
                workbook_sha256 = workbookHash,
                source_unchanged = sourceUnchanged,
                workbook_unchanged = workbookUnchanged,
                patch_stage_passed = passed,
                native_artifact = passed ? candidate : null,
                pdf_review_artifact = passed ? pdf : null,
                chart_recreated_from_workbook = false,
                independent_grader_passed = false,
                visual_approved = false,
                full_acceptance_passed = false,
                failure,
                note = "A new six-slide copy received bounded geometry and table-fill patches. Chart creation from the workbook, independent grading, and human visual approval remain separate gates."
            };
            File.WriteAllText(reportPath,
                new JavaScriptSerializer().Serialize(report));
            Console.WriteLine(new JavaScriptSerializer()
                .Serialize(report));
            return passed && sourceUnchanged &&
                workbookUnchanged ? 0 : 1;
        }

        private static Dictionary<string, object> Geometry(object slide,
            dynamic shape, float left, float top, float width,
            float height)
        {
            dynamic page = slide;
            return new Dictionary<string, object>
            {
                { "kind", "shape_geometry" },
                { "slide_id", (int)page.SlideID },
                { "fingerprint", PresentationInspection
                    .Fingerprint(slide) },
                { "shape_id", (int)shape.Id },
                { "before_left", (float)shape.Left },
                { "before_top", (float)shape.Top },
                { "before_width", (float)shape.Width },
                { "before_height", (float)shape.Height },
                { "left", left }, { "top", top },
                { "width", width }, { "height", height }
            };
        }

        private static dynamic OnlyShape(object slide,
            Func<dynamic, bool> predicate)
        {
            dynamic page = slide;
            var matches = new List<object>();
            for (var index = 1; index <= (int)page.Shapes.Count;
                 index++)
            {
                dynamic shape = page.Shapes[index];
                if (predicate(shape)) matches.Add((object)shape);
            }
            if (matches.Count != 1)
                throw new InvalidOperationException(
                    "PP01_EXPECTED_SHAPE_NOT_UNIQUE");
            return matches[0];
        }

        private static object InvokeStatic(Type type, string method,
            params object[] arguments)
        {
            try { return type.GetMethod(method, BindingFlags.Static |
                BindingFlags.NonPublic).Invoke(null, arguments); }
            catch (TargetInvocationException error)
            { System.Runtime.ExceptionServices.ExceptionDispatchInfo
                .Capture(error.InnerException).Throw(); throw; }
        }

        private static string CopyContent(object slide)
        {
            var method = typeof(PresentationInspection).GetMethod(
                "CopyContentFingerprint", BindingFlags.Static |
                BindingFlags.NonPublic);
            if (method == null)
                throw new InvalidOperationException(
                    "PP01_COPY_PRESERVATION_MISSING");
            return (string)method.Invoke(null,
                new[] { slide });
        }

        private static object Invoke(object target, Type type,
            string method, params object[] arguments)
        {
            try { return type.GetMethod(method,
                BindingFlags.Instance | BindingFlags.NonPublic)
                    .Invoke(target, arguments); }
            catch (TargetInvocationException error)
            { System.Runtime.ExceptionServices.ExceptionDispatchInfo
                .Capture(error.InnerException).Throw(); throw; }
        }

        private static string Hash(string path)
        {
            if (!File.Exists(path) || new FileInfo(path).Length >
                50 * 1024 * 1024)
                throw new InvalidOperationException(
                    "PP01_NATIVE_INPUT_MISSING: " + path);
            using (var stream = File.OpenRead(path))
            using (var hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(stream))
                    .Replace("-", "").ToLowerInvariant();
        }
    }
}
