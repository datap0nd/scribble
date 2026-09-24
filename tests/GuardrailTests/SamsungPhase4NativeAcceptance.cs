using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Web.Script.Serialization;
using Scribble.Office;

namespace GuardrailTests
{
    // Explicit workstation run. Every deck is synthetic and disposable. PDF
    // pages and editable PPTX are kept for an identified visual reviewer.
    internal static class SamsungPhase4NativeAcceptance
    {
        internal static int Run(string reportPath)
        { return RunInternal(reportPath, false); }

        internal static int RunDefects(string reportPath)
        { return RunInternal(reportPath, true); }

        private static int RunInternal(string reportPath, bool defectMode)
        {
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var output = Path.GetDirectoryName(Path.GetFullPath(reportPath));
            Directory.CreateDirectory(output);
            var fixturePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Fixtures", "phase4-reference.json");
            var fixtures = ((IEnumerable)json.DeserializeObject(
                File.ReadAllText(fixturePath)))
                .Cast<Dictionary<string, object>>().ToArray();
            var assemblyHash = FileHash(typeof(SamsungAuthoringPolicy)
                .Assembly.Location);
            var fixtureHash = FileHash(fixturePath);
            var defects = defectMode ? ((IEnumerable)json.DeserializeObject(
                File.ReadAllText(Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, "Fixtures",
                    "phase4-defects.json"))))
                .Cast<Dictionary<string, object>>().ToArray() : null;
            dynamic app = null;
            dynamic deck = null;
            var pages = new List<object>();
            var files = new List<string>();
            var failure = string.Empty;
            var passed = false;
            try
            {
                SamsungSlideTests.Phase4ReferenceMatrix();
                if (defectMode) SamsungSlideTests.Phase4DefectMatrix();
                if (defectMode)
                {
                    var baselinePath = Path.Combine(output,
                        "phase4-reference-report.json");
                    var baseline = json.Deserialize<Dictionary<string, object>>(
                        File.ReadAllText(baselinePath));
                    if (!Convert.ToBoolean(baseline["structural_passed"]) ||
                        Convert.ToInt32(baseline["reference_count"]) != 18 ||
                        Convert.ToString(baseline["assembly_sha256"]) !=
                            assemblyHash ||
                        Convert.ToString(baseline["fixture_sha256"]) !=
                            fixtureHash)
                        throw new InvalidOperationException(
                            "PHASE4_DEFECT_BASELINE_CHANGED");
                }
                app = Activator.CreateInstance(Type.GetTypeFromProgID(
                    "PowerPoint.Application", true));
                app.Visible = -1;
                var writer = typeof(SamsungAuthoringPolicy).Assembly.GetType(
                    "Scribble.Office.PresentationDraftWriter", true);
                var parse = writer.GetMethods(BindingFlags.Static |
                    BindingFlags.NonPublic).Single(method =>
                    method.Name == "ParseSlides" &&
                    method.GetParameters().Length == 1);
                var add = writer.GetMethods(BindingFlags.Static |
                    BindingFlags.NonPublic).Single(method =>
                    method.Name == "AddDraftSlides" &&
                    method.GetParameters().Length == 9);
                for (var batch = 0; batch < 3; batch++)
                {
                    var subset = fixtures.Skip(batch * 6).Take(6).ToArray();
                    var referencePath = Path.Combine(output,
                        "phase4-reference-" + (batch + 1) + ".pptx");
                    var referenceHash = defectMode ?
                        FileHash(referencePath) : null;
                    if (defectMode)
                        deck = app.Presentations.Open(referencePath, -1,
                            -1, 0);
                    else
                    {
                        var input = subset.Select(fixture =>
                        {
                            var slide = (Dictionary<string, object>)fixture["slide"];
                            slide["id"] = fixture["id"];
                            return (object)slide;
                        }).ToArray();
                        var slides = parse.Invoke(null, new object[] { input });
                        add.Invoke(null, new object[] { (object)app,
                            slides, null, true, null, null, null, null, true });
                        deck = app.ActivePresentation;
                    }
                    if ((int)deck.Slides.Count != subset.Length)
                        throw new InvalidOperationException(
                            "PHASE4_NATIVE_PAGE_COUNT_CHANGED");
                    for (var index = 1; index <= subset.Length; index++)
                    {
                        dynamic native = deck.Slides[index];
                        var slide = (Dictionary<string, object>)subset[index - 1]["slide"];
                        var hasChart = false;
                        var hasTable = false;
                        var geometry = new List<object>();
                        for (var shapeIndex = 1; shapeIndex <=
                            (int)native.Shapes.Count; shapeIndex++)
                        {
                            dynamic shape = native.Shapes[shapeIndex];
                            var shapeChart = (int)shape.HasChart != 0;
                            var shapeTable = (int)shape.HasTable != 0;
                            hasChart |= shapeChart;
                            hasTable |= shapeTable;
                        }
                        if (hasChart != slide.ContainsKey("chart") ||
                            hasTable != slide.ContainsKey("table"))
                            throw new InvalidOperationException(
                                "PHASE4_NATIVE_OBJECT_MISSING: " +
                                subset[index - 1]["id"]);
                        var defect = defectMode ?
                            defects[batch * 6 + index - 1] : null;
                        if (defect != null)
                            ApplyDefect(native, defect, slide);
                        for (var shapeIndex = 1; shapeIndex <=
                            (int)native.Shapes.Count; shapeIndex++)
                        {
                            dynamic shape = native.Shapes[shapeIndex];
                            var shapeChart = (int)shape.HasChart != 0;
                            var shapeTable = (int)shape.HasTable != 0;
                            var shapeText = !shapeChart && !shapeTable &&
                                (int)shape.HasTextFrame != 0;
                            geometry.Add(new
                            {
                                id = (int)shape.Id,
                                type = (int)shape.Type,
                                left = (float)shape.Left,
                                top = (float)shape.Top,
                                width = (float)shape.Width,
                                height = (float)shape.Height,
                                visible = (int)shape.Visible != 0,
                                chart = shapeChart,
                                table = shapeTable,
                                text = shapeText ? Convert.ToString(
                                    shape.TextFrame.TextRange.Text) : null,
                                font_size = shapeText ?
                                    (float?)shape.TextFrame.TextRange.Font.Size :
                                    null
                            });
                        }
                        pages.Add(new
                        {
                            fixture_id = subset[index - 1]["id"],
                            defect_id = defect == null ? null : defect["id"],
                            expected_defect = defect == null ? null :
                                defect["expected"],
                            severity = defect == null ? null :
                                defect["severity"],
                            reference_sha256 = referenceHash,
                            family = subset[index - 1]["family"],
                            density = subset[index - 1]["density"],
                            deck_number = batch + 1,
                            page_number = index,
                            native_slide_id = (int)native.SlideID,
                            shape_count = (int)native.Shapes.Count,
                            native_chart = hasChart,
                            native_table = hasTable,
                            geometry
                        });
                    }
                    var prefix = (defectMode ? "phase4-defect-" :
                        "phase4-reference-") + (batch + 1);
                    var pptx = Path.Combine(output, prefix + ".pptx");
                    var pdf = Path.Combine(output, prefix + ".pdf");
                    var afterExport = Path.Combine(output,
                        prefix + "-after-export.pptx");
                    deck.SaveCopyAs(pptx);
                    deck.SaveAs(pdf, 32);
                    deck.SaveCopyAs(afterExport);
                    var packageCheck = typeof(PresentationInspection)
                        .GetMethod("PdfExportPackageEquivalent",
                            BindingFlags.Static | BindingFlags.NonPublic);
                    if (packageCheck == null)
                        throw new InvalidOperationException(
                            "PHASE4_PACKAGE_BOUNDARY_MISSING");
                    var packageArguments = new object[] { pptx, afterExport,
                        null };
                    if (!(bool)packageCheck.Invoke(null, packageArguments))
                        throw new InvalidOperationException(
                            "PHASE4_PDF_EXPORT_CHANGED_NATIVE_PACKAGE: " +
                            Convert.ToString(packageArguments[2]));
                    if (!File.Exists(pptx) ||
                        new FileInfo(pptx).Length < 1000 ||
                        !File.Exists(pdf) ||
                        new FileInfo(pdf).Length < 1000 ||
                        !File.Exists(afterExport) ||
                        new FileInfo(afterExport).Length < 1000)
                        throw new InvalidOperationException(
                            "PHASE4_NATIVE_EXPORT_MISSING");
                    files.Add(pptx);
                    files.Add(pdf);
                    files.Add(afterExport);
                    if (defectMode && FileHash(referencePath) !=
                        referenceHash)
                        throw new InvalidOperationException(
                            "PHASE4_DEFECT_CHANGED_REFERENCE_DECK");
                    deck.Close();
                    deck = null;
                }
                passed = pages.Count == 18;
            }
            catch (Exception error) { failure = error.ToString(); }
            finally
            {
                if (deck != null) try { deck.Close(); } catch { }
                if (app != null) try
                {
                    if ((int)app.Presentations.Count == 0) app.Quit();
                }
                catch { }
            }
            var report = new
            {
                execution_kind = defectMode ?
                    "native_disposable_phase4_defects" :
                    "native_disposable_phase4_references",
                assembly_sha256 = assemblyHash,
                fixture_sha256 = fixtureHash,
                defect_fixture_sha256 = defectMode ? FileHash(Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, "Fixtures",
                    "phase4-defects.json")) : null,
                reference_count = pages.Count,
                structural_passed = !defectMode && passed,
                seeded_defects_rendered = defectMode && passed,
                visual_approved = false,
                reviewer = (string)null,
                pages,
                files,
                failure,
                note = defectMode ?
                    "Each page contains one seeded blocker for reviewer calibration; rendering alone is not a reviewer verdict." :
                    "The reference renders require identified human approval. No reviewer or model calibration is inferred from these exports."
            };
            File.WriteAllText(reportPath, json.Serialize(report));
            Console.WriteLine(json.Serialize(report));
            return passed ? 0 : 1;
        }

        private static string FileHash(string path)
        {
            if (!File.Exists(path) || new FileInfo(path).Length >
                50 * 1024 * 1024)
                throw new InvalidOperationException(
                    "PHASE4_REFERENCE_DECK_REQUIRED: " + path);
            using (var stream = File.OpenRead(path))
            using (var hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(stream))
                    .Replace("-", "").ToLowerInvariant();
        }

        private static void ApplyDefect(dynamic slide,
            Dictionary<string, object> defect,
            Dictionary<string, object> specification)
        {
            var mutation = Convert.ToString(defect["mutation"]);
            var shapes = Enumerable.Range(1, (int)slide.Shapes.Count)
                .Select(index => (object)slide.Shapes[index]).ToArray();
            Func<object, string> shapeText = value =>
            {
                dynamic shape = value;
                return (int)shape.HasTextFrame != 0 ?
                    Convert.ToString(shape.TextFrame.TextRange.Text) ??
                    string.Empty : string.Empty;
            };
            Func<object> chart = () => shapes.Single(value =>
                (int)((dynamic)value).HasChart != 0);
            Func<object> table = () => shapes.Single(value =>
                (int)((dynamic)value).HasTable != 0);
            Func<string, object> text = wanted => shapes.Single(value =>
                shapeText(value) == wanted);
            dynamic target;
            switch (mutation)
            {
                case "hide_primary_metric":
                    target = text("82,992"); target.Visible = 0; break;
                case "shrink_primary_metric":
                    target = text("82,992");
                    target.TextFrame.TextRange.Font.Size = 7f; break;
                case "crop_title":
                    target = text(Convert.ToString(specification["title"]));
                    target.Width = 90f; target.Height = 45f; break;
                case "overlap_chart_table":
                    target = chart(); dynamic leftTable = table();
                    target.Left = (float)leftTable.Left +
                        (float)leftTable.Width - 100f; break;
                case "move_table_off_canvas":
                    target = table(); target.Left = -110f; break;
                case "hide_chart":
                    target = chart(); target.Visible = 0; break;
                case "shrink_chart":
                    target = chart(); target.Width = 95f; break;
                case "move_chart_into_title":
                    target = chart(); target.Top = 15f; break;
                case "move_chart_off_canvas":
                    target = chart(); target.Left =
                        (float)slide.Parent.PageSetup.SlideWidth - 80f; break;
                case "hide_first_card":
                    target = text("144 source rows");
                    target.Visible = 0; break;
                case "shrink_card_copy":
                    target = shapes.First(value => shapeText(value)
                        .Contains("The ledger includes all records"));
                    target.TextFrame.TextRange.Font.Size = 6f; break;
                case "overlap_cards":
                    target = text("Integrity");
                    dynamic firstCard = text("Coverage");
                    target.Left = firstCard.Left;
                    target.Top = firstCard.Top; break;
                case "shrink_table_width":
                    target = table(); target.Width = 175f; break;
                case "shrink_table_height":
                    target = table(); target.Height = 95f; break;
                case "hide_title_color":
                    target = text(Convert.ToString(specification["title"]));
                    target.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF;
                    break;
                case "hide_subtitle_color":
                    target = text(Convert.ToString(specification["subtitle"]));
                    target.TextFrame.TextRange.Font.Color.RGB =
                        (int)slide.Background.Fill.ForeColor.RGB;
                    break;
                default: throw new InvalidOperationException(
                    "PHASE4_DEFECT_UNKNOWN: " + mutation);
            }
        }
    }
}
