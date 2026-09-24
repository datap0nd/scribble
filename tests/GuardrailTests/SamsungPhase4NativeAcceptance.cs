using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using Scribble.Office;

namespace GuardrailTests
{
    // Explicit workstation run. Every deck is synthetic and disposable. PDF
    // pages and editable PPTX are kept for an identified visual reviewer.
    internal static class SamsungPhase4NativeAcceptance
    {
        internal static int Run(string reportPath)
        {
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var output = Path.GetDirectoryName(Path.GetFullPath(reportPath));
            Directory.CreateDirectory(output);
            var fixturePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Fixtures", "phase4-reference.json");
            var fixtures = ((IEnumerable)json.DeserializeObject(
                File.ReadAllText(fixturePath)))
                .Cast<Dictionary<string, object>>().ToArray();
            dynamic app = null;
            dynamic deck = null;
            var pages = new List<object>();
            var files = new List<string>();
            var failure = string.Empty;
            var passed = false;
            try
            {
                SamsungSlideTests.Phase4ReferenceMatrix();
                app = Activator.CreateInstance(Type.GetTypeFromProgID(
                    "PowerPoint.Application", true));
                app.Visible = -1;
                for (var batch = 0; batch < 3; batch++)
                {
                    var subset = fixtures.Skip(batch * 6).Take(6).ToArray();
                    var input = subset.Select(fixture =>
                    {
                        var slide = (Dictionary<string, object>)fixture["slide"];
                        slide["id"] = fixture["id"];
                        return (object)slide;
                    }).ToArray();
                    var slides = PresentationDraftWriter.ParseSlides(input);
                    PresentationDraftWriter.AddDraftSlides((object)app,
                        slides, null, true, null, null, null, null, true);
                    deck = app.ActivePresentation;
                    if ((int)deck.Slides.Count != subset.Length)
                        throw new InvalidOperationException(
                            "PHASE4_NATIVE_PAGE_COUNT_CHANGED");
                    for (var index = 1; index <= subset.Length; index++)
                    {
                        dynamic native = deck.Slides[index];
                        var slide = (Dictionary<string, object>)subset[index - 1]["slide"];
                        var hasChart = false;
                        var hasTable = false;
                        for (var shapeIndex = 1; shapeIndex <=
                            (int)native.Shapes.Count; shapeIndex++)
                        {
                            dynamic shape = native.Shapes[shapeIndex];
                            hasChart |= (int)shape.HasChart != 0;
                            hasTable |= (int)shape.HasTable != 0;
                        }
                        if (hasChart != slide.ContainsKey("chart") ||
                            hasTable != slide.ContainsKey("table"))
                            throw new InvalidOperationException(
                                "PHASE4_NATIVE_OBJECT_MISSING: " +
                                subset[index - 1]["id"]);
                        pages.Add(new
                        {
                            fixture_id = subset[index - 1]["id"],
                            family = subset[index - 1]["family"],
                            density = subset[index - 1]["density"],
                            deck_number = batch + 1,
                            page_number = index,
                            native_slide_id = (int)native.SlideID,
                            shape_count = (int)native.Shapes.Count,
                            native_chart = hasChart,
                            native_table = hasTable
                        });
                    }
                    var prefix = "phase4-reference-" + (batch + 1);
                    var pptx = Path.Combine(output, prefix + ".pptx");
                    var pdf = Path.Combine(output, prefix + ".pdf");
                    deck.SaveAs(pptx, 24);
                    deck.SaveAs(pdf, 32);
                    if (!File.Exists(pptx) ||
                        new FileInfo(pptx).Length < 1000 ||
                        !File.Exists(pdf) ||
                        new FileInfo(pdf).Length < 1000)
                        throw new InvalidOperationException(
                            "PHASE4_NATIVE_EXPORT_MISSING");
                    files.Add(pptx);
                    files.Add(pdf);
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
                execution_kind = "native_disposable_phase4_references",
                reference_count = pages.Count,
                structural_passed = passed,
                visual_approved = false,
                reviewer = (string)null,
                pages,
                files,
                failure,
                note = "The reference renders require identified human approval. No reviewer or model calibration is inferred from these exports."
            };
            File.WriteAllText(reportPath, json.Serialize(report));
            Console.WriteLine(json.Serialize(report));
            return passed ? 0 : 1;
        }
    }
}
