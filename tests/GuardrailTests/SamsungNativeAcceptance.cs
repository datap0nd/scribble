using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Web.Script.Serialization;
using Scribble.Office;

namespace GuardrailTests
{
    // Explicit opt-in. Creates only unsaved synthetic test decks, never opens or
    // changes a user's existing presentation. Does not claim model/UI acceptance.
    internal static class SamsungNativeAcceptance
    {
        private static readonly Type Transaction = typeof(SamsungAuthoringPolicy).Assembly.GetType("Scribble.Office.PresentationRevision", true);
        private static object Invoke(object instance, string name, params object[] args)
        { try { return Transaction.GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(instance, args); } catch (TargetInvocationException ex) { throw ex.InnerException; } }
        private static void Check(bool condition, string message) { if (!condition) throw new Exception(message); }
        private static object NewTransaction(object deck)
        { return Activator.CreateInstance(Transaction, BindingFlags.Instance | BindingFlags.NonPublic, null, new[] { deck }, null); }
        private static void RetainTemps(object transaction, List<object> temps)
        {
            foreach (var name in new[] { "Working", "Recovery" })
                temps.Add(Transaction.GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(transaction));
        }
        private static string Content(object slide)
        { return (string)typeof(PresentationInspection).GetMethod("ContentFingerprint", BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, new[] { slide }); }
        internal static int Run(string reportPath)
        {
            dynamic app = null; dynamic deck = null;
            var temps = new List<object>(); var passed = false; var failure = "";
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application", true));
                deck = app.Presentations.Add(0);
                deck.PageSetup.SlideWidth = 960; deck.PageSetup.SlideHeight = 540;
                dynamic slide = deck.Slides.Add(1, 12);
                dynamic title = slide.Shapes.AddTextbox(1, 60, 35, 700, 80);
                title.TextFrame.TextRange.Text = "Sales Q2"; title.TextFrame.TextRange.Font.Size = 24;
                title.TextFrame.TextRange.Characters(1, 5).Font.Bold = -1;
                slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = "User-authored notes must remain.";
                dynamic untouched = deck.Slides.Add(2, 12);
                untouched.Shapes.AddTextbox(1, 60, 60, 700, 80).TextFrame.TextRange.Text = "Untouched source slide";
                var before = PresentationInspection.Fingerprint((object)slide);
                var unchanged = PresentationInspection.Fingerprint((object)untouched);
                var operation = new Dictionary<string, object> { { "kind", "replace_text" }, { "slide_id", (int)slide.SlideID }, { "shape_id", (int)title.Id },
                    { "fingerprint", before }, { "before", "Sales Q2" }, { "text", "Sales Q3" } };
                var transaction = NewTransaction((object)deck);
                Invoke(transaction, "Stage", (object)app, new object[] { operation });
                temps.Add(Transaction.GetField("Working", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(transaction));
                temps.Add(Transaction.GetField("Recovery", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(transaction));
                Invoke(transaction, "Commit", (Action<string>)(status => { }));
                Check(Convert.ToString(title.TextFrame.TextRange.Text) == "Sales Q3", "Revision did not apply.");
                Check(PresentationInspection.Fingerprint((object)untouched) == unchanged, "Unrelated slide changed.");
                Check(PresentationInspection.Notes((object)slide).Contains("User-authored notes must remain."), "Notes were lost.");
                Invoke(transaction, "Revert");
                Check(PresentationInspection.Fingerprint((object)slide) == before, "Revert failed to restore native content and formatting.");
                var rollback = NewTransaction((object)deck);
                Invoke(rollback, "Stage", (object)app, new object[] { operation });
                temps.Add(Transaction.GetField("Working", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(rollback));
                temps.Add(Transaction.GetField("Recovery", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(rollback));
                try { Invoke(rollback, "Commit", (Action<string>)(status => { if (status.StartsWith("applied:")) throw new InvalidOperationException("Injected journal failure"); })); }
                catch (InvalidOperationException) { }
                Check(PresentationInspection.Fingerprint((object)slide) == before, "Failed batch did not roll back.");
                var concurrent = NewTransaction((object)deck);
                Invoke(concurrent, "Stage", (object)app, new object[] { operation });
                temps.Add(Transaction.GetField("Working", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(concurrent));
                temps.Add(Transaction.GetField("Recovery", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(concurrent));
                title.TextFrame.TextRange.Text = "A concurrent user edit";
                var blocked = false;
                try { Invoke(concurrent, "Commit", (Action<string>)(status => { })); } catch (InvalidOperationException) { blocked = true; }
                Check(blocked && Convert.ToString(title.TextFrame.TextRange.Text) == "A concurrent user edit", "Concurrent edit was not preserved.");
                title.TextFrame.TextRange.Text = "Sales Q2";
                dynamic table = slide.Shapes.AddTable(3, 2, 60, 160, 370, 150);
                table.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Period";
                table.Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Units";
                table.Table.Cell(2, 1).Shape.TextFrame.TextRange.Text = "Q1";
                table.Table.Cell(2, 2).Shape.TextFrame.TextRange.Text = "100";
                table.Table.Cell(3, 1).Shape.TextFrame.TextRange.Text = "Q2";
                table.Table.Cell(3, 2).Shape.TextFrame.TextRange.Text = "120";
                dynamic chart = slide.Shapes.AddChart2(201, 51, 470, 150, 400, 250);
                if ((int)chart.Chart.SeriesCollection().Count == 0) chart.Chart.SeriesCollection().NewSeries();
                chart.Chart.SeriesCollection(1).Name = "Sales";
                chart.Chart.ChartData.Activate();
                dynamic dataWorkbook = chart.Chart.ChartData.Workbook;
                dynamic sourceSheet = dataWorkbook.Worksheets[1];
                sourceSheet.Cells[1, 1].Value2 = "Period"; sourceSheet.Cells[1, 2].Value2 = "Sales";
                sourceSheet.Cells[2, 1].Value2 = "Q1"; sourceSheet.Cells[2, 2].Value2 = 100d;
                sourceSheet.Cells[3, 1].Value2 = "Q2"; sourceSheet.Cells[3, 2].Value2 = 120d;
                chart.Chart.SetSourceData("'" + Convert.ToString(sourceSheet.Name).Replace("'", "''") + "'!$A$1:$B$3", 2);
                dataWorkbook.Close(true);
                var scenarios = new[] {
                    new Dictionary<string, object> { { "kind", "table_cell" }, { "shape_id", (int)table.Id }, { "row", 2 }, { "column", 2 }, { "before", "100" }, { "text", "125" } },
                    new Dictionary<string, object> { { "kind", "chart_point" }, { "shape_id", (int)chart.Id }, { "series", 1 }, { "category", 1 }, { "before_value", 100d }, { "value", 125d } },
                    new Dictionary<string, object> { { "kind", "annotate" }, { "shape_id", (int)table.Id }, { "row", 2 }, { "column", 2 } },
                    new Dictionary<string, object> { { "kind", "notes_append" }, { "notes", "Additional source reference" } },
                    new Dictionary<string, object> { { "kind", "move" }, { "new_index", 2 } },
                    new Dictionary<string, object> { { "kind", "insert" }, { "slide", new Dictionary<string, object> { { "id", "inserted" }, { "layout", "cover" }, { "title", "New analysis" } } } },
                    new Dictionary<string, object> { { "kind", "replace_slide" }, { "slide", new Dictionary<string, object> {
                        { "id", "recomposed" }, { "layout", "dual_visual" }, { "title", "Sales" }, { "subtitle", "Sales increased" },
                        { "table", new { headers = new[] { "Period", "Units" }, rows = new[] { new[] { "Q1", "100" }, new[] { "Q2", "120" } } } },
                        { "chart", new { type = "column", categories = new[] { "Q1", "Q2" }, series = new[] { new { name = "Sales", values = new[] { 100, 120 } } } } }
                    } } },
                    new Dictionary<string, object> { { "kind", "delete" } }
                };
                foreach (var scenario in scenarios)
                {
                    slide = deck.Slides[1];
                    var baseline = Content((object)slide); var otherBaseline = Content((object)deck.Slides[2]);
                    scenario["slide_id"] = (int)slide.SlideID; scenario["fingerprint"] = PresentationInspection.Fingerprint((object)slide);
                    var native = NewTransaction((object)deck);
                    Invoke(native, "Stage", (object)app, new object[] { scenario }); RetainTemps(native, temps);
                    Invoke(native, "Commit", (Action<string>)(status => { }));
                    if (Convert.ToString(scenario["kind"]) == "chart_point")
                    {
                        chart.Chart.ChartData.Activate(); dynamic dataCheck = chart.Chart.ChartData.Workbook;
                        Check(Convert.ToDouble(dataCheck.Worksheets[1].Cells[2, 2].Value2) == 125d, "Chart edit lost its embedded data association.");
                        dataCheck.Close(false);
                    }
                    var snapshot = Invoke(native, "Snapshot");
                    var recovered = Transaction.GetMethod("Recover", BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, new object[] { (object)app, (object)deck, snapshot });
                    Check(Convert.ToString(Invoke(recovered, "Reconcile")) == "applied", "Interrupted completed batch was not reconciled.");
                    Invoke(recovered, "Revert");
                    Check((int)deck.Slides.Count == 2, "Native " + scenario["kind"] + " left extra or missing slides.");
                    Check(Content((object)deck.Slides[1]) == baseline, "Native " + scenario["kind"] + " did not restore original content.");
                    Check(Content((object)deck.Slides[2]) == otherBaseline, "Native " + scenario["kind"] + " changed an unrelated slide.");
                }
                passed = true;
            }
            catch (Exception ex) { failure = ex.ToString(); }
            finally
            {
                foreach (dynamic temp in temps) try { temp.Close(); } catch { }
                if (deck != null) try { deck.Close(); } catch { }
            }
            var report = new { execution_kind = "native", policy = SamsungAuthoringPolicy.Version, assembly_sha256 = PresentationRevisionAcceptance.AssemblyHash(),
                revision_passed = passed, preservation_passed = passed, rollback_passed = passed, all_operations_passed = passed, full_acceptance_passed = false,
                note = "Native operation, preservation, rollback and concurrency checks only. Model/UI workflows and real Samsung fidelity require additional acceptance.", failure };
            var json = new JavaScriptSerializer().Serialize(report);
            File.WriteAllText(reportPath, json); Console.WriteLine(json);
            return passed ? 0 : 1;
        }
    }
}
