using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Office;
using Scribble.Security;

namespace GuardrailTests
{
    // Explicit workstation run. Only disposable unsaved workbooks are used.
    internal static class Phase5NativeExcelAcceptance
    {
        internal static int Run(string reportPath)
        {
            dynamic app = null;
            dynamic source = null;
            dynamic other = null;
            var passed = false;
            var failure = string.Empty;
            var json = new JavaScriptSerializer();
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID(
                    "Excel.Application", true));
                app.Visible = true;
                app.DisplayAlerts = false;
                source = app.Workbooks.Add();
                dynamic originalSheet = source.ActiveSheet;
                originalSheet.Name = "Source";
                originalSheet.Cells[1, 1].Value2 = "source sentinel";
                other = app.Workbooks.Add();
                dynamic otherSheet = other.ActiveSheet;
                otherSheet.Name = "Other";
                otherSheet.Cells[1, 1].Value2 = "other sentinel";
                using (var host = new DocumentDraftHost("excel", (object)app))
                {
                    Action capture = () => typeof(DocumentDraftHost)
                        .GetMethod("BeginExcelDraftRequest",
                            BindingFlags.Instance | BindingFlags.NonPublic)
                        .Invoke(host, null);
                    source.Activate(); originalSheet.Activate();
                    capture();
                    other.Activate(); otherSheet.Activate();
                    var cellsAuth = new OneShotDraftAuthorization(true);
                    var cells = host.Execute(new ChatToolCall
                    {
                        id = "phase5-cells", type = "function",
                        function = new ChatToolCallFunction
                        {
                            name = WorkbookToolCatalog.WriteCells,
                            arguments = json.Serialize(new
                            {
                                start_cell = "B2",
                                rows = new[] { new[] { "bound cell" } }
                            })
                        }
                    }, cellsAuth, true, "Update B2 in my sheet");
                    Check(!cells.Outcome.Failed && cellsAuth.IsCreated &&
                        Convert.ToString(originalSheet.Cells[2, 2].Value2) ==
                            "bound cell" &&
                        otherSheet.Cells[2, 2].Value2 == null,
                        "EXCEL_BOUND_CELLS_REDIRECTED: " + cells.Content);

                    source.Activate(); originalSheet.Activate();
                    capture();
                    other.Activate(); otherSheet.Activate();
                    var draftAuth = new OneShotDraftAuthorization(true);
                    var draft = host.Execute(new ChatToolCall
                    {
                        id = "phase5-draft", type = "function",
                        function = new ChatToolCallFunction
                        {
                            name = WorkbookToolCatalog.WriteDraftSheet,
                            arguments = json.Serialize(new
                            {
                                title = "Bound report",
                                rows = new[] { new[] { "Measure" },
                                    new[] { "Verified" } }
                            })
                        }
                    }, draftAuth, true, "Create a draft sheet in my workbook");
                    Check(!draft.Outcome.Failed && draftAuth.IsCreated &&
                        Enumerable.Range(1, (int)source.Worksheets.Count)
                            .Any(index => Convert.ToString(
                                source.Worksheets[index].Name) ==
                                "Scribble Draft") &&
                        !Enumerable.Range(1, (int)other.Worksheets.Count)
                            .Any(index => Convert.ToString(
                                other.Worksheets[index].Name) ==
                                "Scribble Draft"),
                        "EXCEL_BOUND_DRAFT_REDIRECTED: " + draft.Content);

                    source.Activate(); originalSheet.Activate();
                    capture();
                    originalSheet.Name = "Renamed";
                    var rejectedAuth = new OneShotDraftAuthorization(true);
                    var rejected = host.Execute(new ChatToolCall
                    {
                        id = "phase5-rename", type = "function",
                        function = new ChatToolCallFunction
                        {
                            name = WorkbookToolCatalog.WriteCells,
                            arguments = json.Serialize(new
                            {
                                start_cell = "C2",
                                rows = new[] { new[] { "must not write" } }
                            })
                        }
                    }, rejectedAuth, true, "Update C2 in my sheet");
                    Check(rejected.Outcome.Failed &&
                        rejected.Content.Contains(
                            "EXCEL_DRAFT_TARGET_CHANGED") &&
                        !rejectedAuth.IsConsumed &&
                        originalSheet.Cells[2, 3].Value2 == null &&
                        Convert.ToString(originalSheet.Cells[1, 1].Value2) ==
                            "source sentinel" &&
                        Convert.ToString(otherSheet.Cells[1, 1].Value2) ==
                            "other sentinel",
                        "EXCEL_BOUND_RENAME_DID_NOT_FAIL_CLOSED: " +
                        rejected.Content);

                    originalSheet.Name = "Source";
                    originalSheet.Range("D3:E3").Merge();
                    source.Activate(); originalSheet.Activate();
                    capture();
                    other.Activate(); otherSheet.Activate();
                    var mergeAuth = new OneShotDraftAuthorization(true);
                    var merged = host.Execute(new ChatToolCall
                    {
                        id = "phase5-merged", type = "function",
                        function = new ChatToolCallFunction
                        {
                            name = WorkbookToolCatalog.WriteCells,
                            arguments = json.Serialize(new
                            {
                                start_cell = "C2",
                                rows = new[] { new[] { "early", "safe" },
                                    new[] { "later", "merged" } }
                            })
                        }
                    }, mergeAuth, true, "Update C2:D3 in my sheet");
                    Check(merged.Outcome.Failed &&
                        merged.Content.Contains("DRAFT_TARGET_MERGED") &&
                        !mergeAuth.IsConsumed &&
                        originalSheet.Cells[2, 3].Value2 == null &&
                        originalSheet.Cells[2, 4].Value2 == null,
                        "EXCEL_MERGED_PREFLIGHT_LEFT_PARTIAL_WRITE: " +
                        merged.Content);
                }
                passed = true;
            }
            catch (Exception error) { failure = error.ToString(); }
            finally
            {
                if (other != null) try { other.Close(false); } catch { }
                if (source != null) try { source.Close(false); } catch { }
                if (app != null) try
                {
                    if ((int)app.Workbooks.Count == 0) app.Quit();
                }
                catch { }
            }
            Directory.CreateDirectory(Path.GetDirectoryName(
                Path.GetFullPath(reportPath)));
            File.WriteAllText(reportPath, json.Serialize(new
            {
                execution_kind = "native_disposable_phase5_excel_binding",
                target_binding_passed = passed,
                merged_preflight_passed = passed,
                before_image_recovery_passed = false,
                full_acceptance_passed = false,
                failure
            }));
            Console.WriteLine(File.ReadAllText(reportPath));
            return passed ? 0 : 1;
        }

        private static void Check(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }
    }
}
