using System;
using System.Collections.Generic;
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

                    originalSheet.Cells[6, 6].Value2 = "original first";
                    originalSheet.Cells[6, 7].Value2 = "original second";
                    source.Activate(); originalSheet.Activate();
                    capture();
                    other.Activate(); otherSheet.Activate();
                    var rollbackAuth = new OneShotDraftAuthorization(true);
                    var rollback = host.Execute(new ChatToolCall
                    {
                        id = "phase5-rollback", type = "function",
                        function = new ChatToolCallFunction
                        {
                            name = WorkbookToolCatalog.WriteCells,
                            arguments = json.Serialize(new
                            {
                                start_cell = "F6",
                                rows = new[] { new[] {
                                    "changed first", "=SUM(" } }
                            })
                        }
                    }, rollbackAuth, true, "Update F6:G6 in my sheet");
                    Check(rollback.Outcome.Failed &&
                        rollback.Outcome.ErrorCode ==
                            "DRAFT_WRITE_ROLLED_BACK" &&
                        rollbackAuth.IsConsumed &&
                        Convert.ToString(originalSheet.Cells[6, 6]
                            .Value2) == "original first" &&
                        Convert.ToString(originalSheet.Cells[6, 7]
                            .Value2) == "original second",
                        "EXCEL_ROLLBACK_LOST_SOURCE: " +
                        rollback.Content);

                    originalSheet.Cells[6, 8].Value2 = "original third";
                    originalSheet.Cells[6, 9].Value2 = "original fourth";
                    source.Activate(); originalSheet.Activate();
                    capture();
                    var errorAuth = new OneShotDraftAuthorization(true);
                    var formulaError = host.Execute(new ChatToolCall
                    {
                        id = "phase5-formula-error", type = "function",
                        function = new ChatToolCallFunction
                        {
                            name = WorkbookToolCatalog.WriteCells,
                            arguments = json.Serialize(new
                            {
                                start_cell = "H6",
                                rows = new[] { new[] {
                                    "changed third", "=1/0" } }
                            })
                        }
                    }, errorAuth, true, "Update H6:I6 in my sheet");
                    Check(formulaError.Outcome.Failed &&
                        formulaError.Outcome.ErrorCode ==
                            "DRAFT_WRITE_ROLLED_BACK" &&
                        Convert.ToString(originalSheet.Cells[6, 8]
                            .Value2) == "original third" &&
                        Convert.ToString(originalSheet.Cells[6, 9]
                            .Value2) == "original fourth",
                        "EXCEL_FORMULA_ERROR_WAS_ACCEPTED: " +
                        formulaError.Content);

                    originalSheet.Cells[12, 6].Value2 = 3d;
                    originalSheet.Cells[12, 7].Value2 = 4d;
                    source.Activate(); originalSheet.Activate();
                    capture();
                    other.Activate(); otherSheet.Activate();
                    var goodFormula = host.Execute(new ChatToolCall
                    {
                        id = "phase5-good-formula", type = "function",
                        function = new ChatToolCallFunction
                        {
                            name = WorkbookToolCatalog.WriteCells,
                            arguments = json.Serialize(new
                            {
                                start_cell = "H12",
                                rows = new[] { new[] {
                                    "=SUM($F$12:G12)" } }
                            })
                        }
                    }, new OneShotDraftAuthorization(true), true,
                        "Calculate H12 from F12 and G12");
                    Check(!goodFormula.Outcome.Failed &&
                        Convert.ToDouble(originalSheet.Cells[12, 8]
                            .Value2) == 7d &&
                        Convert.ToString(originalSheet.Cells[12, 8]
                            .Formula).Contains("$F$12:G12") &&
                        otherSheet.Cells[12, 8].Value2 == null,
                        "EXCEL_NATIVE_FORMULA_ADDRESS_FAILED: " +
                        goodFormula.Content);
                }
                VerifyInterruptedRecovery(app, source, other,
                    originalSheet);
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
                in_process_rollback_passed = passed,
                formula_error_rollback_passed = passed,
                formula_address_passed = passed,
                before_image_recovery_passed = passed,
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

        private static void VerifyInterruptedRecovery(dynamic app,
            dynamic source, dynamic other, dynamic sheet)
        {
            foreach (var scenario in new[] { "before", "partial",
                "applied", "user_edit" })
            {
                var root = Path.Combine(Path.GetTempPath(),
                    "scribble-native-grid-recovery-" +
                    Guid.NewGuid().ToString("N"));
                try
                {
                    sheet.Cells[10, 10].NumberFormat = "yyyy-mm-dd";
                    sheet.Cells[10, 10].Value2 = 45100d;
                    sheet.Cells[10, 11].Value2 = "original K10";
                    var writer = typeof(DocumentDraftHost).Assembly.GetType(
                        "Scribble.Office.WorkbookDraftWriter", true);
                    var capture = writer.GetMethod("CaptureCellsReceipt",
                        BindingFlags.Static | BindingFlags.NonPublic);
                    var rows = new List<IReadOnlyList<string>>
                    {
                        new[] { "revised date", "revised K10" }
                    };
                    var callId = "native-recovery-" + scenario;
                    var receipt = capture.Invoke(null, new object[] {
                        (object)app, "J10", rows, (object)sheet, callId });
                    var request = new ChatCompletionRequest
                    {
                        model = "test",
                        messages = new List<object> {
                            new ChatCompletionInputMessage {
                                role = "user", content = "Change J10 and K10" }
                        }
                    };
                    var task = new TaskContextManager(request, "excel",
                        "Change J10 and K10", new TaskCheckpointStore(root));
                    var call = new ChatToolCall
                    {
                        id = callId, type = "function",
                        function = new ChatToolCallFunction
                        {
                            name = WorkbookToolCatalog.WriteCells,
                            arguments = "{\"start_cell\":\"J10\",\"rows\":[[\"revised date\",\"revised K10\"]]}"
                        }
                    };
                    task.BeforeTool(call, true);
                    var receiptId = task.Store.PutEvidence(task.State.Id,
                        new JavaScriptSerializer().Serialize(receipt));
                    task.State.HostData["excel_grid_receipt"] = receiptId;
                    task.State.HostData["excel_grid_call_id"] = callId;
                    task.Checkpoint();
                    if (scenario != "before")
                        sheet.Cells[10, 10].Value2 = "revised date";
                    if (scenario == "applied")
                        sheet.Cells[10, 11].Value2 = "revised K10";
                    if (scenario == "user_edit")
                        sheet.Cells[10, 11].Value2 = "user edit";
                    other.Activate();
                    using (var host = new DocumentDraftHost("excel",
                        (object)app))
                        host.BindTaskAsync(task,
                            System.Threading.CancellationToken.None)
                            .GetAwaiter().GetResult();
                    var uncertain = scenario == "user_edit";
                    var applied = scenario == "applied";
                    Check(task.State.Writes.Single().Status ==
                            (uncertain ? "pending" : "verified") &&
                        Convert.ToString(sheet.Cells[10, 10].Value2) ==
                            (applied || uncertain ? "revised date" : "45100") &&
                        Convert.ToString(sheet.Cells[10, 11].Value2) ==
                            (applied ? "revised K10" : uncertain ?
                                "user edit" : "original K10") &&
                        Convert.ToString(sheet.Cells[10, 10].NumberFormat) ==
                            "yyyy-mm-dd" &&
                        (task.State.HostData.ContainsKey("generic_write_spent")
                            == applied) &&
                        (task.State.HostData.ContainsKey("excel_grid_receipt")
                            == uncertain),
                        "NATIVE_GRID_RECOVERY_FAILED: " + scenario);
                }
                finally
                {
                    if (Directory.Exists(root)) Directory.Delete(root, true);
                }
            }
        }
    }
}
