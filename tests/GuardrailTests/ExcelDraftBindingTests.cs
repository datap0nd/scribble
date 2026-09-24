using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Office;

namespace GuardrailTests
{
    public static class ExcelDraftBindingTests
    {
        public sealed class Application
        {
            public Workbook ActiveWorkbook { get; set; }
        }

        public sealed class Workbook
        {
            public string Name { get; set; }
            public string FullName { get; set; }
            public Sheet ActiveSheet { get; set; }
        }

        public sealed class Sheet
        {
            public string Name { get; set; }
            public Workbook Parent { get; set; }
        }

        public sealed class GridApplication
        {
            public GridSheet ActiveSheet { get; set; }
            public List<GridWorkbook> Workbooks { get; } =
                new List<GridWorkbook>();
        }

        public sealed class GridWorkbook
        {
            public string Name { get; set; } = "Grid.xlsx";
            public string FullName { get; set; } = "C:\\Grid.xlsx";
            public List<GridSheet> Worksheets { get; } =
                new List<GridSheet>();
        }

        public sealed class GridSheet
        {
            public string Name { get; set; } = "Data";
            public GridWorkbook Parent { get; set; }
            public GridCells Cells { get; } = new GridCells();
            public bool ProtectContents { get; set; }
            public GridAnchor Range(string name)
            {
                if (name != "A1") throw new Exception("Unexpected anchor");
                return new GridAnchor { Row = 1, Column = 1 };
            }
        }

        public sealed class GridAnchor
        {
            public int Row { get; set; }
            public int Column { get; set; }
        }

        public sealed class GridCells
        {
            private readonly Dictionary<string, GridCell> _cells =
                new Dictionary<string, GridCell>();
            public GridCell this[int row, int column]
            {
                get
                {
                    var key = row + ":" + column;
                    GridCell found;
                    if (!_cells.TryGetValue(key, out found))
                        _cells[key] = found = new GridCell();
                    return found;
                }
            }
        }

        public sealed class GridCell
        {
            private object _value;
            private object _formula;
            public bool FailNextWrite { get; set; }
            public bool FailNextRead { get; set; }
            public bool MergeCells { get; set; }
            public bool HasFormula { get; private set; }
            public object NumberFormat { get; set; } = "General";
            public object Value2
            {
                get
                {
                    if (FailNextRead)
                    {
                        FailNextRead = false;
                        throw new InvalidOperationException(
                            "Injected restart readback failure");
                    }
                    return _value;
                }
                set
                {
                    if (FailNextWrite)
                    {
                        FailNextWrite = false;
                        throw new InvalidOperationException(
                            "Injected second-cell write failure");
                    }
                    _value = value;
                    HasFormula = false;
                }
            }
            public object Formula
            {
                get { return _formula; }
                set { _formula = value; HasFormula = true; }
            }
        }

        private static object Invoke(object instance, string method,
            params object[] args)
        {
            var type = instance as Type ?? instance.GetType();
            try
            {
                return type.GetMethod(method, BindingFlags.Static |
                    BindingFlags.Instance | BindingFlags.NonPublic)
                    .Invoke(instance is Type ? null : instance, args);
            }
            catch (TargetInvocationException error)
            {
                System.Runtime.ExceptionServices.ExceptionDispatchInfo
                    .Capture(error.InnerException).Throw();
                throw;
            }
        }

        public static void FocusChangesStayBoundAndRenamesFail()
        {
            var first = new Workbook
            {
                Name = "Source.xlsx", FullName = "C:\\Source.xlsx"
            };
            first.ActiveSheet = new Sheet
            {
                Name = "Ledger", Parent = first
            };
            var capturedSheet = first.ActiveSheet;
            var second = new Workbook
            {
                Name = "Other.xlsx", FullName = "C:\\Other.xlsx"
            };
            second.ActiveSheet = new Sheet
            {
                Name = "Other", Parent = second
            };
            var application = new Application { ActiveWorkbook = first };
            var bindingType = typeof(SamsungAuthoringPolicy).Assembly.GetType(
                "Scribble.Office.ExcelDraftBinding", true);
            var binding = Invoke(bindingType, "Capture", application);
            application.ActiveWorkbook = second;
            first.ActiveSheet = second.ActiveSheet;
            if (!ReferenceEquals(Invoke(binding, "Workbook"), first) ||
                !ReferenceEquals(Invoke(binding, "Sheet"),
                    capturedSheet))
                throw new Exception(
                    "A focus change redirected the captured workbook or sheet.");
            first.Name = "Renamed.xlsx";
            try
            {
                Invoke(binding, "Workbook");
                throw new Exception("A renamed target remained writable.");
            }
            catch (InvalidOperationException error)
            {
                if (!error.Message.Contains("EXCEL_DRAFT_TARGET_CHANGED"))
                    throw;
            }
            application.ActiveWorkbook = null;
            var empty = Invoke(bindingType, "Capture", application);
            if (Invoke(empty, "Workbook") != null)
                throw new Exception("An empty request acquired a workbook.");
            try
            {
                Invoke(empty, "Sheet");
                throw new Exception("An absent worksheet remained writable.");
            }
            catch (InvalidOperationException error)
            {
                if (!error.Message.Contains("EXCEL_DRAFT_SHEET_UNAVAILABLE"))
                    throw;
            }
        }

        public static void OversizedRowsFailRatherThanDisappear()
        {
            var writer = typeof(SamsungAuthoringPolicy).Assembly.GetType(
                "Scribble.Office.WorkbookDraftWriter", true);
            Action<object, string> rejects = (rows, code) =>
            {
                try
                {
                    Invoke(writer, "ParseRows", rows);
                    throw new Exception("Oversized input was silently accepted: " +
                        code);
                }
                catch (InvalidOperationException error)
                {
                    if (!error.Message.Contains(code)) throw;
                }
            };
            rejects(Enumerable.Range(0, 201).Select(index =>
                (object)new object[] { "row " + index }).ToArray(),
                "DRAFT_ROWS_LIMIT");
            rejects(new object[] {
                Enumerable.Range(0, 31).Select(index =>
                    (object)("column " + index)).ToArray()
            }, "DRAFT_COLUMNS_LIMIT");
            rejects(new object[] { new object[] { new string('x', 501) } },
                "DRAFT_CELL_LIMIT");
        }

        public static void FailedCellWriteRestoresEarlierCells()
        {
            var workbook = new GridWorkbook();
            var sheet = new GridSheet { Parent = workbook };
            workbook.Worksheets.Add(sheet);
            sheet.Cells[1, 1].Value2 = "original A1";
            sheet.Cells[1, 2].Value2 = "original B1";
            sheet.Cells[1, 2].FailNextWrite = true;
            var application = new GridApplication { ActiveSheet = sheet };
            var writer = typeof(SamsungAuthoringPolicy).Assembly.GetType(
                "Scribble.Office.WorkbookDraftWriter", true);
            var method = writer.GetMethod("WriteCells",
                BindingFlags.Static | BindingFlags.NonPublic,
                null, new[] { typeof(object), typeof(string),
                    typeof(IReadOnlyList<IReadOnlyList<string>>),
                    typeof(object) }, null);
            var rows = new List<IReadOnlyList<string>>
            {
                new[] { "changed A1", "changed B1" }
            };
            try
            {
                method.Invoke(null, new object[] { application, "A1", rows,
                    sheet });
                throw new Exception("The injected failure was ignored.");
            }
            catch (TargetInvocationException error)
            {
                var inner = error.InnerException as InvalidOperationException;
                if (inner == null ||
                    !inner.Message.Contains("DRAFT_WRITE_ROLLED_BACK"))
                    throw;
            }
            if (Convert.ToString(sheet.Cells[1, 1].Value2) !=
                    "original A1" ||
                Convert.ToString(sheet.Cells[1, 2].Value2) !=
                    "original B1")
                throw new Exception("The bounded write lost a before-image.");
        }

        public static void RestoredEditDoesNotPoisonTaskRecovery()
        {
            var root = Path.Combine(Path.GetTempPath(),
                "scribble-cell-rollback-" + Guid.NewGuid().ToString("N"));
            try
            {
                var request = new ChatCompletionRequest
                {
                    model = "test",
                    messages = new List<object>
                    {
                        new ChatCompletionInputMessage
                        {
                            role = "user", content = "Change A1"
                        }
                    }
                };
                var task = new TaskContextManager(request, "excel",
                    "Change A1", new TaskCheckpointStore(root));
                var call = new ChatToolCall
                {
                    id = "rolled-back-cell", type = "function",
                    function = new ChatToolCallFunction
                    {
                        name = WorkbookToolCatalog.WriteCells,
                        arguments = "{\"start_cell\":\"A1\",\"rows\":[[\"x\"]]}"
                    }
                };
                task.BeforeTool(call, true);
                var result = new MailboxToolResult(call.id,
                    new JavaScriptSerializer().Serialize(new
                    {
                        ok = false,
                        error_code = "DRAFT_WRITE_ROLLED_BACK",
                        permission_consumed = true
                    }), "The captured cells were restored.");
                task.AfterTool(call, result);
                if (task.State.Writes.Single().Status != "verified" ||
                    task.State.HostData.Any(entry =>
                        entry.Value == "true" &&
                        entry.Key.Contains("write_cells")))
                    throw new Exception(
                        "A verified rollback was left as an uncertain write.");
            }
            finally
            {
                if (Directory.Exists(root)) Directory.Delete(root, true);
            }
        }

        public static void GridReceiptRetainsTypedBeforeImages()
        {
            var workbook = new GridWorkbook();
            var sheet = new GridSheet { Parent = workbook };
            workbook.Worksheets.Add(sheet);
            sheet.Cells[1, 1].Value2 = "001";
            sheet.Cells[1, 1].NumberFormat = "@";
            sheet.Cells[1, 2].Value2 = 12.5d;
            sheet.Cells[1, 2].NumberFormat = "0.00";
            var writer = typeof(SamsungAuthoringPolicy).Assembly.GetType(
                "Scribble.Office.WorkbookDraftWriter", true);
            var rows = new List<IReadOnlyList<string>>
            {
                new[] { "002", "14.5" }
            };
            var receipt = Invoke(writer, "CaptureCellsReceipt",
                new GridApplication { ActiveSheet = sheet }, "A1", rows,
                sheet, "bound-call");
            var json = new JavaScriptSerializer();
            var data = (Dictionary<string, object>)
                json.DeserializeObject(json.Serialize(receipt));
            var cells = (object[])data["Cells"];
            var first = (Dictionary<string, object>)cells[0];
            var second = (Dictionary<string, object>)cells[1];
            if (Convert.ToString(data["WorkbookFullName"]) !=
                    "C:\\Grid.xlsx" ||
                Convert.ToString(data["SheetName"]) != "Data" ||
                Convert.ToString(first["ValueKind"]) != "text" ||
                Convert.ToString(first["Value"]) != "001" ||
                Convert.ToString(first["NumberFormat"]) != "@" ||
                Convert.ToString(second["ValueKind"]) != "number" ||
                Convert.ToString(second["Value"]) != "12.5" ||
                Convert.ToString(second["Planned"]) != "14.5")
                throw new Exception(
                    "The durable before-image lost workbook identity or typed values.");
        }

        public static void InterruptedGridWriteReconcilesSafely()
        {
            foreach (var scenario in new[] {
                "before", "partial", "applied", "user_edit",
                "wrong_book", "read_error", "resume_write_error",
                "uncertain_rolled_back" })
            {
                var root = Path.Combine(Path.GetTempPath(),
                    "scribble-grid-recovery-" + Guid.NewGuid().ToString("N"));
                try
                {
                    var workbook = new GridWorkbook();
                    var sheet = new GridSheet { Parent = workbook };
                    workbook.Worksheets.Add(sheet);
                    sheet.Cells[1, 1].Value2 = "original A1";
                    sheet.Cells[1, 2].Value2 = "original B1";
                    var application = new GridApplication { ActiveSheet = sheet };
                    application.Workbooks.Add(workbook);
                    var writer = typeof(SamsungAuthoringPolicy).Assembly
                        .GetType("Scribble.Office.WorkbookDraftWriter", true);
                    var rows = new List<IReadOnlyList<string>>
                    {
                        new[] { "planned A1", "planned B1" }
                    };
                    const string callId = "interrupted-grid";
                    var receipt = Invoke(writer, "CaptureCellsReceipt",
                        application, "A1", rows, sheet, callId);
                    var request = new ChatCompletionRequest
                    {
                        model = "test",
                        messages = new List<object> { new ChatCompletionInputMessage
                            { role = "user", content = "Change A1 and B1" } }
                    };
                    var task = new TaskContextManager(request, "excel",
                        "Change A1 and B1", new TaskCheckpointStore(root));
                    var call = new ChatToolCall
                    {
                        id = callId, type = "function",
                        function = new ChatToolCallFunction
                        {
                            name = WorkbookToolCatalog.WriteCells,
                            arguments = "{\"start_cell\":\"A1\",\"rows\":[[\"planned A1\",\"planned B1\"]]}"
                        }
                    };
                    task.BeforeTool(call, true);
                    var receiptId = task.Store.PutEvidence(task.State.Id,
                        new JavaScriptSerializer().Serialize(receipt));
                    task.State.HostData["excel_grid_receipt"] = receiptId;
                    task.State.HostData["excel_grid_call_id"] = callId;
                    task.Checkpoint();
                    if (scenario == "uncertain_rolled_back")
                    {
                        task.AfterTool(call, new MailboxToolResult(call.id,
                            "{\"ok\":false,\"error_code\":\"DRAFT_WRITE_UNCERTAIN\",\"permission_consumed\":true}",
                            "Native outcome unknown"));
                        if (task.State.Writes.Single().Status != "uncertain" ||
                            !task.State.HostData.ContainsKey(
                                "generic_write_spent"))
                            throw new Exception("The interrupted attempt was not recorded as uncertain.");
                    }
                    if (scenario != "before")
                        sheet.Cells[1, 1].Value2 = "planned A1";
                    if (scenario == "applied")
                        sheet.Cells[1, 2].Value2 = "planned B1";
                    if (scenario == "user_edit")
                        sheet.Cells[1, 2].Value2 = "user edit";
                    if (scenario == "read_error")
                        sheet.Cells[1, 1].FailNextRead = true;
                    if (scenario == "resume_write_error")
                        sheet.Cells[1, 1].FailNextWrite = true;
                    if (scenario == "wrong_book")
                    {
                        workbook.Name = "Renamed.xlsx";
                        workbook.FullName = "C:\\Renamed.xlsx";
                    }
                    using (var host = new DocumentDraftHost("excel", application))
                        host.BindTaskAsync(task,
                            System.Threading.CancellationToken.None)
                            .GetAwaiter().GetResult();
                    var expectedStatus = scenario == "user_edit" ||
                        scenario == "wrong_book" ||
                        scenario == "read_error" ||
                        scenario == "resume_write_error"
                        ? "pending" : "verified";
                    var expectedFirst = scenario == "applied" ||
                        scenario == "user_edit" ||
                        scenario == "wrong_book" ||
                        scenario == "read_error" ||
                        scenario == "resume_write_error"
                        ? "planned A1" : "original A1";
                    var expectedSecond = scenario == "applied"
                        ? "planned B1" : scenario == "user_edit"
                            ? "user edit" : "original B1";
                    if (task.State.Writes.Single().Status != expectedStatus ||
                        Convert.ToString(sheet.Cells[1, 1].Value2) !=
                            expectedFirst ||
                        Convert.ToString(sheet.Cells[1, 2].Value2) !=
                            expectedSecond ||
                        (task.State.HostData.ContainsKey("excel_grid_receipt") !=
                            (expectedStatus == "pending")) ||
                        (task.State.HostData.ContainsKey(
                            "generic_write_spent") !=
                            (scenario == "applied")))
                        throw new Exception("Unsafe restart reconciliation: " +
                            scenario);
                }
                finally
                {
                    if (Directory.Exists(root)) Directory.Delete(root, true);
                }
            }
        }
    }
}
