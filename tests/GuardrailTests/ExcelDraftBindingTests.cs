using System;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
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
        }

        public sealed class GridWorkbook
        {
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
            public bool MergeCells { get; set; }
            public bool HasFormula { get; private set; }
            public object NumberFormat { get; set; } = "General";
            public object Value2
            {
                get { return _value; }
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
    }
}
