using System;
using System.Linq;
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
    }
}
