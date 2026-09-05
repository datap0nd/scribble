using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Scribble.Security;

namespace Scribble.Office
{
    // Opens a brand-new, unsaved Excel workbook from outside the
    // Excel add-in process (the browser native host) and writes one
    // bounded table plus an optional native chart into it. The
    // workbook is only ever displayed for the user's review; this
    // class contains no save, print, protect, or close capability
    // and must never gain one.
    public static class ExcelTableLauncher
    {
        public const int MaxColumns = 20;
        public const int MaxRows = 500;
        public const int MaxCellCharacters = 500;
        public const int MaxTitleCharacters = 200;

        public static string OpenTable(
            string title,
            IReadOnlyList<string> columns,
            IReadOnlyList<IReadOnlyList<string>> rows,
            string chartKind,
            string chartTitle,
            Func<string, object> applicationResolver = null)
        {
            var safeColumns = BoundCells(columns, MaxColumns);
            if (safeColumns.Count == 0)
            {
                throw new InvalidOperationException(
                    "The table needs at least one column header.");
            }

            var safeRows = new List<IReadOnlyList<string>>();
            foreach (var row in rows ??
                new IReadOnlyList<string>[0])
            {
                safeRows.Add(BoundCells(row, safeColumns.Count));
                if (safeRows.Count == MaxRows)
                {
                    break;
                }
            }

            if (safeRows.Count == 0)
            {
                throw new InvalidOperationException(
                    "The table needs at least one data row.");
            }

            object application = null;
            object workbooks = null;
            object workbook = null;
            try
            {
                application = applicationResolver == null ? ResolveExcelApplication() : applicationResolver("Excel.Application");
                dynamic excel = application;
                excel.Visible = true;
                workbooks = excel.Workbooks;
                dynamic workbookCollection = workbooks;
                workbook = workbookCollection.Add();
                dynamic newWorkbook = workbook;
                dynamic sheet = newWorkbook.ActiveSheet;

                var safeTitle = TextBoundary.SingleLine(
                    title,
                    MaxTitleCharacters);
                var headerRow = 1;
                if (safeTitle.Length > 0)
                {
                    sheet.Cells[1, 1].Value2 = safeTitle;
                    sheet.Cells[1, 1].Font.Bold = true;
                    sheet.Cells[1, 1].Font.Size = 14;
                    headerRow = 3;
                }

                var values = new object[
                    safeRows.Count + 1,
                    safeColumns.Count];
                for (var column = 0;
                     column < safeColumns.Count;
                     column++)
                {
                    values[0, column] = safeColumns[column];
                }

                for (var row = 0; row < safeRows.Count; row++)
                {
                    for (var column = 0;
                         column < safeColumns.Count;
                         column++)
                    {
                        values[row + 1, column] =
                            CellValue(safeRows[row][column]);
                    }
                }

                dynamic tableRange = sheet.Cells[headerRow, 1]
                    .Resize[safeRows.Count + 1, safeColumns.Count];
                tableRange.Value2 = values;
                sheet.Cells[headerRow, 1]
                    .Resize[1, safeColumns.Count]
                    .Font.Bold = true;
                tableRange.Columns.AutoFit();

                var chartAdded = TryAddChart(
                    sheet,
                    tableRange,
                    headerRow + safeRows.Count + 2,
                    chartKind,
                    chartTitle);

                return
                    "A new unsaved Excel workbook is now open with the " +
                    safeRows.Count.ToString(
                        CultureInfo.InvariantCulture) +
                    "-row table" +
                    (chartAdded ? " and chart" : string.Empty) +
                    " for the user's review. Nothing was saved.";
            }
            finally
            {
                Release(workbook);
                Release(workbooks);
                Release(application);
            }
        }

        private static bool TryAddChart(
            dynamic sheet,
            dynamic tableRange,
            int anchorRow,
            string chartKind,
            string chartTitle)
        {
            var kind = TextBoundary.SingleLine(chartKind, 40)
                .ToLowerInvariant();
            if (kind.Length == 0 || kind == "none")
            {
                return false;
            }

            int chartType;
            switch (kind)
            {
                case "bar":
                    chartType = 57;
                    break;
                case "line":
                    chartType = 4;
                    break;
                case "pie":
                    chartType = 5;
                    break;
                default:
                    // Column chart for "column" and anything else.
                    chartType = 51;
                    break;
            }

            try
            {
                dynamic anchor = sheet.Cells[anchorRow, 1];
                dynamic shape = sheet.Shapes.AddChart2(
                    -1,
                    chartType,
                    (double)anchor.Left,
                    (double)anchor.Top,
                    440.0,
                    280.0);
                dynamic chart = shape.Chart;
                chart.SetSourceData(tableRange);
                var safeChartTitle = TextBoundary.SingleLine(
                    chartTitle,
                    MaxTitleCharacters);
                if (safeChartTitle.Length > 0)
                {
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = safeChartTitle;
                }

                return true;
            }
            catch
            {
                // A chart failure must never lose the table.
                return false;
            }
        }

        private static List<string> BoundCells(
            IReadOnlyList<string> cells,
            int width)
        {
            var result = new List<string>();
            foreach (var cell in cells ?? new string[0])
            {
                result.Add(TextBoundary.SingleLine(
                    cell,
                    MaxCellCharacters));
                if (result.Count == width)
                {
                    break;
                }
            }

            while (result.Count > 0 && result.Count < width)
            {
                result.Add(string.Empty);
            }

            return result;
        }

        // Numeric-looking cells become real numbers so charts and
        // formulas work on them.
        private static object CellValue(string value)
        {
            double numeric;
            if (double.TryParse(
                value,
                NumberStyles.Number,
                CultureInfo.InvariantCulture,
                out numeric))
            {
                return numeric;
            }

            return value;
        }

        private static object ResolveExcelApplication()
        {
            try
            {
                return Marshal.GetActiveObject("Excel.Application");
            }
            catch (COMException)
            {
                // Excel is not running; start a user-visible
                // instance so the workbook has a host.
            }

            var excelType = Type.GetTypeFromProgID(
                "Excel.Application");
            if (excelType == null)
            {
                throw new InvalidOperationException(
                    "Excel is not installed on this machine.");
            }

            return Activator.CreateInstance(excelType);
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                try
                {
                    Marshal.ReleaseComObject(value);
                }
                catch
                {
                    // Releasing a COM wrapper twice must never mask
                    // the table outcome.
                }
            }
        }
    }
}
