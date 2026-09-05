using System;
using System.Collections.Generic;
using System.Globalization;

namespace Scribble.Office
{
    internal sealed class ExcelTransformTarget : IExcelTransformTarget
    {
        private readonly object _application;
        private readonly ExcelSelectionSnapshot _selection;
        private readonly KoreanWorkbookSnapshot _korean;
        internal ExcelTransformTarget(object application, ExcelSelectionSnapshot selection, KoreanWorkbookSnapshot korean)
        {
            _application = application; _selection = selection; _korean = korean;
        }

        public IReadOnlyList<ExcelTaskCell> ReadSources(int offset, int count)
        {
            return Read(offset, count, null, true);
        }
        public IReadOnlyList<ExcelTaskCell> ReadDestination(int offset, int count, string column, bool replaceSource)
        {
            return Read(offset, count, column, replaceSource);
        }

        private IReadOnlyList<ExcelTaskCell> Read(int offset, int count, string column, bool source)
        {
            var result = new List<ExcelTaskCell>();
            if (_selection != null)
            {
                dynamic sheet = WorkbookSelectionOutputWriter.ResolveSheet(_application, _selection);
                var number = source ? _selection.StartColumn : ExcelSelectionOutputPolicy.ColumnNameToNumber(column);
                var first = ExcelSelectionOutputPolicy.ColumnNumberToName(number) + (_selection.StartRow + offset);
                var last = ExcelSelectionOutputPolicy.ColumnNumberToName(number) + (_selection.StartRow + offset + count - 1);
                dynamic range = sheet.Range(first + ":" + last);
                try
                {
                object values = range.Value2;
                object formulas = range.Formula;
                object hasFormula = range.HasFormula;
                object merged = range.MergeCells;
                for (var i = 0; i < count; i++)
                {
                    var address = ExcelSelectionOutputPolicy.ColumnNumberToName(number) + (_selection.StartRow + offset + i);
                    bool formula = hasFormula is bool ? (bool)hasFormula : Convert.ToBoolean(sheet.Range(address).HasFormula);
                    bool merge = merged is bool ? (bool)merged : Convert.ToBoolean(sheet.Range(address).MergeCells);
                    result.Add(new ExcelTaskCell { Id = "excel:" + _selection.WorksheetName + "!" + address,
                        Sheet = _selection.WorksheetName, Address = address, Value = Text(At(values, i)),
                        Formula = formula ? Text(At(formulas, i)) : "", HasFormula = formula, Merged = merge });
                }
                }
                finally { Release((object)range); Release((object)sheet); }
            }
            else
            {
                dynamic workbook = WorkbookSelectionOutputWriter.ResolveWorkbook(_application, _korean);
                try
                {
                for (var i = 0; i < count; i++)
                {
                    var original = _korean.Cells[offset + i];
                    dynamic sheet = WorkbookSelectionOutputWriter.FindWorksheet(workbook, original.WorksheetName);
                    dynamic cell = sheet.Range(original.Address);
                    try
                    {
                    bool formula = Convert.ToBoolean(cell.HasFormula);
                    result.Add(new ExcelTaskCell { Id = "excel:" + original.WorksheetName + "!" + original.Address,
                        Sheet = original.WorksheetName, Address = original.Address, Value = Text((object)cell.Value2),
                        Formula = formula ? Text((object)cell.Formula) : "", HasFormula = formula, Merged = Convert.ToBoolean(cell.MergeCells) });
                    }
                    finally { Release((object)cell); Release((object)sheet); }
                }
                }
                finally { Release((object)workbook); }
            }
            return result;
        }

        public void Write(int offset, IReadOnlyList<string> values, string column, bool replaceSource)
        {
            if (_selection != null)
            {
                dynamic sheet = WorkbookSelectionOutputWriter.ResolveSheet(_application, _selection);
                var number = replaceSource ? _selection.StartColumn : ExcelSelectionOutputPolicy.ColumnNameToNumber(column);
                var letter = ExcelSelectionOutputPolicy.ColumnNumberToName(number);
                dynamic range = sheet.Range(letter + (_selection.StartRow + offset) + ":" + letter + (_selection.StartRow + offset + values.Count - 1));
                try
                {
                var grid = new object[values.Count, 1];
                for (var i = 0; i < values.Count; i++) grid[i, 0] = values[i];
                range.NumberFormat = "@";
                range.Value2 = grid;
                }
                finally { Release((object)range); Release((object)sheet); }
            }
            else
            {
                dynamic workbook = WorkbookSelectionOutputWriter.ResolveWorkbook(_application, _korean);
                try
                {
                for (var i = 0; i < values.Count; i++)
                {
                    var original = _korean.Cells[offset + i];
                    dynamic sheet = WorkbookSelectionOutputWriter.FindWorksheet(workbook, original.WorksheetName);
                    dynamic cell = sheet.Range(original.Address);
                    try { cell.NumberFormat = "@"; cell.Value2 = values[i]; }
                    finally { Release((object)cell); Release((object)sheet); }
                }
                }
                finally { Release((object)workbook); }
            }
        }

        private static object At(object value, int index)
        {
            var grid = value as object[,];
            return grid == null ? value : grid[index + grid.GetLowerBound(0), grid.GetLowerBound(1)];
        }
        private static string Text(object value)
        {
            return value == null ? "" : Convert.ToString(value, CultureInfo.InvariantCulture);
        }
        private static void Release(object value)
        {
            if (value != null && System.Runtime.InteropServices.Marshal.IsComObject(value)) System.Runtime.InteropServices.Marshal.ReleaseComObject(value);
        }
    }
}
