using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Scribble.Security;

namespace Scribble.Office
{
    // Reconciles a bounded in-place edit after the process stopped between
    // its encrypted before-image and its final tool receipt. No unknown cell
    // is overwritten. The workbook remains unsaved.
    internal static class ExcelGridWriteRecovery
    {
        internal const string Applied = "applied";
        internal const string RolledBack = "rolled_back";
        internal const string Uncertain = "uncertain";

        internal static string Reconcile(object excelApplication,
            ExcelGridWriteReceipt receipt)
        {
            if (receipt == null || receipt.Version != 1 ||
                string.IsNullOrWhiteSpace(receipt.CallId) ||
                string.IsNullOrWhiteSpace(receipt.WorkbookFullName) ||
                string.IsNullOrWhiteSpace(receipt.SheetName) ||
                receipt.Cells == null || receipt.Cells.Count == 0 ||
                receipt.Cells.Count > WorkbookDraftWriter.MaxDraftRows *
                    WorkbookDraftWriter.MaxDraftColumns ||
                receipt.Cells.Any(cell => cell.Row < 1 ||
                    cell.Row > ExcelSelectionOutputPolicy.MaxExcelRows ||
                    cell.Column < 1 ||
                    cell.Column > ExcelSelectionOutputPolicy.MaxExcelColumns))
                return Uncertain;
            try
            {
                dynamic application = excelApplication;
                dynamic workbook = null;
                foreach (dynamic candidate in application.Workbooks)
                    if (Convert.ToString(candidate.Name) ==
                            receipt.WorkbookName &&
                        Convert.ToString(candidate.FullName) ==
                            receipt.WorkbookFullName)
                    {
                        workbook = candidate;
                        break;
                    }
                if (workbook == null) return Uncertain;
                dynamic sheet = null;
                var sheetNames = new List<string>();
                foreach (dynamic candidate in workbook.Worksheets)
                {
                    var name = Convert.ToString(candidate.Name);
                    sheetNames.Add(name);
                    if (name == receipt.SheetName) sheet = candidate;
                }
                if (sheet == null ||
                    Convert.ToBoolean(sheet.ProtectContents))
                    return Uncertain;
                var planned = new List<ExcelGridCellReceipt>();
                var formulaError = false;
                foreach (var item in receipt.Cells)
                {
                    dynamic current = sheet.Cells[item.Row, item.Column];
                    if (Convert.ToBoolean(current.MergeCells) ||
                        Convert.ToString(current.NumberFormat) !=
                            item.NumberFormat)
                        return Uncertain;
                    if (BeforeMatches(current, item)) continue;
                    if (!PlannedMatches(current, item, sheetNames))
                        return Uncertain;
                    planned.Add(item);
                    if (Convert.ToBoolean(current.HasFormula) &&
                        ExcelErrorValue.Text((object)current.Value2) != null)
                        formulaError = true;
                }
                if (planned.Count == 0) return RolledBack;
                if (planned.Count == receipt.Cells.Count && !formulaError)
                    return Applied;
                // A mixture of original and planned cells, or a native
                // formula error, cannot be reported as a completed edit.
                // Recheck each cell immediately before restoring it.
                var restored = true;
                foreach (var item in planned.AsEnumerable().Reverse())
                {
                    dynamic current = sheet.Cells[item.Row, item.Column];
                    if (!PlannedMatches(current, item, sheetNames))
                    {
                        restored = false;
                        break;
                    }
                    Restore(current, item);
                    if (!BeforeMatches(current, item))
                    {
                        restored = false;
                        break;
                    }
                }
                return restored ? RolledBack : Uncertain;
            }
            catch
            {
                return Uncertain;
            }
        }

        private static bool BeforeMatches(dynamic current,
            ExcelGridCellReceipt item)
        {
            if (Convert.ToBoolean(current.HasFormula) != item.HasFormula)
                return false;
            if (item.HasFormula)
                return string.Equals(
                    Convert.ToString(current.Formula), item.Formula,
                    StringComparison.OrdinalIgnoreCase);
            object actual = current.Value2;
            switch (item.ValueKind)
            {
                case "empty": return actual == null ||
                    actual is string && (string)actual == string.Empty;
                case "text": return actual is string &&
                    (string)actual == item.Value;
                case "boolean": return actual is bool &&
                    (bool)actual == (item.Value == "true");
                case "number":
                    double expected;
                    return double.TryParse(item.Value,
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture, out expected) &&
                        actual is double && (double)actual == expected;
                default: return false;
            }
        }

        private static bool PlannedMatches(dynamic current,
            ExcelGridCellReceipt item, IEnumerable<string> sheetNames)
        {
            var proposed = item.Planned ?? string.Empty;
            if (proposed.StartsWith("=", StringComparison.Ordinal) &&
                DraftFormulaPolicy.IsAllowedFormula(proposed))
            {
                if (!Convert.ToBoolean(current.HasFormula))
                    return false;
                var normalized = WorkbookDraftWriter
                    .NormalizeSheetReferences(proposed, sheetNames);
                var actual = Convert.ToString(current.Formula);
                return string.Equals(actual, normalized,
                    StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(actual, proposed,
                        StringComparison.OrdinalIgnoreCase);
            }
            if (Convert.ToBoolean(current.HasFormula)) return false;
            object value = current.Value2;
            if (value == null)
                return proposed.Length == 0;
            if (value is string) return (string)value == proposed;
            if (value is bool)
                return string.Equals(proposed,
                    (bool)value ? "TRUE" : "FALSE",
                    StringComparison.OrdinalIgnoreCase);
            if (value is double)
                return Convert.ToString(value,
                    CultureInfo.InvariantCulture) == proposed;
            return false;
        }

        private static void Restore(dynamic cell,
            ExcelGridCellReceipt item)
        {
            cell.NumberFormat = item.NumberFormat;
            if (item.HasFormula)
            {
                cell.Formula = item.Formula;
                return;
            }
            switch (item.ValueKind)
            {
                case "empty": cell.Value2 = null; break;
                case "text": cell.Value2 = item.Value; break;
                case "boolean": cell.Value2 = item.Value == "true"; break;
                case "number": cell.Value2 = double.Parse(item.Value,
                    CultureInfo.InvariantCulture); break;
                default: throw new InvalidOperationException(
                    "DRAFT_BEFORE_IMAGE_UNSUPPORTED");
            }
        }
    }
}
