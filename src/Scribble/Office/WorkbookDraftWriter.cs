using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Scribble.Security;

namespace Scribble.Office
{
    // The Excel draft write surface of the suite. Every draft call
    // lands on its own numbered "Scribble Draft" worksheet appended at
    // the end of the workbook; earlier drafts and the user's own
    // sheets are never touched, and the workbook is never saved -
    // saving stays a human action. Cells starting with '=' become
    // live formulas only when DraftFormulaPolicy allows them (no
    // network, native-code, or external-workbook functions);
    // everything else lands as text.
    internal static class WorkbookDraftWriter
    {
        internal const string DraftSheetName = "Scribble Draft";
        internal const int MaxDraftRows = 200;
        internal const int MaxDraftColumns = 30;
        internal const int MaxCellCharacters = 500;

        internal static string WriteDraftSheet(
            object excelApplication,
            string title,
            IReadOnlyList<IReadOnlyList<string>> rows)
        {
            return WriteDraftSheet(
                excelApplication,
                title,
                rows,
                null);
        }

        internal static string WriteDraftSheet(
            object excelApplication,
            string title,
            IReadOnlyList<IReadOnlyList<string>> rows,
            DraftSheetChart chart)
        {
            return WriteDraftSheet(
                excelApplication,
                title,
                rows,
                chart,
                false);
        }

        // inNewWorkbook forces a brand-new unsaved workbook - the
        // cross-app send tools use it so content handed over from
        // another app never lands quietly inside whatever workbook
        // happens to be open.
        internal static string WriteDraftSheet(
            object excelApplication,
            string title,
            IReadOnlyList<IReadOnlyList<string>> rows,
            DraftSheetChart chart,
            bool inNewWorkbook)
        {
            if (rows == null || rows.Count == 0)
            {
                throw new InvalidOperationException(
                    "The draft table needs at least one row.");
            }

            dynamic application = excelApplication;
            dynamic workbook = inNewWorkbook
                ? null
                : application.ActiveWorkbook;
            if (workbook == null)
            {
                workbook = application.Workbooks.Add();
                Scribble.Testing.TestLab.RegisterOutput((object)workbook, "xlsx");
            }

            if (inNewWorkbook)
            {
                try
                {
                    application.Visible = true;
                    workbook.Activate();
                }
                catch
                {
                }
            }

            // Every draft call gets its own numbered sheet
            // ('Scribble Draft', 'Scribble Draft 2', ...): an earlier
            // draft is NEVER overwritten, so a follow-up request
            // ("now a summary table with charts") can never destroy
            // the previous result - Excel cannot undo add-in
            // writes, so keeping every draft is the undo.
            var existingNames = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            foreach (dynamic candidate in workbook.Worksheets)
            {
                existingNames.Add(
                    Convert.ToString(candidate.Name) ??
                    string.Empty);
            }

            var sheetName = DraftSheetName;
            var suffix = 2;
            while (existingNames.Contains(sheetName))
            {
                if (suffix > 99)
                {
                    throw new InvalidOperationException(
                        "Too many Scribble Draft sheets - delete " +
                        "some drafts and try again.");
                }

                sheetName = DraftSheetName + " " + suffix;
                suffix++;
            }

            dynamic sheets = workbook.Worksheets;
            dynamic sheet = sheets.Add(
                Type.Missing,
                sheets[sheets.Count]);
            sheet.Name = sheetName;

            // The layout is deterministic so model formulas can
            // reference the draft table itself: title in A1, table
            // always starting at A3 (header row 3, first data row
            // 4).
            var boundedTitle = TextBoundary.SingleLine(
                SafeModelText.Format(title, 180).PlainText,
                180);
            const int startRow = 3;
            if (boundedTitle.Length > 0)
            {
                sheet.Cells[1, 1].Value2 = boundedTitle;
            }

            var rowCount = Math.Min(rows.Count, MaxDraftRows);
            var columnCount = 1;
            for (var index = 0; index < rowCount; index++)
            {
                if (rows[index] != null &&
                    rows[index].Count > columnCount)
                {
                    columnCount = Math.Min(
                        rows[index].Count,
                        MaxDraftColumns);
                }
            }

            // Formula cells stay out of the bulk write: they are
            // set one by one below so a rejected or broken formula
            // degrades to text without failing the whole draft.
            var formulas =
                new List<KeyValuePair<int[], string>>();
            var grid = new object[rowCount, columnCount];
            for (var row = 0; row < rowCount; row++)
            {
                var source = rows[row];
                for (var column = 0;
                     column < columnCount;
                     column++)
                {
                    var cell =
                        source != null && column < source.Count
                            ? TextBoundary.SingleLine(
                                source[column],
                                MaxCellCharacters)
                            : string.Empty;
                    if (cell.Length > 0 &&
                        cell[0] != '=' &&
                        cell.IndexOf(
                            "**",
                            StringComparison.Ordinal) >= 0)
                    {
                        // Text cells never show literal bold
                        // markers; formulas stay untouched.
                        cell = TextBoundary.SingleLine(
                            SafeModelText.Format(
                                cell,
                                MaxCellCharacters).PlainText,
                            MaxCellCharacters);
                    }

                    if (cell.Length > 0 && cell[0] == '=')
                    {
                        if (DraftFormulaPolicy.IsAllowedFormula(
                            cell))
                        {
                            grid[row, column] = string.Empty;
                            formulas.Add(
                                new KeyValuePair<int[], string>(
                                    new[] { row, column },
                                    cell));
                            continue;
                        }

                        // The apostrophe keeps blocked formula
                        // text visible as plain text.
                        cell = "'" + cell;
                    }

                    grid[row, column] = cell;
                }
            }

            dynamic target = sheet.Range(
                sheet.Cells[startRow, 1],
                sheet.Cells[
                    startRow + rowCount - 1,
                    columnCount]);
            target.Value2 = grid;
            var formulaCount = 0;
            var liveFormulas =
                new List<KeyValuePair<int[], string>>();
            foreach (var formula in formulas)
            {
                dynamic cell = sheet.Cells[
                    startRow + formula.Key[0],
                    formula.Key[1] + 1];
                try
                {
                    cell.Formula = formula.Value;
                    formulaCount++;
                    liveFormulas.Add(formula);
                }
                catch
                {
                    // Models sometimes emit locale-style formulas
                    // (semicolon argument separators); FormulaLocal
                    // accepts the current locale's syntax before
                    // the final degrade-to-text fallback.
                    try
                    {
                        cell.FormulaLocal = formula.Value;
                        formulaCount++;
                        liveFormulas.Add(formula);
                    }
                    catch
                    {
                        try
                        {
                            cell.Value2 = "'" + formula.Value;
                        }
                        catch
                        {
                        }
                    }
                }
            }

            // A formula that parses but evaluates to an Excel error is
            // definitely wrong: repair a one-row header offset when safe,
            // otherwise degrade it to visible text so the draft never
            // shows a silently broken live formula.
            var brokenFormulas = 0;
            if (liveFormulas.Count > 0)
            {
                try
                {
                    sheet.Calculate();
                }
                catch
                {
                }

                foreach (var formula in liveFormulas)
                {
                    try
                    {
                        dynamic cell = sheet.Cells[
                            startRow + formula.Key[0],
                            formula.Key[1] + 1];
                        object value = cell.Value2;
                        // Excel error cells marshal as Int32 error
                        // codes; real numbers arrive as Double.
                        if (IsExcelError(value))
                        {
                            if (TryRepairAdjacentRowFormula(
                                sheet,
                                cell,
                                formula.Value,
                                startRow + formula.Key[0]))
                                continue;
                            cell.Value2 = "'" + formula.Value;
                            formulaCount--;
                            brokenFormulas++;
                        }
                    }
                    catch
                    {
                    }
                }
            }

            ApplyDraftFormatting(
                sheet,
                boundedTitle,
                startRow,
                rowCount,
                target);
            var chartAdded =
                chart != null &&
                AddDraftChart(
                    sheet,
                    chart,
                    startRow,
                    rowCount,
                    columnCount,
                    rows,
                    target);
            try
            {
                sheet.Activate();
            }
            catch
            {
            }

            return "Wrote " + rowCount + " rows x " +
                columnCount + " columns" +
                (formulaCount > 0
                    ? " including " + formulaCount +
                      " live formulas"
                    : string.Empty) +
                (chartAdded
                    ? " and a native chart"
                    : string.Empty) +
                " to the new '" +
                sheetName +
                "' sheet" +
                (inNewWorkbook
                    ? " of a new unsaved draft workbook"
                    : string.Empty) +
                ". Earlier draft sheets were left untouched." +
                (brokenFormulas > 0
                    ? " " + brokenFormulas +
                      (brokenFormulas == 1
                          ? " formula evaluated"
                          : " formulas evaluated") +
                      " to an Excel error and " +
                      (brokenFormulas == 1 ? "was" : "were") +
                      " kept as visible text - check the " +
                      "function names and sheet references."
                    : string.Empty) +
                " Nothing was saved.";
        }

        // Writes a bounded grid into the ACTIVE worksheet starting
        // at the given A1-style cell - the user explicitly asked to
        // work on their own sheet. Existing cells in the target
        // area are overwritten in memory; nothing is ever saved,
        // but Excel cannot undo add-in changes, so the draft sheet
        // stays the default surface.
        internal static string WriteCells(
            object excelApplication,
            string startCell,
            IReadOnlyList<IReadOnlyList<string>> rows)
        {
            if (rows == null || rows.Count == 0)
            {
                throw new InvalidOperationException(
                    "At least one row of cells is required.");
            }

            var anchorName = TextBoundary.SingleLine(
                startCell,
                12).Replace("$", string.Empty);
            if (!IsCellName(anchorName))
            {
                throw new InvalidOperationException(
                    "start_cell must be a single A1-style cell " +
                    "such as B2.");
            }

            dynamic application = excelApplication;
            dynamic sheet = application.ActiveSheet;
            if (sheet == null)
            {
                throw new InvalidOperationException(
                    "No worksheet is active.");
            }

            dynamic anchor = sheet.Range(anchorName);
            int startRow = anchor.Row;
            int startColumn = anchor.Column;
            var rowCount = Math.Min(rows.Count, MaxDraftRows);
            var written = 0;
            var formulaCount = 0;
            var brokenFormulas = 0;
            var liveFormulas =
                new List<KeyValuePair<int[], string>>();
            for (var row = 0; row < rowCount; row++)
            {
                var source = rows[row];
                if (source == null)
                {
                    continue;
                }

                var columnCount = Math.Min(
                    source.Count,
                    MaxDraftColumns);
                for (var column = 0;
                     column < columnCount;
                     column++)
                {
                    var cell = TextBoundary.SingleLine(
                        source[column],
                        MaxCellCharacters);
                    if (cell.Length > 0 &&
                        cell[0] != '=' &&
                        cell.IndexOf(
                            "**",
                            StringComparison.Ordinal) >= 0)
                    {
                        cell = TextBoundary.SingleLine(
                            SafeModelText.Format(
                                cell,
                                MaxCellCharacters).PlainText,
                            MaxCellCharacters);
                    }

                    dynamic target = sheet.Cells[
                        startRow + row,
                        startColumn + column];
                    if (cell.Length > 0 && cell[0] == '=')
                    {
                        if (!DraftFormulaPolicy.IsAllowedFormula(
                            cell))
                        {
                            target.Value2 = "'" + cell;
                            written++;
                            continue;
                        }

                        try
                        {
                            target.Formula = cell;
                            formulaCount++;
                            liveFormulas.Add(
                                new KeyValuePair<int[], string>(
                                    new[]
                                    {
                                        startRow + row,
                                        startColumn + column
                                    },
                                    cell));
                        }
                        catch
                        {
                            try
                            {
                                target.FormulaLocal = cell;
                                formulaCount++;
                                liveFormulas.Add(
                                    new KeyValuePair<int[], string>(
                                        new[]
                                        {
                                            startRow + row,
                                            startColumn + column
                                        },
                                        cell));
                            }
                            catch
                            {
                                try
                                {
                                    target.Value2 = "'" + cell;
                                }
                                catch
                                {
                                }
                            }
                        }

                        written++;
                        continue;
                    }

                    target.Value2 = cell;
                    written++;
                }
            }

            if (liveFormulas.Count > 0)
            {
                try
                {
                    sheet.Calculate();
                }
                catch
                {
                }

                foreach (var formula in liveFormulas)
                {
                    try
                    {
                        dynamic cell = sheet.Cells[
                            formula.Key[0],
                            formula.Key[1]];
                        object value = cell.Value2;
                        if (IsExcelError(value))
                        {
                            if (TryRepairAdjacentRowFormula(
                                sheet,
                                cell,
                                formula.Value,
                                formula.Key[0]))
                                continue;
                            cell.Value2 = "'" + formula.Value;
                            formulaCount--;
                            brokenFormulas++;
                        }
                    }
                    catch
                    {
                    }
                }
            }

            return "Wrote " + written + " cells starting at " +
                anchorName + " on the active sheet" +
                (formulaCount > 0
                    ? " including " + formulaCount +
                      " live formulas"
                    : string.Empty) +
                "." +
                (brokenFormulas > 0
                    ? " " + brokenFormulas +
                      " formula(s) evaluated to an Excel error " +
                      "and were kept as visible text."
                    : string.Empty) +
                " Nothing was saved, but Excel cannot undo " +
                "add-in changes - close without saving to " +
                "discard.";
        }

        private static bool IsExcelError(object value)
        {
            if (!(value is int)) return false;
            return new[] { 2000, 2007, 2015, 2023, 2029, 2036, 2042 }
                .Contains((int)value);
        }

        private static bool TryRepairAdjacentRowFormula(
            dynamic sheet,
            dynamic cell,
            string formula,
            int targetRow)
        {
            try
            {
                var references = Regex.Matches(
                    formula ?? "",
                    @"(?<![A-Za-z0-9_])(\$?[A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})");
                if (references.Count == 0) return false;
                var rows = references.Cast<Match>()
                    .Select(match => Convert.ToInt32(match.Groups[3].Value))
                    .Distinct()
                    .ToArray();
                if (rows.Length != 1 ||
                    Math.Abs(rows[0] - targetRow) != 1)
                    return false;
                var repaired = Regex.Replace(
                    formula,
                    @"(?<![A-Za-z0-9_])(\$?[A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})",
                    match => match.Groups[1].Value +
                        match.Groups[2].Value +
                        targetRow.ToString(CultureInfo.InvariantCulture));
                cell.Formula = repaired;
                sheet.Calculate();
                if (!IsExcelError((object)cell.Value2)) return true;
                cell.Formula = formula;
                sheet.Calculate();
            }
            catch
            {
                try { cell.Formula = formula; } catch { }
            }
            return false;
        }

        private static bool IsCellName(string value)
        {
            var letters = 0;
            var digits = 0;
            foreach (var character in value)
            {
                if (char.IsLetter(character) && digits == 0)
                {
                    letters++;
                }
                else if (char.IsDigit(character) && letters > 0)
                {
                    digits++;
                }
                else
                {
                    return false;
                }
            }

            return letters >= 1 &&
                   letters <= 3 &&
                   digits >= 1 &&
                   digits <= 7;
        }

        // The chart definition can name one rectangular section of a
        // multi-section draft. If omitted, the writer finds the last
        // chartable contiguous table instead of charting the whole report.
        internal sealed class DraftSheetChart
        {
            internal DraftSheetChart(
                int typeCode,
                string title,
                string sourceRange)
            {
                TypeCode = typeCode;
                Title = title ?? string.Empty;
                SourceRange = sourceRange ?? string.Empty;
            }

            internal int TypeCode { get; }

            internal string Title { get; }

            internal string SourceRange { get; }
        }

        // Reads the optional chart argument of write_draft_sheet;
        // anything malformed simply yields no chart.
        internal static DraftSheetChart ParseChart(object value)
        {
            var map = value as IDictionary<string, object>;
            if (map == null)
            {
                return null;
            }

            object typeValue;
            object titleValue;
            object rangeValue;
            map.TryGetValue("type", out typeValue);
            map.TryGetValue("title", out titleValue);
            map.TryGetValue("range", out rangeValue);
            return new DraftSheetChart(
                DraftChartTypes.Resolve(
                    Convert.ToString(typeValue)),
                TextBoundary.SingleLine(
                    SafeModelText.Format(
                        Convert.ToString(titleValue),
                        180).PlainText,
                    180),
                TextBoundary.SingleLine(
                    Convert.ToString(rangeValue),
                    32));
        }

        // Draws a native chart below the table on the draft sheet,
        // sourced live from the table range. Chart failures never
        // fail the draft - the table is the primary deliverable.
        private static bool AddDraftChart(
            dynamic sheet,
            DraftSheetChart chart,
            int startRow,
            int rowCount,
            int columnCount,
            IReadOnlyList<IReadOnlyList<string>> rows,
            dynamic target)
        {
            try
            {
                dynamic anchor = sheet.Cells[
                    startRow + rowCount + 2,
                    1];
                dynamic chartObject = sheet.ChartObjects().Add(
                    (double)anchor.Left,
                    (double)anchor.Top,
                    440.0,
                    270.0);
                chartObject.Chart.SetSourceData(
                    SelectChartRange(
                        sheet,
                        chart,
                        startRow,
                        rowCount,
                        columnCount,
                        rows,
                        target));
                chartObject.Chart.ChartType = chart.TypeCode;
                if (chart.Title.Length > 0)
                {
                    try
                    {
                        chartObject.Chart.HasTitle = true;
                        chartObject.Chart.ChartTitle.Text =
                            chart.Title;
                    }
                    catch
                    {
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static dynamic SelectChartRange(
            dynamic sheet,
            DraftSheetChart chart,
            int startRow,
            int rowCount,
            int columnCount,
            IReadOnlyList<IReadOnlyList<string>> rows,
            dynamic fallback)
        {
            if (Regex.IsMatch(
                chart.SourceRange,
                @"^[A-Za-z]{1,3}[1-9][0-9]{0,6}:[A-Za-z]{1,3}[1-9][0-9]{0,6}$"))
            {
                try
                {
                    dynamic selected = sheet.Range(chart.SourceRange);
                    var lastRow = (int)selected.Row +
                        (int)selected.Rows.Count - 1;
                    var lastColumn = (int)selected.Column +
                        (int)selected.Columns.Count - 1;
                    if ((int)selected.Row >= startRow &&
                        (int)selected.Column >= 1 &&
                        lastRow < startRow + rowCount &&
                        lastColumn <= columnCount)
                        return selected;
                }
                catch
                {
                }
            }

            var blocks = new List<int[]>();
            var blockStart = -1;
            for (var row = 0; row <= rowCount; row++)
            {
                var blank = row == rowCount || IsBlankRow(rows[row]);
                if (!blank && blockStart < 0) blockStart = row;
                if (blank && blockStart >= 0)
                {
                    if (row - blockStart >= 2 &&
                        IsChartableBlock(rows, blockStart, row))
                        blocks.Add(new[] { blockStart, row });
                    blockStart = -1;
                }
            }
            if (blocks.Count == 0) return fallback;
            var block = blocks[blocks.Count - 1];
            var lastColumnIndex = 1;
            for (var row = block[0]; row < block[1]; row++)
                if (rows[row] != null)
                    for (var column = 1;
                         column < Math.Min(rows[row].Count, columnCount);
                         column++)
                        if (!string.IsNullOrWhiteSpace(rows[row][column]))
                            lastColumnIndex = Math.Max(lastColumnIndex, column);
            return sheet.Range(
                sheet.Cells[startRow + block[0], 1],
                sheet.Cells[startRow + block[1] - 1, lastColumnIndex + 1]);
        }

        private static bool IsBlankRow(IReadOnlyList<string> row)
        {
            return row == null || row.All(string.IsNullOrWhiteSpace);
        }

        private static bool IsChartableBlock(
            IReadOnlyList<IReadOnlyList<string>> rows,
            int start,
            int end)
        {
            if (rows[start] == null ||
                rows[start].Count < 2 ||
                string.IsNullOrWhiteSpace(rows[start][0]))
                return false;
            var numericRows = 0;
            for (var row = start + 1; row < end; row++)
            {
                var values = rows[row];
                if (values == null || values.Count < 2) continue;
                if (values.Skip(1).Any(IsChartValue)) numericRows++;
            }
            return numericRows > 0;
        }

        private static bool IsChartValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return false;
            if (value.TrimStart().StartsWith("=", StringComparison.Ordinal))
                return true;
            double number;
            return double.TryParse(
                value.Replace(",", "")
                    .Replace("%", "")
                    .Replace("€", "")
                    .Trim(),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out number);
        }

        // Cosmetic polish for the draft sheet: bold title, bold
        // header row with a divider, and autofitted columns. Any
        // failure here must never fail the draft itself.
        private static void ApplyDraftFormatting(
            dynamic sheet,
            string boundedTitle,
            int startRow,
            int rowCount,
            dynamic target)
        {
            try
            {
                if (boundedTitle.Length > 0)
                {
                    dynamic titleCell = sheet.Cells[1, 1];
                    titleCell.Font.Bold = true;
                    titleCell.Font.Size = 12;
                }

                if (rowCount > 1)
                {
                    dynamic header = target.Rows[1];
                    header.Font.Bold = true;
                    // 9 = xlEdgeBottom, 1 = xlContinuous.
                    header.Borders[9].LineStyle = 1;
                }

                target.EntireColumn.AutoFit();
            }
            catch
            {
            }
        }

        // Converts the model-supplied JSON rows value into bounded
        // string rows, rejecting anything but arrays of arrays.
        internal static IReadOnlyList<IReadOnlyList<string>> ParseRows(
            object value)
        {
            var outer = AsEnumerable(value);
            if (outer == null)
            {
                throw new InvalidOperationException(
                    "rows must be an array of arrays of strings.");
            }

            var rows = new List<IReadOnlyList<string>>();
            foreach (var rowValue in outer)
            {
                if (rows.Count == MaxDraftRows)
                {
                    break;
                }

                var inner = AsEnumerable(rowValue);
                if (inner == null)
                {
                    throw new InvalidOperationException(
                        "rows must be an array of arrays of strings.");
                }

                var cells = new List<string>();
                foreach (var cell in inner)
                {
                    if (cells.Count == MaxDraftColumns)
                    {
                        break;
                    }

                    cells.Add(TextBoundary.SingleLine(
                        Convert.ToString(cell),
                        MaxCellCharacters));
                }

                rows.Add(cells);
            }

            return rows;
        }

        private static IEnumerable AsEnumerable(object value)
        {
            if (value == null || value is string)
            {
                return null;
            }

            return value as IEnumerable;
        }
    }
}
