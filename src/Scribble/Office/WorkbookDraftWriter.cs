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

        // A model occasionally omits the separator in an otherwise
        // unambiguous existing-sheet reference (Ledger$B$2). Repair only
        // that narrow form; never guess at a bare name or at quoted text.
        internal static string NormalizeSheetReferences(
            string formula,
            IEnumerable<string> worksheetNames)
        {
            if (string.IsNullOrEmpty(formula) || worksheetNames == null)
                return formula;
            var names = worksheetNames.Where(name =>
                    !string.IsNullOrWhiteSpace(name) &&
                    Regex.IsMatch(name, @"^[A-Za-z_][A-Za-z0-9_]*$"))
                .OrderByDescending(name => name.Length).ToArray();
            var parts = formula.Split('"');
            for (var part = 0; part < parts.Length; part += 2)
                foreach (var name in names)
                    parts[part] = Regex.Replace(
                        parts[part],
                        @"(?<![A-Za-z0-9_.!'])" + Regex.Escape(name) +
                        @"(?=\$[A-Z]{1,3}\$?\d+\b)",
                        match => match.Value + "!",
                        RegexOptions.IgnoreCase);
            return string.Join("\"", parts);
        }

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
            // set one by one below. A syntax rejection stays visible as
            // text; calculated errors retain their native formula receipts.
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
                                    NormalizeSheetReferences(cell, existingNames)));
                            continue;
                        }

                        // The apostrophe keeps blocked formula
                        // text visible as plain text.
                        cell = "'" + cell;
                    }

                    // The model supplies text month keys, not Excel date serials.
                    // Preserve the same key type as source strings such as 2026-06.
                    if (Regex.IsMatch(cell, @"^\d{4}-(?:0[1-9]|1[0-2])$")) cell = "'" + cell;
                    grid[row, column] = cell;
                }
            }

            dynamic target = sheet.Range(
                sheet.Cells[startRow, 1],
                sheet.Cells[
                    startRow + rowCount - 1,
                    columnCount]);
            // Headers such as 2026-05 are labels. Excel otherwise coerces
            // them into date serials (46143), losing the supplied meaning.
            target.Rows[1].NumberFormat = "@";
            target.Value2 = grid;
            var formulaCount = 0;
            var rejectedFormulas = new List<string>();
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
                        rejectedFormulas.Add("R" + (startRow + formula.Key[0]) +
                            "C" + (formula.Key[1] + 1));
                    }
                }
            }

            if (rejectedFormulas.Count > 0)
                throw new InvalidOperationException(
                    "DRAFT_FORMULA_INVALID: Excel rejected " +
                    rejectedFormulas.Count + " formula(s) at " +
                    string.Join(", ", rejectedFormulas.Take(8)) +
                    ". They remain visible as text on the new draft sheet, but this is not a valid analytical output. Correct the syntax and create a fresh draft before continuing to PowerPoint or claiming completion.");

            // A formula that parses but evaluates to an Excel error is
            // definitely wrong: repair a one-row header offset when safe,
            // otherwise retain the visible error and formula for correction.
            // Turning it into text hides failures from native readback.
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
                columnCount,
                target,
                rows);
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
                      " to an Excel error. The formulas remain visible for correction; " +
                      "do not claim the analysis is complete. Check source types and sheet references."
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
            var existingNames = new List<string>();
            foreach (dynamic candidate in sheet.Parent.Worksheets)
                existingNames.Add(Convert.ToString(candidate.Name) ?? string.Empty);

            dynamic anchor = sheet.Range(anchorName);
            int startRow = anchor.Row;
            int startColumn = anchor.Column;
            var rowCount = Math.Min(rows.Count, MaxDraftRows);
            var written = 0;
            var formulaCount = 0;
            var brokenFormulas = 0;
            var rejectedFormulas = new List<string>();
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
                        cell = NormalizeSheetReferences(cell, existingNames);

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
                                rejectedFormulas.Add("R" + (startRow + row) +
                                    "C" + (startColumn + column));
                            }
                        }

                        written++;
                        continue;
                    }

                    target.Value2 = cell;
                    written++;
                }
            }

            if (rejectedFormulas.Count > 0)
                throw new InvalidOperationException(
                    "DRAFT_FORMULA_INVALID: Excel rejected " +
                    rejectedFormulas.Count + " formula(s) at " +
                    string.Join(", ", rejectedFormulas.Take(8)) +
                    ". The affected cells remain visible as text, but this is not a valid analytical output. Correct the syntax before continuing or claiming completion.");

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
                      "and remain visible for correction. Do not claim the analysis is complete."
                    : string.Empty) +
                " Nothing was saved, but Excel cannot undo " +
                "add-in changes - close without saving to " +
                "discard.";
        }

        private static bool IsExcelError(object value)
        {
            return ExcelErrorValue.Text(value) != null;
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

        // Cosmetic polish for the draft sheet. Formatting is deliberately
        // applied after values and formulas, so it can improve readability
        // without changing the model's data, formula text, or source links.
        // Any failure here must never fail the draft itself.
        private static void ApplyDraftFormatting(
            dynamic sheet,
            string boundedTitle,
            int startRow,
            int rowCount,
            int columnCount,
            dynamic target,
            IReadOnlyList<IReadOnlyList<string>> rows)
        {
            try
            {
                if (boundedTitle.Length > 0)
                {
                    dynamic titleCell = sheet.Cells[1, 1];
                    titleCell.Font.Bold = true;
                    titleCell.Font.Size = 16;
                    titleCell.Font.Name = "Aptos Display";
                    // RGB(31, 78, 121), Excel's OLE/BGR integer.
                    titleCell.Font.Color = 0x794E1F;
                }

                dynamic header = target.Rows[1];
                header.Font.Bold = true;
                header.Font.Color = 0xFFFFFF;
                header.Font.Name = "Aptos";
                header.Interior.Color = 0x794E1F;
                header.HorizontalAlignment = -4108; // xlCenter.
                header.VerticalAlignment = -4108;
                header.RowHeight = 22;
                // 9 = xlEdgeBottom, 1 = xlContinuous.
                header.Borders[9].LineStyle = 1;
                header.Borders[9].Color = 0x794E1F;

                target.EntireColumn.AutoFit();
            }
            catch
            {
            }

            if (rowCount > 1)
            {
                try
                {
                    dynamic body = sheet.Range(
                        sheet.Cells[startRow + 1, 1],
                        sheet.Cells[startRow + rowCount - 1, columnCount]);
                    body.Font.Name = "Aptos";
                    // Add separators while preserving up to five displayed
                    // decimals. Formula precision remains unchanged.
                    body.NumberFormat = "#,##0.#####";
                    body.Borders.LineStyle = 1; // xlContinuous.
                    body.Borders.Color = 0xD9D9D9;
                    body.VerticalAlignment = -4108;
                    for (var offset = 1; offset < rowCount; offset += 2)
                    {
                        dynamic band = target.Rows[offset + 1];
                        band.Interior.Color = 0xF7EBDD;
                    }
                    if (columnCount > 1)
                        for (var offset = 1; offset < rowCount; offset++)
                        {
                            var label = rows != null && offset < rows.Count && rows[offset] != null && rows[offset].Count > 0
                                ? rows[offset][0] ?? "" : "";
                            if (!Regex.IsMatch(label, @"(?i)(?:\bmargin\b|\brate\b|\bpercent(?:age)?\b|\bshare\b|%)"))
                                continue;
                            dynamic percentageRow = sheet.Range(
                                sheet.Cells[startRow + offset, 2],
                                sheet.Cells[startRow + offset, columnCount]);
                            percentageRow.NumberFormat = "0.00%";
                        }
                }
                catch
                {
                }

                try
                {
                    target.AutoFilter();
                }
                catch
                {
                }
            }

            try
            {
                // AutoFit can make a prose/disclosure column hundreds of
                // characters wide. Refit after applying the displayed number
                // format so numeric columns reserve room for separators and
                // decimals, then cap prose columns and wrap only when needed.
                target.EntireColumn.AutoFit();
                for (var column = 1; column <= columnCount; column++)
                {
                    dynamic draftColumn = sheet.Columns[column];
                    var width = Convert.ToDouble(draftColumn.ColumnWidth);
                    if (width > 36d)
                    {
                        draftColumn.ColumnWidth = 36d;
                        draftColumn.WrapText = true;
                        width = 36d;
                    }
                    // Excel can AutoFit a formula column against its short
                    // header before the recalculated value is displayed.
                    // Reserve enough space for separators and decimals in
                    // numeric output columns, including formula results.
                    var numeric = column > 1 && rows != null &&
                        rows.Skip(1).Any(row => row != null &&
                            row.Count >= column &&
                            IsChartValue(row[column - 1]));
                    if (numeric && width < 14d)
                        draftColumn.ColumnWidth = 14d;
                }
                target.EntireRow.AutoFit();
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
