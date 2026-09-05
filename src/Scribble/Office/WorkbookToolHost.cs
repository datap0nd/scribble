using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Web.Script.Serialization;
using Scribble.Chat;
using Scribble.Security;
using Scribble.Utilities;

namespace Scribble.Office
{
    // Read-only Excel access for the Scribble pane. Every read is
    // bounded before it reaches the model and cell text always
    // travels as untrusted data. This host holds no write
    // capability: draft writes live in WorkbookDraftHost behind the
    // one-shot authorization.
    public sealed class WorkbookToolHost
    {
        // Per-call transport windows. They never limit the selected
        // range: read_cells returns the next offsets until complete.
        public const int MaxReadRows = 500;
        public const int MaxReadColumns = 50;
        public const int MaxCellCharacters = 32767;
        public const int MaxSheets = 100;

        private readonly object _excelApplication;
        private readonly JavaScriptSerializer _serializer =
            new JavaScriptSerializer();

        public WorkbookToolHost(object excelApplication)
        {
            _excelApplication = excelApplication ??
                throw new ArgumentNullException(
                    nameof(excelApplication));
        }

        public MailboxToolResult Execute(ChatToolCall call)
        {
            if (call?.function == null ||
                string.IsNullOrWhiteSpace(call.id))
            {
                return Error(
                    call?.id,
                    "WORKBOOK_TOOL_CALL_INVALID",
                    "The model returned an invalid tool call.");
            }

            var name = call.function.name ?? string.Empty;
            if (!WorkbookToolCatalog.IsApproved(name))
            {
                return Error(
                    call.id,
                    "WORKBOOK_TOOL_NOT_ALLOWED",
                    "The requested workbook tool is not allowed.");
            }

            try
            {
                var arguments = ToolArguments.Parse(
                    _serializer,
                    call.function.arguments);
                switch (name)
                {
                    case WorkbookToolCatalog.ListWorksheets:
                        return ListWorksheets(call.id);
                    case WorkbookToolCatalog.ReadCells:
                        return ReadCells(call.id, arguments);
                    default:
                        return Error(
                            call.id,
                            "WORKBOOK_TOOL_NOT_ALLOWED",
                            "The requested workbook tool is not allowed.");
                }
            }
            catch (Exception exception)
            {
                Log.Error("WorkbookTool." + name, exception);
                return Error(
                    call.id,
                    "WORKBOOK_TOOL_FAILED",
                    DiagnosticDetails.ForException(
                        exception,
                        "WORKBOOK_TOOL_FAILED"));
            }
        }

        // Bounded description of the open workbook for the request
        // context: workbook name, sheet inventory, and the current
        // selection address.
        public string DescribeActiveContext()
        {
            try
            {
                dynamic application = _excelApplication;
                dynamic workbook = application.ActiveWorkbook;
                if (workbook == null)
                {
                    return "No workbook is open in Excel.";
                }

                var lines = new List<string>
                {
                    "Workbook: " + TextBoundary.SingleLine(
                        Convert.ToString(workbook.Name),
                        180)
                };
                var saved = string.Empty;
                try
                {
                    saved = Convert.ToString(workbook.Path) ??
                        string.Empty;
                }
                catch
                {
                }

                lines.Add(saved.Length > 0
                    ? "Saved on disk: yes"
                    : "Saved on disk: no (unsaved workbook)");
                var count = 0;
                foreach (dynamic sheet in workbook.Worksheets)
                {
                    if (count == MaxSheets)
                    {
                        lines.Add("(more sheets not listed)");
                        break;
                    }

                    string used;
                    try
                    {
                        used = Convert.ToString(
                            sheet.UsedRange.Address(false, false));
                    }
                    catch
                    {
                        used = "empty";
                    }

                    lines.Add(
                        "Sheet: " +
                        TextBoundary.SingleLine(
                            Convert.ToString(sheet.Name),
                            120) +
                        " (used range " +
                        TextBoundary.SingleLine(used, 60) +
                        ")");
                    count++;
                }

                try
                {
                    dynamic active = workbook.ActiveSheet;
                    lines.Add(
                        "Active sheet: " +
                        TextBoundary.SingleLine(
                            Convert.ToString(active.Name),
                            120));
                }
                catch
                {
                }

                try
                {
                    dynamic selection = application.Selection;
                    lines.Add(
                        "Current selection: " +
                        TextBoundary.SingleLine(
                            Convert.ToString(
                                selection.Address(false, false)),
                            80));
                }
                catch
                {
                }

                return string.Join("\n", lines);
            }
            catch (Exception exception)
            {
                Log.Error("WorkbookDescribe", exception);
                return "The workbook context could not be read.";
            }
        }

        // Bounded snapshot of the current selection for the context
        // tray ("add current selection").
        public string DescribeSelection(out string title)
        {
            var snapshot = CaptureSelection();
            title = TextBoundary.SingleLine(
                "Excel " + snapshot.WorksheetName + "!" +
                snapshot.Address,
                120);
            return snapshot.BuildContextText(string.Empty);
        }

        // Captures identity and a bounded preview before a Ribbon
        // callback creates or focuses the task pane. Whole-column
        // and whole-row selections are reduced to the used range;
        // discontiguous selections fail closed because a single
        // output column cannot preserve their row alignment.
        public ExcelSelectionSnapshot CaptureSelection()
        {
            dynamic application = _excelApplication;
            dynamic selection = application.Selection;
            if (selection == null)
            {
                throw new InvalidOperationException(
                    "Select cells in Excel first.");
            }

            try
            {
                if ((int)selection.Areas.Count != 1)
                {
                    throw new InvalidOperationException(
                        "Scribble cannot send a multi-area Excel selection. " +
                        "Filtered selections can create separate areas; " +
                        "select one contiguous block and try again.");
                }
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch
            {
                throw new InvalidOperationException(
                    "Select one contiguous range of Excel cells.");
            }

            dynamic sheet;
            try
            {
                sheet = selection.Worksheet;
            }
            catch
            {
                throw new InvalidOperationException(
                    "The Excel selection is not a cell range.");
            }

            dynamic normalized = selection;
            var selectedRows = (int)selection.Rows.Count;
            var selectedColumns = (int)selection.Columns.Count;
            if (selectedRows >= 1048576 ||
                selectedColumns >=
                    ExcelSelectionOutputPolicy.MaxExcelColumns)
            {
                try
                {
                    normalized = application.Intersect(
                        selection,
                        sheet.UsedRange);
                }
                catch
                {
                    normalized = null;
                }

                if (normalized == null)
                {
                    throw new InvalidOperationException(
                        "The selected rows or columns do not contain used cells.");
                }
            }

            var rowCount = (int)normalized.Rows.Count;
            var columnCount = (int)normalized.Columns.Count;
            if (rowCount < 1 || columnCount < 1)
            {
                throw new InvalidOperationException(
                    "The selected Excel range is empty.");
            }

            dynamic workbook = sheet.Parent;
            var workbookName = TextBoundary.SingleLine(
                Convert.ToString(workbook.Name),
                180);
            var workbookPath = string.Empty;
            var workbookFullName = string.Empty;
            try
            {
                workbookPath = Convert.ToString(workbook.Path) ??
                    string.Empty;
                workbookFullName = Convert.ToString(
                    workbook.FullName) ?? string.Empty;
            }
            catch
            {
            }

            var windowHandle = 0;
            try
            {
                windowHandle = (int)application.ActiveWindow.Hwnd;
            }
            catch
            {
            }

            var address = TextBoundary.SingleLine(
                Convert.ToString(
                    normalized.Address(false, false, 1)),
                80);
            var sheetName = TextBoundary.SingleLine(
                Convert.ToString(sheet.Name),
                120);
            bool truncated;
            var text = ReadRangeText(
                normalized,
                out truncated);
            if (text.Trim().Length == 0)
            {
                throw new InvalidOperationException(
                    "The selected Excel range does not contain any values.");
            }

            return new ExcelSelectionSnapshot(
                Guid.NewGuid().ToString("N"),
                workbookPath.Length > 0,
                workbookPath.Length > 0
                    ? workbookFullName
                    : workbookName,
                workbookName,
                windowHandle,
                sheetName,
                address,
                (int)normalized.Row,
                (int)normalized.Column,
                rowCount,
                columnCount,
                text,
                truncated);
        }

        // Deterministically discovers literal Korean text throughout
        // the active workbook for the built-in translation skill.
        // Values are bulk-read in transport tiles; only sparse matching
        // cells are retained and no workbook content is changed here.
        public KoreanWorkbookSnapshot CaptureKoreanWorkbook()
        {
            dynamic application = _excelApplication;
            dynamic workbook = application.ActiveWorkbook;
            if (workbook == null)
            {
                throw new InvalidOperationException(
                    "Open an Excel workbook first.");
            }

            var workbookName = TextBoundary.SingleLine(
                Convert.ToString(workbook.Name),
                180);
            var workbookPath = string.Empty;
            var workbookFullName = string.Empty;
            try
            {
                workbookPath = Convert.ToString(workbook.Path) ??
                    string.Empty;
                workbookFullName = Convert.ToString(
                    workbook.FullName) ?? string.Empty;
            }
            catch
            {
            }

            var windowHandle = 0;
            try
            {
                windowHandle = (int)application.ActiveWindow.Hwnd;
            }
            catch
            {
            }

            var cells = new List<KoreanWorkbookCellSnapshot>();
            var skippedFormulaCells = 0;
            var skippedMergedCells = 0;
            foreach (dynamic sheet in workbook.Worksheets)
            {
                dynamic used = sheet.UsedRange;
                var totalRows = (int)used.Rows.Count;
                var totalColumns = (int)used.Columns.Count;
                var firstRow = (int)used.Row;
                var firstColumn = (int)used.Column;
                var sheetName = TextBoundary.SingleLine(
                    Convert.ToString(sheet.Name),
                    120);
                for (var rowOffset = 0;
                     rowOffset < totalRows;
                     rowOffset += MaxReadRows)
                {
                    var rows = Math.Min(
                        MaxReadRows,
                        totalRows - rowOffset);
                    for (var columnOffset = 0;
                         columnOffset < totalColumns;
                         columnOffset += MaxReadColumns)
                    {
                        var columns = Math.Min(
                            MaxReadColumns,
                            totalColumns - columnOffset);
                        dynamic tile = sheet.Range(
                            sheet.Cells[
                                firstRow + rowOffset,
                                firstColumn + columnOffset],
                            sheet.Cells[
                                firstRow + rowOffset + rows - 1,
                                firstColumn + columnOffset + columns - 1]);
                        object raw = tile.Value2;
                        var grid = raw as object[,];
                        if (grid == null)
                        {
                            AddKoreanWorkbookCell(
                                sheet,
                                sheetName,
                                firstRow + rowOffset,
                                firstColumn + columnOffset,
                                raw,
                                cells,
                                ref skippedFormulaCells,
                                ref skippedMergedCells);
                            continue;
                        }

                        var rowBase = grid.GetLowerBound(0);
                        var columnBase = grid.GetLowerBound(1);
                        for (var row = 0; row < rows; row++)
                        {
                            for (var column = 0;
                                 column < columns;
                                 column++)
                            {
                                AddKoreanWorkbookCell(
                                    sheet,
                                    sheetName,
                                    firstRow + rowOffset + row,
                                    firstColumn + columnOffset + column,
                                    grid[
                                        rowBase + row,
                                        columnBase + column],
                                    cells,
                                    ref skippedFormulaCells,
                                    ref skippedMergedCells);
                            }
                        }
                    }
                }
            }

            return new KoreanWorkbookSnapshot(
                workbookPath.Length > 0,
                workbookPath.Length > 0
                    ? workbookFullName
                    : workbookName,
                workbookName,
                windowHandle,
                cells,
                skippedFormulaCells,
                skippedMergedCells);
        }

        private static void AddKoreanWorkbookCell(
            dynamic sheet,
            string sheetName,
            int row,
            int column,
            object rawValue,
            ICollection<KoreanWorkbookCellSnapshot> cells,
            ref int skippedFormulaCells,
            ref int skippedMergedCells)
        {
            var text = CellText(rawValue);
            if (!ExcelSelectionOutputPolicy.ContainsKorean(text))
            {
                return;
            }

            dynamic cell = sheet.Cells[row, column];
            try
            {
                object hasFormula = cell.HasFormula;
                if (!(hasFormula is bool) || (bool)hasFormula)
                {
                    skippedFormulaCells++;
                    return;
                }

                object merged = cell.MergeCells;
                if (!(merged is bool) || (bool)merged)
                {
                    skippedMergedCells++;
                    return;
                }
            }
            catch
            {
                // Ambiguous cell state must never broaden an
                // automatic workbook-wide overwrite.
                skippedFormulaCells++;
                return;
            }

            cells.Add(new KoreanWorkbookCellSnapshot(
                sheetName,
                ExcelSelectionOutputPolicy.ColumnNumberToName(column) +
                    row.ToString(CultureInfo.InvariantCulture),
                text));
        }

        private MailboxToolResult ListWorksheets(string callId)
        {
            dynamic application = _excelApplication;
            dynamic workbook = application.ActiveWorkbook;
            if (workbook == null)
            {
                return Error(
                    callId,
                    "WORKBOOK_NOT_OPEN",
                    "No workbook is open in Excel.");
            }

            var sheets = new List<object>();
            var count = 0;
            foreach (dynamic sheet in workbook.Worksheets)
            {
                if (count == MaxSheets)
                {
                    break;
                }

                var entry = new Dictionary<string, object>
                {
                    {
                        "name",
                        TextBoundary.SingleLine(
                            Convert.ToString(sheet.Name),
                            120)
                    }
                };
                try
                {
                    dynamic used = sheet.UsedRange;
                    entry["used_range"] = TextBoundary.SingleLine(
                        Convert.ToString(
                            used.Address(false, false)),
                        60);
                    entry["rows"] = (int)used.Rows.Count;
                    entry["columns"] = (int)used.Columns.Count;
                }
                catch
                {
                    entry["used_range"] = "empty";
                }

                sheets.Add(entry);
                count++;
            }

            return Success(
                callId,
                new Dictionary<string, object>
                {
                    { "untrusted_document_data", true },
                    {
                        "workbook",
                        TextBoundary.SingleLine(
                            Convert.ToString(workbook.Name),
                            180)
                    },
                    { "sheet_count", sheets.Count },
                    { "sheets", sheets }
                },
                "Listed " +
                sheets.Count.ToString(CultureInfo.InvariantCulture) +
                " worksheets.");
        }

        private MailboxToolResult ReadCells(
            string callId,
            IDictionary<string, object> arguments)
        {
            dynamic application = _excelApplication;
            dynamic workbook = application.ActiveWorkbook;
            if (workbook == null)
            {
                return Error(
                    callId,
                    "WORKBOOK_NOT_OPEN",
                    "No workbook is open in Excel.");
            }

            var sheetName = ToolArguments.GetString(
                arguments,
                "sheet",
                string.Empty);
            var rangeText = TextBoundary.SingleLine(
                ToolArguments.GetString(
                    arguments,
                    "range",
                    string.Empty),
                60);

            dynamic sheet = null;
            if (sheetName.Length > 0)
            {
                foreach (dynamic candidate in workbook.Worksheets)
                {
                    if (string.Equals(
                        Convert.ToString(candidate.Name),
                        sheetName,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        sheet = candidate;
                        break;
                    }
                }

                if (sheet == null)
                {
                    return Error(
                        callId,
                        "WORKBOOK_SHEET_UNKNOWN",
                        "No worksheet with that name exists. Call list_worksheets first.");
                }
            }
            else
            {
                sheet = workbook.ActiveSheet;
            }

            dynamic range;
            if (rangeText.Length > 0)
            {
                try
                {
                    range = sheet.Range(rangeText);
                }
                catch
                {
                    return Error(
                        callId,
                        "WORKBOOK_RANGE_INVALID",
                        "The range must be A1-style, such as A1:F40.");
                }
            }
            else
            {
                range = sheet.UsedRange;
            }

            var totalRows = (int)range.Rows.Count;
            var totalColumns = (int)range.Columns.Count;
            var rowOffset = ToolArguments.GetInteger(
                arguments,
                "row_offset",
                0,
                0,
                Math.Max(0, totalRows - 1));
            var columnOffset = ToolArguments.GetInteger(
                arguments,
                "column_offset",
                0,
                0,
                Math.Max(0, totalColumns - 1));
            var requestedRows = ToolArguments.GetInteger(
                arguments,
                "max_rows",
                MaxReadRows,
                1,
                MaxReadRows);
            var requestedColumns = ToolArguments.GetInteger(
                arguments,
                "max_columns",
                MaxReadColumns,
                1,
                MaxReadColumns);
            var rows = Math.Min(requestedRows, totalRows - rowOffset);
            var columns = Math.Min(
                requestedColumns,
                totalColumns - columnOffset);
            // When columns require more than one page, keep each
            // tile to one row. That makes (row, column) a complete
            // cursor and prevents a later column page from changing
            // row height and skipping cells.
            if (columnOffset > 0 || requestedColumns < totalColumns)
            {
                rows = 1;
            }
            dynamic page = range.Worksheet.Range(
                range.Cells[rowOffset + 1, columnOffset + 1],
                range.Cells[rowOffset + rows, columnOffset + columns]);
            bool pageTruncated;
            var text = ReadRangeText(page, out pageTruncated);
            if (pageTruncated)
            {
                return Error(
                    callId,
                    "WORKBOOK_PAGE_TOO_LARGE",
                    "This transport page is too large. Retry the same " +
                    "row_offset and column_offset with smaller max_rows " +
                    "or max_columns; the selected range remains available.");
            }

            var nextRowOffset = rowOffset;
            var nextColumnOffset = columnOffset + columns;
            if (nextColumnOffset >= totalColumns)
            {
                nextColumnOffset = 0;
                nextRowOffset = rowOffset + rows;
            }
            var complete = nextRowOffset >= totalRows;
            return Success(
                callId,
                new Dictionary<string, object>
                {
                    { "untrusted_document_data", true },
                    {
                        "sheet",
                        TextBoundary.SingleLine(
                            Convert.ToString(sheet.Name),
                            120)
                    },
                    {
                        "range",
                        TextBoundary.SingleLine(
                            Convert.ToString(
                                range.Address(false, false)),
                            60)
                    },
                    { "total_rows", totalRows },
                    { "total_columns", totalColumns },
                    { "row_offset", rowOffset },
                    { "column_offset", columnOffset },
                    { "returned_rows", rows },
                    { "returned_columns", columns },
                    { "complete", complete },
                    { "next_row_offset", complete ? 0 : nextRowOffset },
                    { "next_column_offset", complete ? 0 : nextColumnOffset },
                    { "cells_tsv", text }
                },
                "Read cells from " +
                TextBoundary.SingleLine(
                    Convert.ToString(sheet.Name),
                    120) +
                ".");
        }

        // Bulk-reads range.Value2 and renders a bounded TSV block.
        private static string ReadRangeText(
            dynamic range,
            out bool truncated)
        {
            truncated = false;
            var totalRows = (int)range.Rows.Count;
            var totalColumns = (int)range.Columns.Count;
            var rows = Math.Min(totalRows, MaxReadRows);
            var columns = Math.Min(totalColumns, MaxReadColumns);
            if (rows < totalRows || columns < totalColumns)
            {
                truncated = true;
                range = range.Resize(rows, columns);
            }

            object value = range.Value2;
            var builder = new StringBuilder();
            var grid = value as object[,];
            if (grid == null)
            {
                builder.Append(CellText(value));
            }
            else
            {
                var rowBase = grid.GetLowerBound(0);
                var columnBase = grid.GetLowerBound(1);
                for (var row = 0; row < rows; row++)
                {
                    if (row > 0)
                    {
                        builder.Append('\n');
                    }

                    for (var column = 0;
                         column < columns;
                         column++)
                    {
                        if (column > 0)
                        {
                            builder.Append('\t');
                        }

                        builder.Append(
                            CellText(
                                grid[
                                    rowBase + row,
                                    columnBase + column]));
                    }
                }
            }

            var text = builder.ToString();
            var cap = ContextScale.Scaled(
                TextBoundary.MaxToolResultCharacters) / 2;
            if (text.Length > cap)
            {
                truncated = true;
                text = text.Substring(0, cap);
            }

            return text;
        }

        private static string CellText(object value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            string text;
            if (value is double)
            {
                text = ((double)value).ToString(
                    "0.############",
                    CultureInfo.InvariantCulture);
            }
            else
            {
                text = Convert.ToString(
                    value,
                    CultureInfo.InvariantCulture) ?? string.Empty;
            }

            return TextBoundary.SingleLine(
                text,
                MaxCellCharacters);
        }

        private MailboxToolResult Success(
            string callId,
            object payload,
            string status)
        {
            var json = _serializer.Serialize(payload);
            if (json.Length >
                ContextScale.Scaled(
                    TextBoundary.MaxToolResultCharacters))
            {
                return Error(
                    callId,
                    "WORKBOOK_TOOL_RESULT_TOO_LARGE",
                    "The bounded workbook result was still too large to return safely.");
            }

            return new MailboxToolResult(callId, json, status);
        }

        private MailboxToolResult Error(
            string callId,
            string code,
            string message)
        {
            return new MailboxToolResult(
                callId ?? string.Empty,
                _serializer.Serialize(
                    new Dictionary<string, object>
                    {
                        { "ok", false },
                        { "error_code", code },
                        {
                            "message",
                            TextBoundary.PlainText(message, 1200)
                        }
                    }),
                "[" + code + "] " + message);
        }
    }

    // Shared argument parsing for the document tool hosts, matching
    // the mailbox host's bounded conversions.
    internal static class ToolArguments
    {
        internal static IDictionary<string, object> Parse(
            JavaScriptSerializer serializer,
            string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return new Dictionary<string, object>();
            }

            try
            {
                var parsed = serializer.DeserializeObject(json) as
                    IDictionary<string, object>;
                if (parsed == null)
                {
                    throw new InvalidOperationException(
                        "Tool arguments must be a JSON object.");
                }

                return parsed;
            }
            catch (Exception exception)
            {
                throw new AiEndpointException(
                    "TOOL_ARGUMENTS_INVALID_JSON",
                    "The model returned invalid JSON tool arguments.",
                    exception,
                    responseSnippet: json);
            }
        }

        internal static string GetString(
            IDictionary<string, object> arguments,
            string key,
            string fallback)
        {
            object value;
            return arguments.TryGetValue(key, out value)
                ? TextBoundary.PlainText(
                    Convert.ToString(value),
                    1000)
                : fallback;
        }

        internal static int GetInteger(
            IDictionary<string, object> arguments,
            string key,
            int fallback,
            int minimum,
            int maximum)
        {
            object value;
            int parsed;
            if (!arguments.TryGetValue(key, out value) ||
                !int.TryParse(
                    Convert.ToString(value),
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out parsed))
            {
                return fallback;
            }

            return Math.Max(minimum, Math.Min(maximum, parsed));
        }

        internal static bool GetBoolean(
            IDictionary<string, object> arguments,
            string key)
        {
            object value;
            if (!arguments.TryGetValue(key, out value))
            {
                return false;
            }

            if (value is bool)
            {
                return (bool)value;
            }

            bool parsed;
            return bool.TryParse(
                Convert.ToString(value),
                out parsed) && parsed;
        }
    }
}
