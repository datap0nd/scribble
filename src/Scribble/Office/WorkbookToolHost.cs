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
        public const int MaxReadRows = 500;
        public const int MaxReadColumns = 50;
        public const int MaxCellCharacters = 500;
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

            bool truncated;
            var text = ReadRangeText(range, out truncated);
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
                    { "truncated", truncated },
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
