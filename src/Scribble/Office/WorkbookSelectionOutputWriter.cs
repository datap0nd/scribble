using System;
using System.Collections.Generic;
using Scribble.Security;

namespace Scribble.Office
{
    internal sealed class ExcelSelectionDestinationException :
        InvalidOperationException
    {
        internal ExcelSelectionDestinationException(
            string code,
            string message,
            IReadOnlyList<string> candidates)
            : base(message)
        {
            Code = code ?? "SELECTION_DESTINATION_INVALID";
            Candidates = candidates ?? new string[0];
        }

        internal string Code { get; }

        internal IReadOnlyList<string> Candidates { get; }
    }

    // The only COM-writing boundary for staged selection output.
    // Validation is repeated immediately before the one bulk write;
    // no source cell is changed and Scribble never saves the file.
    internal static class WorkbookSelectionOutputWriter
    {
        internal static void ValidateDestination(
            object excelApplication,
            ExcelSelectionSnapshot snapshot,
            string destinationColumn)
        {
            dynamic sheet = ResolveSheet(
                excelApplication,
                snapshot);
            var destinationNumber =
                ExcelSelectionOutputPolicy.ColumnNameToNumber(
                    destinationColumn);
            if (destinationNumber == 0 ||
                destinationNumber == snapshot.StartColumn)
            {
                throw DestinationError(
                    (object)sheet,
                    snapshot,
                    "SELECTION_DESTINATION_INVALID",
                    "Choose a valid destination column outside the source.");
            }

            dynamic destination = DestinationRange(
                sheet,
                snapshot,
                destinationNumber);
            if (!IsWritable(destination, snapshot.RowCount))
            {
                throw DestinationError(
                    (object)sheet,
                    snapshot,
                    "SELECTION_DESTINATION_OCCUPIED",
                    "The destination contains data, formulas, or merged " +
                    "cells. Choose one of the empty columns returned.");
            }
        }

        internal static string Commit(
            object excelApplication,
            ExcelSelectionSnapshot snapshot,
            string destinationColumn,
            IReadOnlyList<string> values)
        {
            if (values == null || values.Count != snapshot.RowCount)
            {
                throw new InvalidOperationException(
                    "The completed output must contain exactly one " +
                    "value per source row.");
            }

            dynamic application = excelApplication;
            dynamic sheet = application.ActiveSheet;
            var destinationNumber =
                ExcelSelectionOutputPolicy.ColumnNameToNumber(
                    destinationColumn);
            dynamic destination = DestinationRange(
                sheet,
                snapshot,
                destinationNumber);
            var grid = new object[snapshot.RowCount, 1];
            for (var row = 0; row < snapshot.RowCount; row++)
            {
                grid[row, 0] = values[row];
            }

            // Text format plus apostrophe sanitization in the
            // COM-free policy keeps =, +, -, and @ values inert.
            destination.NumberFormat = "@";
            destination.Value2 = grid;
            return "Wrote " + snapshot.RowCount +
                " literal values to " + snapshot.WorksheetName +
                "!" + destinationColumn + snapshot.StartRow + ":" +
                destinationColumn +
                (snapshot.StartRow + snapshot.RowCount - 1) +
                ". Source cells were unchanged. Nothing was saved.";
        }

        private static dynamic ResolveSheet(
            object excelApplication,
            ExcelSelectionSnapshot snapshot)
        {
            if (snapshot == null)
            {
                throw new InvalidOperationException(
                    "The selection handle is unknown or expired.");
            }

            dynamic application = excelApplication;
            dynamic workbook = application.ActiveWorkbook;
            dynamic sheet = application.ActiveSheet;
            if (workbook == null || sheet == null)
            {
                throw new InvalidOperationException(
                    "The captured workbook and worksheet are no longer active.");
            }

            var path = string.Empty;
            var fullName = string.Empty;
            try
            {
                path = Convert.ToString(workbook.Path) ?? string.Empty;
                fullName = Convert.ToString(workbook.FullName) ??
                    string.Empty;
            }
            catch
            {
            }

            var workbookName = Convert.ToString(workbook.Name) ??
                string.Empty;
            var sheetName = Convert.ToString(sheet.Name) ??
                string.Empty;
            var windowHandle = 0;
            try
            {
                windowHandle = (int)application.ActiveWindow.Hwnd;
            }
            catch
            {
            }

            var saved = path.Length > 0;
            if (!ExcelSelectionOutputPolicy.IdentityMatches(
                snapshot,
                saved,
                saved ? fullName : workbookName,
                workbookName,
                windowHandle,
                sheetName))
            {
                throw new InvalidOperationException(
                    "The workbook, window, or worksheet changed after " +
                    "the selection was sent to Scribble.");
            }

            dynamic source;
            try
            {
                source = sheet.Range(snapshot.Address);
            }
            catch
            {
                throw new InvalidOperationException(
                    "The captured source range no longer resolves.");
            }

            var resolvedAddress = TextBoundary.SingleLine(
                Convert.ToString(source.Address(false, false, 1)),
                80);
            if (!string.Equals(
                    resolvedAddress,
                    snapshot.Address,
                    StringComparison.OrdinalIgnoreCase) ||
                (int)source.Row != snapshot.StartRow ||
                (int)source.Column != snapshot.StartColumn ||
                (int)source.Rows.Count != snapshot.RowCount ||
                (int)source.Columns.Count != snapshot.ColumnCount)
            {
                throw new InvalidOperationException(
                    "The captured source range changed before the write.");
            }

            return sheet;
        }

        private static dynamic DestinationRange(
            dynamic sheet,
            ExcelSelectionSnapshot snapshot,
            int destinationColumn)
        {
            dynamic first = sheet.Cells[
                snapshot.StartRow,
                destinationColumn];
            dynamic last = sheet.Cells[
                snapshot.StartRow + snapshot.RowCount - 1,
                destinationColumn];
            return sheet.Range(first, last);
        }

        private static bool IsWritable(
            dynamic range,
            int rowCount)
        {
            var states = new List<ExcelDestinationCellState>();
            var hasFormula = true;
            var merged = true;
            try
            {
                // Excel returns null for a mixed state. Only an
                // unambiguous false is accepted for the full range.
                hasFormula = !IsFalse(range.HasFormula);
                merged = !IsFalse(range.MergeCells);
            }
            catch
            {
                // Ambiguous COM state fails closed.
            }

            object raw = range.Value2;
            var grid = raw as object[,];
            if (grid == null)
            {
                states.Add(new ExcelDestinationCellState(
                    Convert.ToString(raw) ?? string.Empty,
                    hasFormula,
                    merged));
            }
            else
            {
                var rowBase = grid.GetLowerBound(0);
                var columnBase = grid.GetLowerBound(1);
                for (var row = 0; row < rowCount; row++)
                {
                    states.Add(new ExcelDestinationCellState(
                        Convert.ToString(
                            grid[rowBase + row, columnBase]) ??
                            string.Empty,
                        hasFormula,
                        merged));
                }
            }

            return ExcelSelectionOutputPolicy.IsDestinationWritable(
                states);
        }

        private static bool IsFalse(object value)
        {
            return value is bool && !(bool)value;
        }

        private static ExcelSelectionDestinationException
            DestinationError(
                object sheetValue,
                ExcelSelectionSnapshot snapshot,
                string code,
                string message)
        {
            dynamic sheet = sheetValue;
            return new ExcelSelectionDestinationException(
                code,
                message,
                FindCandidates(sheet, snapshot));
        }

        private static IReadOnlyList<string> FindCandidates(
            dynamic sheet,
            ExcelSelectionSnapshot snapshot)
        {
            var candidates = new List<string>();
            for (var distance = 1;
                 distance <= 20 && candidates.Count < 3;
                 distance++)
            {
                AddCandidate(
                    sheet,
                    snapshot,
                    snapshot.StartColumn + distance,
                    candidates);
                if (candidates.Count < 3)
                {
                    AddCandidate(
                        sheet,
                        snapshot,
                        snapshot.StartColumn - distance,
                        candidates);
                }
            }

            return candidates;
        }

        private static void AddCandidate(
            dynamic sheet,
            ExcelSelectionSnapshot snapshot,
            int column,
            ICollection<string> candidates)
        {
            if (column < 1 ||
                column > ExcelSelectionOutputPolicy.MaxExcelColumns ||
                column == snapshot.StartColumn)
            {
                return;
            }

            dynamic range = DestinationRange(
                sheet,
                snapshot,
                column);
            if (IsWritable(range, snapshot.RowCount))
            {
                candidates.Add(
                    ExcelSelectionOutputPolicy.ColumnNumberToName(
                        column));
            }
        }
    }
}
