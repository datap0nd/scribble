using System;
using System.Collections.Generic;
using System.Globalization;
using Scribble.Security;
using Scribble.Utilities;

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
    // source cells change only under explicit replacement intent,
    // and Scribble never saves the file.
    internal static class WorkbookSelectionOutputWriter
    {
        internal static IReadOnlyList<Dictionary<string, string>>
            ReadKoreanSourceWindow(
                KoreanWorkbookSnapshot snapshot,
                int startOffset,
                int count)
        {
            if (snapshot == null ||
                startOffset < 0 ||
                startOffset > snapshot.Cells.Count ||
                count < 0)
            {
                throw new ArgumentOutOfRangeException();
            }

            var actual = Math.Min(
                count,
                snapshot.Cells.Count - startOffset);
            var result = new List<Dictionary<string, string>>();
            var characters = 0;
            var characterWindow = Math.Max(
                ExcelSelectionOutputPolicy.MaxCellCharacters,
                ContextScale.Scaled(
                    TextBoundary.MaxToolResultCharacters) / 3);
            for (var index = 0; index < actual; index++)
            {
                var cell = snapshot.Cells[startOffset + index];
                var addition = cell.WorksheetName.Length +
                    cell.Address.Length + cell.SourceText.Length;
                if (result.Count > 0 &&
                    characters + addition > characterWindow)
                {
                    break;
                }

                result.Add(new Dictionary<string, string>
                {
                    { "worksheet", cell.WorksheetName },
                    { "address", cell.Address },
                    { "source", cell.SourceText }
                });
                characters += addition;
            }

            return result;
        }

        internal static IReadOnlyList<string> ReadSourceValues(
            object excelApplication,
            ExcelSelectionSnapshot snapshot,
            int startOffset,
            int count)
        {
            if (startOffset < 0 ||
                startOffset > snapshot.RowCount ||
                count < 0)
            {
                throw new ArgumentOutOfRangeException();
            }

            var actual = Math.Min(
                count,
                snapshot.RowCount - startOffset);
            var values = new List<string>(actual);
            if (actual == 0)
            {
                return values;
            }

            dynamic sheet = ResolveSheet(excelApplication, snapshot);
            dynamic page = sheet.Range(
                sheet.Cells[
                    snapshot.StartRow + startOffset,
                    snapshot.StartColumn],
                sheet.Cells[
                    snapshot.StartRow + startOffset + actual - 1,
                    snapshot.StartColumn]);
            object raw = page.Value2;
            var grid = raw as object[,];
            if (grid == null)
            {
                values.Add(SourceText(raw));
                return values;
            }

            var rowBase = grid.GetLowerBound(0);
            var columnBase = grid.GetLowerBound(1);
            var characters = 0;
            var characterWindow = Math.Max(
                ExcelSelectionOutputPolicy.MaxCellCharacters,
                ContextScale.Scaled(
                    TextBoundary.MaxToolResultCharacters) / 3);
            for (var row = 0; row < actual; row++)
            {
                var text = SourceText(
                    grid[rowBase + row, columnBase]);
                if (values.Count > 0 &&
                    characters + text.Length > characterWindow)
                {
                    break;
                }

                values.Add(text);
                characters += text.Length;
            }

            return values;
        }

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

        internal static void ValidateSource(
            object excelApplication,
            ExcelSelectionSnapshot snapshot)
        {
            dynamic sheet = ResolveSheet(
                excelApplication,
                snapshot);
            dynamic source = sheet.Range(snapshot.Address);
            if (!IsFalse(source.MergeCells))
            {
                throw new InvalidOperationException(
                    "The selected source contains merged cells and " +
                    "cannot be safely replaced.");
            }
        }

        internal static string Commit(
            object excelApplication,
            ExcelSelectionSnapshot snapshot,
            string destinationColumn,
            IReadOnlyList<string> values,
            bool replaceSource)
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
            dynamic destination = replaceSource
                ? sheet.Range(snapshot.Address)
                : DestinationRange(
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
            var rangeAddress = destinationColumn + snapshot.StartRow + ":" +
                destinationColumn +
                (snapshot.StartRow + snapshot.RowCount - 1);
            return "Wrote " + snapshot.RowCount +
                " literal values to " + snapshot.WorksheetName +
                "!" + rangeAddress + ". " +
                (replaceSource
                    ? "The explicitly selected source cells were replaced. "
                    : "Source cells were unchanged. ") +
                "Nothing was saved.";
        }

        internal static string CommitKoreanTranslations(
            object excelApplication,
            KoreanWorkbookSnapshot snapshot,
            IReadOnlyList<string> translations)
        {
            if (snapshot == null ||
                translations == null ||
                translations.Count != snapshot.Cells.Count)
            {
                throw new InvalidOperationException(
                    "Complete output requires one translation per " +
                    "detected Korean cell.");
            }

            dynamic workbook = ResolveWorkbook(
                excelApplication,
                snapshot);
            var targets = new List<object>();
            var formats = new List<object>();
            foreach (var source in snapshot.Cells)
            {
                dynamic sheet = FindWorksheet(
                    workbook,
                    source.WorksheetName);
                dynamic cell = sheet.Range(source.Address);
                object hasFormula = cell.HasFormula;
                object merged = cell.MergeCells;
                if (!(hasFormula is bool) ||
                    (bool)hasFormula ||
                    !(merged is bool) ||
                    (bool)merged ||
                    !string.Equals(
                        SourceText(cell.Value2),
                        source.SourceText,
                        StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "Korean source cell " + source.WorksheetName +
                        "!" + source.Address +
                        " changed after discovery. No cells were written.");
                }

                targets.Add((object)cell);
                formats.Add(cell.NumberFormat);
            }

            var written = 0;
            try
            {
                for (var index = 0; index < targets.Count; index++)
                {
                    dynamic target = targets[index];
                    target.NumberFormat = "@";
                    target.Value2 = ExcelSelectionOutputPolicy
                        .SanitizeLiteral(translations[index]);
                    written++;
                }
            }
            catch
            {
                // Multi-sheet sparse writes cannot be one Excel bulk
                // assignment. Roll back every completed cell before
                // surfacing the failure, so the operation remains
                // all-or-nothing from the user's perspective.
                for (var index = written - 1; index >= 0; index--)
                {
                    try
                    {
                        dynamic target = targets[index];
                        target.NumberFormat = formats[index];
                        target.Value2 = snapshot.Cells[index].SourceText;
                    }
                    catch (Exception rollbackException)
                    {
                        Log.Error(
                            "KoreanWorkbookTranslationRollback",
                            rollbackException);
                    }
                }

                throw;
            }

            return "Translated and replaced " + translations.Count +
                " Korean text cells across the active workbook. " +
                "Formula and merged cells were left unchanged. " +
                "Nothing was saved.";
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

        private static dynamic ResolveWorkbook(
            object excelApplication,
            KoreanWorkbookSnapshot snapshot)
        {
            dynamic application = excelApplication;
            dynamic workbook = application.ActiveWorkbook;
            if (workbook == null)
            {
                throw new InvalidOperationException(
                    "The captured workbook is no longer active.");
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
            var windowHandle = 0;
            try
            {
                windowHandle = (int)application.ActiveWindow.Hwnd;
            }
            catch
            {
            }

            var saved = path.Length > 0;
            var identityMatches =
                snapshot.Saved == saved &&
                snapshot.WindowHandle == windowHandle &&
                (saved
                    ? string.Equals(
                        snapshot.WorkbookIdentity,
                        fullName,
                        StringComparison.OrdinalIgnoreCase)
                    : string.Equals(
                        snapshot.WorkbookName,
                        workbookName,
                        StringComparison.OrdinalIgnoreCase));
            if (!identityMatches)
            {
                throw new InvalidOperationException(
                    "The captured workbook or window changed before " +
                    "translation. No Korean cells were changed.");
            }

            return workbook;
        }

        private static dynamic FindWorksheet(
            dynamic workbook,
            string worksheetName)
        {
            foreach (dynamic sheet in workbook.Worksheets)
            {
                if (string.Equals(
                    Convert.ToString(sheet.Name),
                    worksheetName,
                    StringComparison.Ordinal))
                {
                    return sheet;
                }
            }

            throw new InvalidOperationException(
                "Worksheet '" + worksheetName +
                "' changed after Korean text discovery.");
        }

        private static string SourceText(object value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            var text = value is double
                ? ((double)value).ToString(
                    "0.############",
                    CultureInfo.InvariantCulture)
                : Convert.ToString(
                    value,
                    CultureInfo.InvariantCulture) ?? string.Empty;
            return TextBoundary.SingleLine(
                text,
                ExcelSelectionOutputPolicy.MaxCellCharacters);
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
