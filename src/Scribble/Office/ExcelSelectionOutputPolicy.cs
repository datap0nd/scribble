using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Scribble.Security;

namespace Scribble.Office
{
    // Immutable, COM-free description of an Excel selection. The
    // workbook host creates it while Excel is active; request-scoped
    // handles and the output policy consume it without retaining COM
    // objects beyond the callback that captured the selection.
    public sealed class ExcelSelectionSnapshot
    {
        public ExcelSelectionSnapshot(
            string attachmentId,
            bool saved,
            string workbookIdentity,
            string workbookName,
            int windowHandle,
            string worksheetName,
            string address,
            int startRow,
            int startColumn,
            int rowCount,
            int columnCount,
            string preview,
            bool previewTruncated)
        {
            AttachmentId = attachmentId ?? string.Empty;
            Saved = saved;
            WorkbookIdentity = workbookIdentity ?? string.Empty;
            WorkbookName = workbookName ?? string.Empty;
            WindowHandle = windowHandle;
            WorksheetName = worksheetName ?? string.Empty;
            Address = address ?? string.Empty;
            StartRow = startRow;
            StartColumn = startColumn;
            RowCount = rowCount;
            ColumnCount = columnCount;
            Preview = preview ?? string.Empty;
            PreviewTruncated = previewTruncated;
        }

        public string AttachmentId { get; }

        public bool Saved { get; }

        public string WorkbookIdentity { get; }

        public string WorkbookName { get; }

        public int WindowHandle { get; }

        public string WorksheetName { get; }

        public string Address { get; }

        public int StartRow { get; }

        public int StartColumn { get; }

        public int RowCount { get; }

        public int ColumnCount { get; }

        public string Preview { get; }

        public bool PreviewTruncated { get; }

        public string BuildContextText(string requestHandle)
        {
            var header = "Selected Excel cells " + WorksheetName + "!" +
                Address +
                (PreviewTruncated ? " (preview truncated)" : string.Empty);
            if (!string.IsNullOrWhiteSpace(requestHandle))
            {
                header += ":\nSelection handle for this request: " +
                    requestHandle + "\nSelected values";
            }

            return header + ":\n" + Preview;
        }
    }

    public sealed class ExcelSelectionRequestContext
    {
        public ExcelSelectionRequestContext(
            string handle,
            ExcelSelectionSnapshot snapshot,
            bool allowSourceReplacement = false)
        {
            Handle = handle ?? string.Empty;
            Snapshot = snapshot ??
                throw new ArgumentNullException(nameof(snapshot));
            AllowSourceReplacement = allowSourceReplacement;
        }

        public string Handle { get; }

        public ExcelSelectionSnapshot Snapshot { get; }

        public bool AllowSourceReplacement { get; }
    }

    public sealed class ExcelDestinationCellState
    {
        public ExcelDestinationCellState(
            string value,
            bool hasFormula,
            bool isMerged)
        {
            Value = value ?? string.Empty;
            HasFormula = hasFormula;
            IsMerged = isMerged;
        }

        public string Value { get; }

        public bool HasFormula { get; }

        public bool IsMerged { get; }
    }

    public static class ExcelSelectionOutputPolicy
    {
        public const int MaxSelectedCells = 500;
        public const int MaxBatchValues = 125;
        public const int MaxBatchCharacters = 10000;
        public const int MaxBatches = 4;
        public const int MaxCellCharacters = 500;
        public const int MaxExcelColumns = 16384;

        public static bool IdentityMatches(
            ExcelSelectionSnapshot snapshot,
            bool saved,
            string workbookIdentity,
            string workbookName,
            int windowHandle,
            string worksheetName)
        {
            if (snapshot == null ||
                snapshot.Saved != saved ||
                snapshot.WindowHandle != windowHandle ||
                !string.Equals(
                    snapshot.WorksheetName,
                    worksheetName ?? string.Empty,
                    StringComparison.Ordinal))
            {
                return false;
            }

            return saved
                ? string.Equals(
                    snapshot.WorkbookIdentity,
                    workbookIdentity ?? string.Empty,
                    StringComparison.OrdinalIgnoreCase)
                : string.Equals(
                      snapshot.WorkbookName,
                      workbookName ?? string.Empty,
                      StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsDestinationWritable(
            IEnumerable<ExcelDestinationCellState> cells)
        {
            if (cells == null)
            {
                return false;
            }

            foreach (var cell in cells)
            {
                if (cell == null ||
                    cell.HasFormula ||
                    cell.IsMerged ||
                    cell.Value.Length > 0)
                {
                    return false;
                }
            }

            return true;
        }

        public static string SanitizeLiteral(string value)
        {
            var bounded = TextBoundary.PlainText(
                    value,
                    MaxCellCharacters)
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Replace('\t', ' ');
            if (bounded.Length > 0 &&
                (bounded[0] == '=' ||
                 bounded[0] == '+' ||
                 bounded[0] == '-' ||
                 bounded[0] == '@'))
            {
                return "'" + bounded;
            }

            return bounded;
        }

        // Source replacement is a separate, explicit user choice.
        // Negated and preservation phrases win over the positive
        // terms so document or model wording cannot broaden it.
        public static bool AllowsSourceReplacement(string userText)
        {
            var text = TextBoundary.PlainText(userText, 500)
                .ToLowerInvariant();
            if (text.Length == 0 ||
                ContainsAny(
                    text,
                    "do not replace",
                    "don't replace",
                    "dont replace",
                    "without replacing",
                    "do not overwrite",
                    "don't overwrite",
                    "dont overwrite",
                    "keep the source",
                    "keep source",
                    "preserve the source",
                    "preserve source"))
            {
                return false;
            }

            return ContainsAny(
                text,
                "replace",
                "overwrite",
                "in place",
                "in-place",
                "same cells",
                "source cells",
                "original cells");
        }

        private static bool ContainsAny(
            string value,
            params string[] terms)
        {
            foreach (var term in terms)
            {
                if (value.IndexOf(
                    term,
                    StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        public static int ColumnNameToNumber(string value)
        {
            var name = (value ?? string.Empty)
                .Trim()
                .ToUpperInvariant();
            if (name.Length == 0 || name.Length > 3)
            {
                return 0;
            }

            var number = 0;
            foreach (var character in name)
            {
                if (character < 'A' || character > 'Z')
                {
                    return 0;
                }

                number = (number * 26) + (character - 'A' + 1);
            }

            return number <= MaxExcelColumns ? number : 0;
        }

        public static string ColumnNumberToName(int number)
        {
            if (number < 1 || number > MaxExcelColumns)
            {
                return string.Empty;
            }

            var result = string.Empty;
            while (number > 0)
            {
                number--;
                result = (char)('A' + (number % 26)) + result;
                number /= 26;
            }

            return result;
        }
    }

    // Request-scoped, COM-free batch assembler. It does not grant
    // permission or touch Excel; the draft host consumes permission
    // only after this session reports a complete one-to-one result.
    public sealed class ExcelSelectionOutputSession
    {
        private readonly string _handle;
        private readonly int _expectedValues;
        private readonly List<string> _values = new List<string>();
        private string _destinationColumn = string.Empty;
        private int _batches;
        private bool _complete;

        public ExcelSelectionOutputSession(
            string handle,
            int expectedValues)
        {
            if (string.IsNullOrWhiteSpace(handle))
            {
                throw new ArgumentException(
                    "A request-scoped selection handle is required.",
                    nameof(handle));
            }

            if (expectedValues < 1 ||
                expectedValues >
                    ExcelSelectionOutputPolicy.MaxSelectedCells)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(expectedValues));
            }

            _handle = handle;
            _expectedValues = expectedValues;
        }

        public string DestinationColumn
        {
            get { return _destinationColumn; }
        }

        public int StagedCount
        {
            get { return _values.Count; }
        }

        public bool IsComplete
        {
            get { return _complete; }
        }

        public IReadOnlyList<string> Values
        {
            get
            {
                return new ReadOnlyCollection<string>(
                    new List<string>(_values));
            }
        }

        public bool Stage(
            string handle,
            string destinationColumn,
            int startOffset,
            IReadOnlyList<string> values,
            bool complete)
        {
            if (_complete)
            {
                throw new InvalidOperationException(
                    "The selection output is already complete.");
            }

            if (!string.Equals(
                _handle,
                handle ?? string.Empty,
                StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "The selection handle is unknown or expired.");
            }

            var normalizedDestination =
                ExcelSelectionOutputPolicy.ColumnNumberToName(
                    ExcelSelectionOutputPolicy.ColumnNameToNumber(
                        destinationColumn));
            if (normalizedDestination.Length == 0)
            {
                throw new InvalidOperationException(
                    "A valid destination column is required.");
            }

            if (_destinationColumn.Length > 0 &&
                !string.Equals(
                _destinationColumn,
                normalizedDestination,
                StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "All batches must use the same destination column.");
            }

            if (_batches >= ExcelSelectionOutputPolicy.MaxBatches)
            {
                throw new InvalidOperationException(
                    "The selection output exceeded the batch limit.");
            }

            if (startOffset != _values.Count)
            {
                throw new InvalidOperationException(
                    "Selection output batches must be contiguous and ordered.");
            }

            if (values == null || values.Count == 0 ||
                values.Count > ExcelSelectionOutputPolicy.MaxBatchValues)
            {
                throw new InvalidOperationException(
                    "A batch must contain between 1 and " +
                    ExcelSelectionOutputPolicy.MaxBatchValues +
                    " values.");
            }

            var bounded = new List<string>();
            var characters = 0;
            foreach (var value in values)
            {
                var sanitized =
                    ExcelSelectionOutputPolicy.SanitizeLiteral(value);
                characters += sanitized.Length;
                if (characters >
                    ExcelSelectionOutputPolicy.MaxBatchCharacters)
                {
                    throw new InvalidOperationException(
                        "The selection output batch is too large.");
                }

                bounded.Add(sanitized);
            }

            if (_values.Count + bounded.Count > _expectedValues)
            {
                throw new InvalidOperationException(
                    "The selection output contains more values than the source.");
            }

            if (!complete &&
                _values.Count + bounded.Count == _expectedValues)
            {
                throw new InvalidOperationException(
                    "The final batch must set complete to true.");
            }

            if (complete &&
                _values.Count + bounded.Count != _expectedValues)
            {
                throw new InvalidOperationException(
                    "The completed output must contain exactly one value per source row.");
            }

            if (_destinationColumn.Length == 0)
            {
                _destinationColumn = normalizedDestination;
            }

            _values.AddRange(bounded);
            _batches++;
            _complete = complete;
            return _complete;
        }
    }
}
