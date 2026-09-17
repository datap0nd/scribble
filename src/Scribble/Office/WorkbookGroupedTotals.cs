using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Scribble.Office
{
    // Host arithmetic for "analyze by group" requests. A model that adds a
    // hundred ledger rows in its head produces totals no source contains, and
    // the evidence gate rightly refuses them. This read-only pivot sums with
    // decimal arithmetic and returns a small table that is itself a verified
    // read receipt. Blank or non-numeric cells are disclosed, never imputed.
    public static class WorkbookGroupedTotals
    {
        public const int MaxGroupColumns = 3;
        public const int MaxSumColumns = 6;
        public const int MaxGroups = 200;

        public sealed class Result
        {
            public string Table { get; set; }
            public int SourceRows { get; set; }
            public int MatchedRows { get; set; }
            public int Groups { get; set; }
            public int SkippedCells { get; set; }
        }

        // table[0] holds the header row; every cell is invariant text.
        public static Result Compute(IReadOnlyList<IReadOnlyList<string>> table, IReadOnlyList<string> groupBy,
            IReadOnlyList<string> sumColumns, string filterColumn, string filterEquals)
        {
            if (table == null || table.Count < 2) throw new InvalidOperationException("The range needs a header row and at least one data row.");
            if (groupBy == null || groupBy.Count == 0 || groupBy.Count > MaxGroupColumns)
                throw new InvalidOperationException("group_by needs 1 to " + MaxGroupColumns + " header names.");
            if (sumColumns == null || sumColumns.Count == 0 || sumColumns.Count > MaxSumColumns)
                throw new InvalidOperationException("sum_columns needs 1 to " + MaxSumColumns + " header names.");
            var headers = table[0].Select(cell => (cell ?? "").Trim()).ToArray();
            var groupIndexes = groupBy.Select(name => HeaderIndex(headers, name)).ToArray();
            var sumIndexes = sumColumns.Select(name => HeaderIndex(headers, name)).ToArray();
            var filterIndex = string.IsNullOrWhiteSpace(filterColumn) ? -1 : HeaderIndex(headers, filterColumn);
            var wanted = (filterEquals ?? "").Trim();

            var order = new List<string>();
            var keys = new Dictionary<string, string[]>(StringComparer.Ordinal);
            var sums = new Dictionary<string, decimal[]>(StringComparer.Ordinal);
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            var skipped = new Dictionary<string, int>(StringComparer.Ordinal);
            var matched = 0;
            for (var row = 1; row < table.Count; row++)
            {
                var cells = table[row];
                if (cells == null || cells.All(string.IsNullOrWhiteSpace)) continue;
                if (filterIndex >= 0 && !string.Equals(Cell(cells, filterIndex), wanted, StringComparison.OrdinalIgnoreCase)) continue;
                matched++;
                var parts = groupIndexes.Select(index => Cell(cells, index)).ToArray();
                var key = string.Join("\t", parts);
                if (!sums.ContainsKey(key))
                {
                    if (order.Count >= MaxGroups) throw new InvalidOperationException("More than " + MaxGroups + " groups. Group by a coarser column or add a filter.");
                    order.Add(key); keys[key] = parts; sums[key] = new decimal[sumIndexes.Length]; counts[key] = 0; skipped[key] = 0;
                }
                counts[key]++;
                for (var i = 0; i < sumIndexes.Length; i++)
                {
                    decimal value;
                    if (decimal.TryParse(Cell(cells, sumIndexes[i]), NumberStyles.Float, CultureInfo.InvariantCulture, out value)) sums[key][i] += value;
                    else skipped[key]++;
                }
            }
            if (matched == 0) throw new InvalidOperationException("No data row matched the filter. Check the literal cell text of " + filterColumn + ".");

            var text = new StringBuilder();
            text.Append(string.Join("\t", groupIndexes.Select(index => headers[index]))).Append("\tRows\t")
                .Append(string.Join("\t", sumIndexes.Select(index => headers[index]))).Append("\tBlank or non-numeric cells");
            foreach (var key in order)
                text.Append('\n').Append(string.Join("\t", keys[key])).Append('\t').Append(counts[key]).Append('\t')
                    .Append(string.Join("\t", sums[key].Select(Format))).Append('\t').Append(skipped[key]);
            if (order.Count > 1)
            {
                var totals = new decimal[sumIndexes.Length];
                foreach (var key in order) for (var i = 0; i < totals.Length; i++) totals[i] += sums[key][i];
                text.Append('\n').Append("All groups").Append(new string('\t', groupIndexes.Length)).Append(order.Sum(key => counts[key])).Append('\t')
                    .Append(string.Join("\t", totals.Select(Format))).Append('\t').Append(order.Sum(key => skipped[key]));
            }
            return new Result { Table = text.ToString(), SourceRows = table.Count - 1, MatchedRows = matched,
                Groups = order.Count, SkippedCells = order.Sum(key => skipped[key]) };
        }

        private static string Format(decimal value) { return value.ToString("0.############", CultureInfo.InvariantCulture); }
        private static string Cell(IReadOnlyList<string> cells, int index) { return index < cells.Count ? (cells[index] ?? "").Trim() : ""; }

        private static int HeaderIndex(string[] headers, string name)
        {
            var wanted = (name ?? "").Trim();
            var matches = Enumerable.Range(0, headers.Length)
                .Where(index => headers[index].Length > 0 && string.Equals(headers[index], wanted, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matches.Length == 1) return matches[0];
            throw new InvalidOperationException((matches.Length == 0 ? "No header named '" : "More than one header named '") + wanted +
                "'. Use one literal header from the first row of the range: " + string.Join(", ", headers.Where(header => header.Length > 0).Take(40)) + ".");
        }
    }
}
