using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace Scribble.Office
{
    // Read-only pilot adapter for a sales ledger. Chart amounts are recomputed
    // from literal operands; cached formula results are cross-checks only.
    public static class WorkbookMonthlyChartFacts
    {
        public sealed class Result
        {
            public string SourceSha256 { get; set; }
            public string SnapshotId { get; set; }
            public string[] Categories { get; set; }
            public decimal[] RevenueEur { get; set; }
            public decimal[] CostEur { get; set; }
            public int SourceRows { get; set; }
            public int UniqueRows { get; set; }
        }

        public static Result ReadSalesLedger(string path,
            CancellationToken cancellationToken)
        {
            var before = Hash(path);
            var snapshot = OpenXmlWorkbookSnapshotReader.Capture(path,
                Path.GetFullPath(path), before, cancellationToken);
            var table = snapshot.Tables.SingleOrDefault(item =>
                item.Name == "Ledger") ?? throw new InvalidOperationException(
                    "MONTHLY_LEDGER_MISSING");
            if (table.Rows < 13 || table.Rows > 20000 ||
                table.Columns < 10 || table.Columns > 256)
                throw new InvalidOperationException("MONTHLY_LEDGER_SCOPE_INVALID");
            var cells = table.Cells.ToDictionary(cell =>
                cell.Row.ToString(CultureInfo.InvariantCulture) + ":" +
                cell.Column.ToString(CultureInfo.InvariantCulture),
                StringComparer.Ordinal);
            Func<int, int, DatasetCell> at = (row, column) =>
            {
                DatasetCell cell;
                cells.TryGetValue(row.ToString(CultureInfo.InvariantCulture) +
                    ":" + column.ToString(CultureInfo.InvariantCulture),
                    out cell);
                return cell;
            };
            var id = Column(table, at, "RowID");
            var period = Column(table, at, "Period");
            var group = Column(table, at, "Group");
            var item = Column(table, at, "Item");
            var units = Column(table, at, "Units");
            var price = Column(table, at, "UnitPriceEUR");
            var cost = Column(table, at, "UnitCostEUR");
            var revenue = Column(table, at, "RevenueEUR");
            var costTotal = Column(table, at, "CostEUR");
            var seen = new Dictionary<string, string>(StringComparer.Ordinal);
            var sums = new SortedDictionary<string, decimal[]>(
                StringComparer.Ordinal);
            for (var row = 1; row < table.Rows; row++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var rowId = Literal(at(row, id));
                var month = Literal(at(row, period));
                var grouping = Literal(at(row, group));
                var product = Literal(at(row, item));
                if (string.IsNullOrWhiteSpace(rowId) ||
                    string.IsNullOrWhiteSpace(grouping) ||
                    string.IsNullOrWhiteSpace(product) ||
                    !Regex.IsMatch(month, @"^\d{4}-(0[1-9]|1[0-2])$"))
                    throw new InvalidOperationException(
                        "MONTHLY_LEDGER_ROW_INVALID: " + (row + 1));
                var unitCount = Decimal(at(row, units));
                var unitPrice = Decimal(at(row, price));
                var unitCost = Decimal(at(row, cost));
                var signature = string.Join("|", new[] { month, grouping,
                    product, unitCount.ToString(CultureInfo.InvariantCulture),
                    unitPrice.ToString(CultureInfo.InvariantCulture),
                    unitCost.ToString(CultureInfo.InvariantCulture) });
                string prior;
                if (seen.TryGetValue(rowId, out prior))
                {
                    if (prior != signature)
                        throw new InvalidOperationException(
                            "MONTHLY_LEDGER_DUPLICATE_CONFLICT: " + rowId);
                    continue;
                }
                seen.Add(rowId, signature);
                var amount = checked(unitCount * unitPrice);
                var expense = checked(unitCount * unitCost);
                VerifyCachedAmount(at(row, revenue), amount);
                VerifyCachedAmount(at(row, costTotal), expense);
                decimal[] totals;
                if (!sums.TryGetValue(month, out totals))
                    sums[month] = totals = new decimal[2];
                totals[0] = checked(totals[0] + amount);
                totals[1] = checked(totals[1] + expense);
            }
            if (sums.Count != 6 || seen.Count < 12)
                throw new InvalidOperationException(
                    "MONTHLY_LEDGER_PERIOD_COVERAGE_INVALID");
            var after = Hash(path);
            if (before != after)
                throw new InvalidOperationException(
                    "MONTHLY_LEDGER_SOURCE_CHANGED");
            return new Result
            {
                SourceSha256 = before,
                SnapshotId = snapshot.SnapshotId,
                Categories = sums.Keys.ToArray(),
                RevenueEur = sums.Values.Select(value => value[0]).ToArray(),
                CostEur = sums.Values.Select(value => value[1]).ToArray(),
                SourceRows = table.Rows - 1,
                UniqueRows = seen.Count
            };
        }

        private static int Column(TableDataset table,
            Func<int, int, DatasetCell> at, string name)
        {
            var matches = Enumerable.Range(0, table.Columns).Where(index =>
                Literal(at(0, index)) == name).ToArray();
            if (matches.Length != 1)
                throw new InvalidOperationException(
                    "MONTHLY_LEDGER_HEADER_INVALID: " + name);
            return matches[0];
        }

        private static string Literal(DatasetCell cell)
        {
            if (cell == null || cell.Status != AnalysisContract.Verified ||
                !string.IsNullOrEmpty(cell.Formula))
                return string.Empty;
            return cell.Value ?? string.Empty;
        }

        private static decimal Decimal(DatasetCell cell)
        {
            decimal number;
            if (!decimal.TryParse(Literal(cell), NumberStyles.Float,
                    CultureInfo.InvariantCulture, out number))
                throw new InvalidOperationException(
                    "MONTHLY_LEDGER_OPERAND_UNVERIFIED");
            return number;
        }

        private static void VerifyCachedAmount(DatasetCell cell,
            decimal calculated)
        {
            decimal cached;
            if (cell == null || !decimal.TryParse(cell.RawValue,
                    NumberStyles.Float, CultureInfo.InvariantCulture,
                    out cached) || Math.Abs(cached - calculated) > .01m)
                throw new InvalidOperationException(
                    "MONTHLY_LEDGER_CACHE_MISMATCH");
        }

        private static string Hash(string path)
        {
            using (var stream = File.OpenRead(path))
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(stream))
                    .Replace("-", "").ToLowerInvariant();
        }
    }
}
