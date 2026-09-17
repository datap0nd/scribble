using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Scribble.Security;

namespace Scribble.Office
{
    // Pre-write audit of the model's draft-table formulas. The Scribble
    // Draft layout is deterministic (title in A1, header row 3, first
    // data row 4), and a model that counts rows loosely produces formulas
    // that evaluate cleanly yet describe the wrong cells: a same-row rate
    // that mixes in the row above, a total that sums the neighbouring
    // column, or a range that swallows the header. Nothing here evaluates
    // a formula. It only compares the cells a formula names against the
    // table the same call is writing, so every finding is returned before
    // any draft-write permission is spent and the model can retry.
    public static class DraftTableFormulaAudit
    {
        public const string ErrorCode = "DRAFT_FORMULA_ROW_MISMATCH";
        public const int HeaderRow = 3;
        private const int MaxReportedFindings = 6;

        private static readonly Regex QualifiedReference = new Regex(
            @"(?:'[^']*'|[A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Za-z]{1,3}(?:\$?\d+)?(?::\$?[A-Za-z]{1,3}(?:\$?\d+)?)?",
            RegexOptions.Compiled);
        private static readonly Regex StringLiteral = new Regex(@"""[^""]*""", RegexOptions.Compiled);
        private static readonly Regex RangeReference = new Regex(
            @"(?<![A-Za-z0-9_!])\$?(?<c1>[A-Za-z]{1,3})(?<a1>\$?)(?<r1>\d+):\$?(?<c2>[A-Za-z]{1,3})\$?(?<r2>\d+)(?![A-Za-z0-9_(])",
            RegexOptions.Compiled);
        private static readonly Regex CellReference = new Regex(
            @"(?<![A-Za-z0-9_!$])\$?(?<c>[A-Za-z]{1,3})(?<a>\$?)(?<r>\d+)(?![A-Za-z0-9_(])",
            RegexOptions.Compiled);
        private static readonly Regex BareAggregate = new Regex(
            @"^=\s*(?:SUM|AVERAGE|MIN|MAX|COUNT|COUNTA|MEDIAN)\s*\(\s*\$?(?<c1>[A-Za-z]{1,3})\$?(?<r1>\d+):\$?(?<c2>[A-Za-z]{1,3})\$?(?<r2>\d+)\s*\)\s*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex PeriodLabel = new Regex(
            @"(?<![A-Za-z0-9])(?:\d{4}[-/]\d{1,2}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?(?:\s+\d{2,4})?|Q[1-4](?:\s*\d{4})?|H[12](?:\s*\d{4})?|FY\s*\d{2,4}|W(?:eek)?\s*\d{1,2}|\d{4})(?![A-Za-z0-9])",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex TotalLabel = new Regex(
            @"\b(?:total|sum|overall|all|grand|subtotal)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private struct Reference
        {
            public int Column;
            public int Row;
            public bool AbsoluteRow;
        }

        // Throws the model-facing rejection when any formula names cells that
        // cannot belong to the table this call writes.
        public static void Require(IReadOnlyList<IReadOnlyList<string>> rows)
        {
            var findings = Findings(rows);
            if (findings.Count == 0) return;
            var reported = findings.Take(MaxReportedFindings).ToArray();
            throw new InvalidOperationException(
                ErrorCode + ": " + findings.Count +
                (findings.Count == 1 ? " draft formula names" : " draft formulas name") +
                " cells that do not match the table this call writes. The header is row " +
                HeaderRow.ToString(CultureInfo.InvariantCulture) + " and the first data row is row " +
                (HeaderRow + 1).ToString(CultureInfo.InvariantCulture) +
                "; count blank spacer rows and section headers exactly. A same-row rate uses only that row's metric cells, and a total sums only its section's data rows in its own column. " +
                "No draft was written and no write permission was spent. Fix every listed formula, then retry as the only tool call. " +
                string.Join(" ", reported) +
                (findings.Count > reported.Length ? " (" + (findings.Count - reported.Length) + " more.)" : ""));
        }

        public static IReadOnlyList<string> Findings(IReadOnlyList<IReadOnlyList<string>> rows)
        {
            var findings = new List<string>();
            if (rows == null || rows.Count == 0) return findings;
            var rowCount = rows.Count;
            var columnCount = 1;
            foreach (var row in rows)
                if (row != null && row.Count > columnCount) columnCount = row.Count;
            var lastRow = HeaderRow + rowCount - 1;

            for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var row = rows[rowIndex];
                if (row == null) continue;
                for (var columnIndex = 0; columnIndex < row.Count; columnIndex++)
                {
                    var formula = (row[columnIndex] ?? "").Trim();
                    if (formula.Length < 2 || formula[0] != '=' || !DraftFormulaPolicy.IsAllowedFormula(formula)) continue;
                    var cellRow = HeaderRow + rowIndex;
                    var cellColumn = columnIndex + 1;
                    var address = ColumnName(cellColumn) + cellRow.ToString(CultureInfo.InvariantCulture);
                    var shown = address + " '" + Shorten(formula) + "'";

                    // Other worksheets and quoted text are outside this table.
                    var local = StringLiteral.Replace(QualifiedReference.Replace(formula, " "), " ");
                    var ranges = RangeReference.Matches(local).Cast<Match>().ToArray();
                    var withoutRanges = RangeReference.Replace(local, " ");
                    var cells = CellReference.Matches(withoutRanges).Cast<Match>()
                        .Select(match => new Reference
                        {
                            Column = ColumnNumber(match.Groups["c"].Value),
                            Row = int.Parse(match.Groups["r"].Value, CultureInfo.InvariantCulture),
                            AbsoluteRow = match.Groups["a"].Value.Length > 0
                        }).ToArray();
                    var endpoints = ranges.SelectMany(match => new[]
                    {
                        new Reference { Column = ColumnNumber(match.Groups["c1"].Value), Row = int.Parse(match.Groups["r1"].Value, CultureInfo.InvariantCulture) },
                        new Reference { Column = ColumnNumber(match.Groups["c2"].Value), Row = int.Parse(match.Groups["r2"].Value, CultureInfo.InvariantCulture) }
                    }).ToArray();

                    var outside = cells.Concat(endpoints)
                        .Where(reference => reference.Row < HeaderRow || reference.Row > lastRow || reference.Column > columnCount)
                        .Select(reference => ColumnName(reference.Column) + reference.Row.ToString(CultureInfo.InvariantCulture))
                        .Distinct(StringComparer.Ordinal).ToArray();
                    if (outside.Length > 0)
                    {
                        findings.Add(shown + " references " + string.Join(", ", outside) +
                            " outside this draft table (rows " + HeaderRow.ToString(CultureInfo.InvariantCulture) +
                            " to " + lastRow.ToString(CultureInfo.InvariantCulture) + ", columns A to " +
                            ColumnName(columnCount) + ").");
                        continue;
                    }

                    var rangeFinding = false;
                    foreach (var range in ranges)
                    {
                        var from = int.Parse(range.Groups["r1"].Value, CultureInfo.InvariantCulture);
                        var to = int.Parse(range.Groups["r2"].Value, CultureInfo.InvariantCulture);
                        var low = Math.Min(from, to);
                        var high = Math.Max(from, to);
                        if (low <= HeaderRow && HeaderRow <= high)
                        {
                            findings.Add(shown + " includes the header row " + HeaderRow.ToString(CultureInfo.InvariantCulture) +
                                "; start the range at the section's first data row.");
                            rangeFinding = true;
                            break;
                        }
                        if (low <= cellRow && cellRow <= high)
                        {
                            findings.Add(shown + " includes its own cell " + address + ", which is circular; end the range at the last data row above it.");
                            rangeFinding = true;
                            break;
                        }
                        // A numeric aggregate that starts or ends on a label row
                        // (a section header, a spacer caption) is off by one.
                        var fromColumn = ColumnNumber(range.Groups["c1"].Value);
                        var toColumn = ColumnNumber(range.Groups["c2"].Value);
                        if (fromColumn != toColumn || !Regex.IsMatch(local,
                                @"(?:SUM|AVERAGE|MIN|MAX|MEDIAN)\s*\(\s*" + Regex.Escape(range.Value) + @"\s*[,)]",
                                RegexOptions.IgnoreCase)) continue;
                        foreach (var endpoint in new[] { low, high })
                        {
                            var content = CellText(rows, endpoint, fromColumn);
                            if (content.Length == 0 || content[0] == '=' || Regex.IsMatch(content, @"^[-+]?[\d,.]+%?$")) continue;
                            findings.Add(shown + " includes the text cell " + ColumnName(fromColumn) + endpoint.ToString(CultureInfo.InvariantCulture) +
                                " ('" + content + "'); start and end the range on the section's data rows.");
                            rangeFinding = true;
                            break;
                        }
                        if (rangeFinding) break;
                    }
                    if (rangeFinding) continue;

                    var aggregate = BareAggregate.Match(formula);
                    if (aggregate.Success)
                    {
                        var fromColumn = ColumnNumber(aggregate.Groups["c1"].Value);
                        var toColumn = ColumnNumber(aggregate.Groups["c2"].Value);
                        if (fromColumn == toColumn && fromColumn != cellColumn && cellColumn > 1)
                        {
                            findings.Add(shown + " totals column " + ColumnName(fromColumn) + " from column " + ColumnName(cellColumn) +
                                "; a total sums the data rows of its own column, for example =SUM(" + ColumnName(cellColumn) +
                                aggregate.Groups["r1"].Value + ":" + ColumnName(cellColumn) + aggregate.Groups["r2"].Value + ").");
                            continue;
                        }
                    }

                    if (ranges.Length > 0 || cells.Length == 0) continue;
                    var referencedRows = cells.Select(reference => reference.Row).Distinct().OrderBy(value => value).ToArray();
                    if (referencedRows.Length != 2 || !referencedRows.Contains(cellRow)) continue;
                    var adjacentRow = referencedRows.First(value => value != cellRow);
                    if (Math.Abs(adjacentRow - cellRow) != 1) continue;
                    var adjacentReferences = cells.Where(reference => reference.Row == adjacentRow).ToArray();
                    // An anchored row ($11) is a deliberate base such as a total.
                    if (adjacentReferences.Any(reference => reference.AbsoluteRow)) continue;
                    var ownLabel = Label(rows, cellRow);
                    var adjacentLabel = Label(rows, adjacentRow);
                    if (adjacentRow != HeaderRow)
                    {
                        // Period rows legitimately compare with their neighbour
                        // (month-over-month), and shares are taken against totals.
                        if (PeriodLabel.IsMatch(ownLabel) || PeriodLabel.IsMatch(adjacentLabel) || TotalLabel.IsMatch(adjacentLabel)) continue;
                        // A different column per row is a running figure, not a shifted rate.
                        var ownColumns = cells.Where(reference => reference.Row == cellRow).Select(reference => reference.Column);
                        if (!adjacentReferences.Select(reference => reference.Column).Intersect(ownColumns).Any()) continue;
                    }
                    findings.Add(shown + " mixes row " + adjacentRow.ToString(CultureInfo.InvariantCulture) +
                        (adjacentRow == HeaderRow ? " (the header)" : (adjacentLabel.Length > 0 ? " ('" + adjacentLabel + "')" : "")) +
                        " with its own row " + cellRow.ToString(CultureInfo.InvariantCulture) +
                        (ownLabel.Length > 0 ? " ('" + ownLabel + "')" : "") +
                        "; a same-row rate uses only row " + cellRow.ToString(CultureInfo.InvariantCulture) + " cells.");
                }
            }
            return findings;
        }

        private static string Label(IReadOnlyList<IReadOnlyList<string>> rows, int sheetRow)
        {
            return CellText(rows, sheetRow, 1);
        }

        private static string CellText(IReadOnlyList<IReadOnlyList<string>> rows, int sheetRow, int column)
        {
            var index = sheetRow - HeaderRow;
            if (index < 0 || index >= rows.Count || rows[index] == null || rows[index].Count < column) return "";
            var text = (rows[index][column - 1] ?? "").Trim();
            return text.Length > 40 ? text.Substring(0, 40) : text;
        }

        private static string Shorten(string formula)
        {
            return formula.Length <= 80 ? formula : formula.Substring(0, 77) + "...";
        }

        private static int ColumnNumber(string letters)
        {
            var value = 0;
            foreach (var letter in letters.ToUpperInvariant()) value = value * 26 + (letter - 'A' + 1);
            return value;
        }

        private static string ColumnName(int column)
        {
            var name = "";
            while (column > 0)
            {
                var remainder = (column - 1) % 26;
                name = (char)('A' + remainder) + name;
                column = (column - 1) / 26;
            }
            return name;
        }
    }
}
