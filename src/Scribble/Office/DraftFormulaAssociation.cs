using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Scribble.Office
{
    // Preflight for model-written formulas in a new Scribble Draft sheet. The
    // layout is deterministic (rows[0] lands on sheet row 3, column A), so the
    // host can resolve every same-sheet reference before anything is written.
    // A model that plans one table layout and emits another produces formulas
    // such as D13 =(B13-B14)/B13: they evaluate to a plausible number from the
    // wrong cells. Only definite association faults are reported; formulas that
    // use functions, other sheets, or that stand alone are left to Excel.
    public static class DraftFormulaAssociation
    {
        public const int DraftFirstRow = 3;
        public const int MaxReportedIssues = 12;

        private sealed class CellReference
        {
            public int Row, Column, EndRow, EndColumn;
            public bool IsRange;
            public string Text;
        }

        private sealed class ParsedFormula
        {
            public int Row, Column;
            public string Formula;
            public string Skeleton;
            public List<CellReference> References = new List<CellReference>();
            public bool HasFunction;
            public bool IsAggregate;
        }

        private static readonly Regex StringLiteral = new Regex("\"[^\"]*\"", RegexOptions.Compiled);
        private static readonly Regex AnyReference = new Regex(
            @"(?<sheet>(?:'[^']+'|[A-Za-z0-9_.]+)!)?(?<![A-Za-z0-9_.$])\$?(?<c1>[A-Za-z]{1,3})\$?(?<r1>[1-9][0-9]{0,6})(?::\$?(?<c2>[A-Za-z]{1,3})\$?(?<r2>[1-9][0-9]{0,6}))?(?![A-Za-z0-9_(])",
            RegexOptions.Compiled);
        private static readonly Regex FunctionCall = new Regex(@"[A-Za-z_][A-Za-z0-9_.]*\s*\(", RegexOptions.Compiled);
        private static readonly Regex Aggregate = new Regex(@"^=\s*(?:SUM|AVERAGE|MIN|MAX|COUNT|COUNTA)\s*\(\s*\{ref\}\s*\)\s*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static IReadOnlyList<string> Validate(IReadOnlyList<IReadOnlyList<string>> rows)
        {
            var issues = new List<string>();
            if (rows == null) return issues;
            var formulas = new Dictionary<long, ParsedFormula>();
            for (var i = 0; i < rows.Count; i++)
                for (var j = 0; rows[i] != null && j < rows[i].Count; j++)
                {
                    var parsed = Parse(rows[i][j], DraftFirstRow + i, j + 1);
                    if (parsed != null) formulas[Key(parsed.Row, parsed.Column)] = parsed;
                }
            foreach (var formula in formulas.Values.OrderBy(f => f.Row).ThenBy(f => f.Column))
            {
                if (issues.Count >= MaxReportedIssues) break;
                var issue = formula.IsAggregate ? CheckAggregate(formula, rows, formulas)
                    : formula.HasFunction ? null : CheckArithmetic(formula, rows, formulas);
                if (issue != null) issues.Add(Address(formula.Row, formula.Column) + " " + formula.Formula + ": " + issue);
            }
            return issues;
        }

        public static string RepairMessage(IReadOnlyList<string> issues)
        {
            return "DRAFT_FORMULA_ASSOCIATION: No sheet was written and no draft permission was consumed. " +
                "The draft table always starts at A3: rows[0] is sheet row 3, rows[i] is sheet row i+3, and column index j is letter j+1 (A, B, C...). " +
                "A per-row metric must reference its own row's cells; a total must cover exactly the data rows above it, never the header. " +
                "Recount every formula reference against that layout and resend the complete table. " +
                string.Join(" ", issues ?? new string[0]);
        }

        public static IReadOnlyList<string> ValidatePromptRequirements(
            string prompt,
            IReadOnlyList<IReadOnlyList<string>> rows)
        {
            var issues = new List<string>();
            var instruction = prompt ?? "";
            if (!Regex.IsMatch(instruction, @"(?is)\b(?:live\s+(?:Excel\s+)?|linked\s+)formulas?\b"))
                return issues;

            var required = Regex.Matches(instruction,
                    @"(?is)\b(?:live\s+(?:Excel\s+)?|linked\s+)formulas?\s+(?:for|in)\s+([A-Z]{1,3}[1-9][0-9]*)(?::([A-Z]{1,3}[1-9][0-9]*))?")
                .Cast<Match>().ToArray();
            foreach (var match in required)
            {
                var first = ParseAddress(match.Groups[1].Value);
                var last = match.Groups[2].Success ? ParseAddress(match.Groups[2].Value) : first;
                for (var row = Math.Min(first.Item1, last.Item1); row <= Math.Max(first.Item1, last.Item1); row++)
                    for (var column = Math.Min(first.Item2, last.Item2); column <= Math.Max(first.Item2, last.Item2); column++)
                    {
                        var value = CellText(rows, row, column);
                        if (!value.StartsWith("=", StringComparison.Ordinal))
                            issues.Add(Address(row, column) + " must be a live formula, not the pasted value '" + Shorten(value) + "'.");
                        else if (value.IndexOf('!') < 0)
                            issues.Add(Address(row, column) + " must link to a source worksheet as requested.");
                    }
            }
            if (required.Length == 0 && !(rows ?? new IReadOnlyList<string>[0])
                .Where(row => row != null).SelectMany(row => row)
                .Any(value => (value ?? "").TrimStart().StartsWith("=", StringComparison.Ordinal)))
                issues.Add("The requested live formulas are missing from the draft table.");
            return issues.Take(MaxReportedIssues).ToArray();
        }

        private static ParsedFormula Parse(string cell, int row, int column)
        {
            var text = (cell ?? "").Trim();
            if (text.Length < 2 || text[0] != '=') return null;
            var stripped = StringLiteral.Replace(text, "\"\"");
            var external = false;
            var parsed = new ParsedFormula { Row = row, Column = column, Formula = text };
            var skeleton = AnyReference.Replace(stripped, match =>
            {
                if (match.Groups["sheet"].Success) { external = true; return match.Value; }
                var reference = new CellReference
                {
                    Column = ColumnIndex(match.Groups["c1"].Value), Row = int.Parse(match.Groups["r1"].Value, CultureInfo.InvariantCulture),
                    IsRange = match.Groups["c2"].Success, Text = match.Value.Replace("$", "").ToUpperInvariant()
                };
                reference.EndColumn = reference.IsRange ? ColumnIndex(match.Groups["c2"].Value) : reference.Column;
                reference.EndRow = reference.IsRange ? int.Parse(match.Groups["r2"].Value, CultureInfo.InvariantCulture) : reference.Row;
                parsed.References.Add(reference);
                return "{ref}";
            });
            // Cross-sheet formulas read the user's data; their criteria differ
            // legitimately between neighbours and are verified by Excel itself.
            if (external || parsed.References.Count == 0) return null;
            parsed.Skeleton = Regex.Replace(skeleton, @"\s+", "").ToUpperInvariant();
            parsed.IsAggregate = parsed.References.Count == 1 && Aggregate.IsMatch(skeleton);
            parsed.HasFunction = FunctionCall.IsMatch(skeleton);
            return parsed;
        }

        private static string CheckArithmetic(ParsedFormula formula, IReadOnlyList<IReadOnlyList<string>> rows,
            Dictionary<long, ParsedFormula> formulas)
        {
            if (formula.References.Any(reference => reference.IsRange)) return null;
            foreach (var reference in formula.References)
            {
                var value = CellText(rows, reference.Row, reference.Column);
                if (value.Length == 0)
                    return "operand " + reference.Text + " is blank in this draft, so the arithmetic uses the wrong cell.";
                if (!IsNumericOrFormula(value))
                    return "operand " + reference.Text + " is the text cell '" + Shorten(value) + "', so the arithmetic cannot evaluate.";
            }
            if (formula.References.All(reference => reference.Row == formula.Row) ||
                formula.References.All(reference => reference.Column == formula.Column)) return null;
            // Only a record row can be misassociated: the formula sits beside
            // its own row's quantities and ignores them. A derived-metric list
            // under a table (B7 =C4-B4) has a label and nothing else in its row.
            if (!InRecordRow(formula, rows)) return null;
            if (!InconsistentWithEveryNeighbour(formula, formulas)) return null;
            return "its operands are neither all in row " + formula.Row + " nor all in column " + ColumnName(formula.Column) +
                ", and it does not follow the adjacent formulas. A per-row metric must use the cells of row " + formula.Row + ".";
        }

        private static string CheckAggregate(ParsedFormula formula, IReadOnlyList<IReadOnlyList<string>> rows,
            Dictionary<long, ParsedFormula> formulas)
        {
            var range = formula.References[0];
            var top = Math.Min(range.Row, range.EndRow);
            var bottom = Math.Max(range.Row, range.EndRow);
            var singleColumn = range.Column == range.EndColumn;
            var singleRow = top == bottom;
            if (singleColumn && range.Column == formula.Column && bottom < formula.Row && bottom >= formula.Row - 3)
            {
                // The contiguous numeric block directly above a total is its data.
                var end = formula.Row - 1;
                if (CellText(rows, end, formula.Column).Length == 0) end--;
                var start = end;
                while (start >= DraftFirstRow && IsNumericOrFormula(CellText(rows, start, formula.Column))) start--;
                start++;
                // A numeric-looking header (a year) may sit directly above the
                // data, so a shorter top is tolerated; a missed last data row
                // or an included text header is not.
                var includesText = Enumerable.Range(top, Math.Max(0, start - top))
                    .Any(row => CellText(rows, row, formula.Column).Length > 0);
                if (start <= end && (bottom != end || includesText))
                    return "the total must cover exactly the data block " + ColumnName(formula.Column) + start + ":" +
                        ColumnName(formula.Column) + end + " directly above it, not " + range.Text + ".";
                return null;
            }
            if ((singleColumn && range.Column == formula.Column) || (singleRow && top == formula.Row)) return null;
            if (!InRecordRow(formula, rows)) return null;
            if (!InconsistentWithEveryNeighbour(formula, formulas)) return null;
            return "its range " + range.Text + " is outside column " + ColumnName(formula.Column) + " and row " + formula.Row +
                ", and it does not follow the adjacent totals. Total the data rows of its own column.";
        }

        private static bool InRecordRow(ParsedFormula formula, IReadOnlyList<IReadOnlyList<string>> rows)
        {
            var row = rows[formula.Row - DraftFirstRow];
            var quantities = 0;
            for (var column = 1; row != null && column <= row.Count; column++)
                if (column != formula.Column && IsNumericOrFormula(CellText(rows, formula.Row, column))) quantities++;
            return quantities >= 2;
        }

        // True only when at least one adjacent same-sheet formula exists and
        // none of them is the same formula filled across or down.
        private static bool InconsistentWithEveryNeighbour(ParsedFormula formula, Dictionary<long, ParsedFormula> formulas)
        {
            var neighbours = new[] { new[] { -1, 0 }, new[] { 1, 0 }, new[] { 0, -1 }, new[] { 0, 1 } }
                .Select(offset => { ParsedFormula other; return formulas.TryGetValue(Key(formula.Row + offset[0], formula.Column + offset[1]), out other) ? other : null; })
                .Where(other => other != null).ToArray();
            return neighbours.Length > 0 && !neighbours.Any(other => Consistent(formula, other));
        }

        private static bool Consistent(ParsedFormula a, ParsedFormula b)
        {
            if (a.Skeleton != b.Skeleton || a.References.Count != b.References.Count) return false;
            for (var i = 0; i < a.References.Count; i++)
            {
                var x = a.References[i]; var y = b.References[i];
                var relative = x.Row - a.Row == y.Row - b.Row && x.Column - a.Column == y.Column - b.Column &&
                    x.EndRow - a.Row == y.EndRow - b.Row && x.EndColumn - a.Column == y.EndColumn - b.Column;
                // A shared anchor such as a grand total is the same address.
                var anchored = x.Row == y.Row && x.Column == y.Column && x.EndRow == y.EndRow && x.EndColumn == y.EndColumn;
                if (!relative && !anchored) return false;
            }
            return true;
        }

        private static string CellText(IReadOnlyList<IReadOnlyList<string>> rows, int row, int column)
        {
            var i = row - DraftFirstRow; var j = column - 1;
            if (i < 0 || i >= rows.Count || rows[i] == null || j < 0 || j >= rows[i].Count) return "";
            return (rows[i][j] ?? "").Trim();
        }

        private static bool IsNumericOrFormula(string value)
        {
            if (value.Length == 0) return false;
            if (value[0] == '=') return true;
            // The writer keeps month keys such as 2026-06 as text labels.
            if (Regex.IsMatch(value, @"^\d{4}-(?:0[1-9]|1[0-2])$")) return false;
            decimal number; DateTime date;
            var plain = value.TrimEnd('%').TrimStart('$', '\u20AC', '\u00A3', '\u00A5', '\u20A9').Replace(",", "").Trim();
            return decimal.TryParse(plain, NumberStyles.Float, CultureInfo.InvariantCulture, out number) ||
                // Excel coerces date text to a serial; date arithmetic is valid.
                DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
        }

        private static string Shorten(string value) { return value.Length <= 40 ? value : value.Substring(0, 37) + "..."; }
        private static long Key(int row, int column) { return row * 100000L + column; }
        private static string Address(int row, int column) { return ColumnName(column) + row.ToString(CultureInfo.InvariantCulture); }
        private static int ColumnIndex(string letters)
        {
            var index = 0;
            foreach (var letter in letters.ToUpperInvariant()) index = index * 26 + (letter - 'A' + 1);
            return index;
        }
        private static Tuple<int, int> ParseAddress(string address)
        {
            var match = Regex.Match(address ?? "", @"^([A-Z]{1,3})([1-9][0-9]*)$", RegexOptions.IgnoreCase);
            return Tuple.Create(int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture), ColumnIndex(match.Groups[1].Value));
        }
        private static string ColumnName(int column)
        {
            var name = "";
            while (column > 0) { column--; name = (char)('A' + column % 26) + name; column /= 26; }
            return name;
        }
    }
}
