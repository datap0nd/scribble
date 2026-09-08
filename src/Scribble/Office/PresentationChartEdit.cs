using System;
using System.Collections;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Scribble.Office
{
    internal static class PresentationChartEdit
    {
        internal sealed class RangeTarget
        {
            internal string Sheet, Cell;
        }
        internal static RangeTarget Resolve(string formula, int category)
        {
            // Only a direct local SERIES value range is writable. Computed,
            // external, named, disjoint and multidimensional ranges stay intact.
            if (category < 1 || formula == null || !formula.StartsWith("=SERIES(", StringComparison.OrdinalIgnoreCase) || !formula.EndsWith(")"))
                throw new InvalidOperationException("REVISION_CHART_SOURCE_UNSUPPORTED: Expected a direct local series range.");
            var fields = new System.Collections.Generic.List<string>(); var start = 8; var quoted = false;
            for (var i = start; i < formula.Length - 1; i++)
            {
                if (formula[i] == '\'')
                { if (quoted && i + 1 < formula.Length && formula[i + 1] == '\'') { i++; continue; } quoted = !quoted; }
                if (!quoted && formula[i] == ',') { fields.Add(formula.Substring(start, i - start)); start = i + 1; }
            }
            fields.Add(formula.Substring(start, formula.Length - 1 - start));
            if (quoted || fields.Count != 4) throw new InvalidOperationException("REVISION_CHART_SOURCE_UNSUPPORTED: Complex series formulas cannot be edited safely.");
            var match = Regex.Match(fields[2].Trim(), @"^(?:'(?<sheet>(?:[^']|'')+)'|(?<sheet>[\p{L}_][\p{L}\p{N}_ ]*))!\$(?<col>[A-Z]{1,3})\$(?<row>[1-9][0-9]*)(?::\$(?<endcol>[A-Z]{1,3})\$(?<endrow>[1-9][0-9]*))?$", RegexOptions.IgnoreCase);
            if (!match.Success || match.Groups["sheet"].Value.IndexOfAny(new[] { '[', ']', ':', '/', '\\' }) >= 0)
                throw new InvalidOperationException("REVISION_CHART_SOURCE_UNSUPPORTED: Only a local rectangular range is supported.");
            var col = match.Groups["col"].Value.ToUpperInvariant(); var endCol = match.Groups["endcol"].Success ? match.Groups["endcol"].Value.ToUpperInvariant() : col;
            var row = int.Parse(match.Groups["row"].Value, CultureInfo.InvariantCulture); var endRow = match.Groups["endrow"].Success ? int.Parse(match.Groups["endrow"].Value, CultureInfo.InvariantCulture) : row;
            var column = Column(col); var lastColumn = Column(endCol);
            if (endRow > 1048576 || lastColumn > 16384 || endRow < row || lastColumn < column || (col != endCol && row != endRow) || category > Math.Max(endRow - row, lastColumn - column) + 1)
                throw new InvalidOperationException("REVISION_CHART_SOURCE_UNSUPPORTED: Category must identify one cell in a single row or column.");
            return new RangeTarget { Sheet = match.Groups["sheet"].Value.Replace("''", "'"), Cell = Letter(column + (row == endRow ? category - 1 : 0)) + (row + (col == endCol ? category - 1 : 0)).ToString(CultureInfo.InvariantCulture) };
        }
        private static int Column(string text) { var value = 0; foreach (var c in text) value = value * 26 + c - 'A' + 1; return value; }
        private static string Letter(int column) { var result = ""; while (column > 0) { column--; result = (char)('A' + column % 26) + result; column /= 26; } return result; }
        internal static void SetPoint(object value, int seriesIndex, int category, double expected, double replacement)
        {
            dynamic chart = value;
            if ((bool)chart.ChartData.IsLinked || double.IsNaN(replacement) || double.IsInfinity(replacement)) throw new InvalidOperationException("REVISION_CHART_UNSUPPORTED");
            dynamic series = chart.SeriesCollection(seriesIndex);
            var formula = Convert.ToString(series.Formula); var target = Resolve(formula, category);
            var values = ((IEnumerable)series.Values).Cast<object>().ToArray();
            if (category > values.Length || values[category - 1] == null || Convert.ToDouble(values[category - 1], CultureInfo.InvariantCulture) != expected) throw new InvalidOperationException("REVISION_CHART_CHANGED");
            chart.ChartData.Activate(); dynamic dataWorkbook = chart.ChartData.Workbook;
            var changed = false;
            try
            {
                dynamic cell = dataWorkbook.Worksheets[target.Sheet].Range[target.Cell];
                if ((bool)cell.HasFormula || cell.Value2 == null || Convert.ToDouble(cell.Value2, CultureInfo.InvariantCulture) != expected)
                    throw new InvalidOperationException("REVISION_CHART_SOURCE_CHANGED: The embedded source cell must match the inspected point and contain a constant.");
                cell.Value2 = replacement; changed = true;
                if (Convert.ToDouble(cell.Value2, CultureInfo.InvariantCulture) != replacement) throw new InvalidOperationException("REVISION_CHART_SOURCE_WRITE_FAILED");
            }
            finally { dataWorkbook.Close(changed); }
            chart.Refresh();
            values = ((IEnumerable)series.Values).Cast<object>().ToArray();
            if (Convert.ToString(series.Formula) != formula || Convert.ToDouble(values[category - 1], CultureInfo.InvariantCulture) != replacement)
                throw new InvalidOperationException("REVISION_CHART_ASSOCIATION_CHANGED: The chart did not retain the reviewed workbook association.");
        }
    }
}
