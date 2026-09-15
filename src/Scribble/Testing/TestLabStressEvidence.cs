using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;

namespace Scribble.Testing
{
    // Native measurements of operator-owned output. Never supplied to the model.
    public static class TestLabStressEvidence
    {
        public static StressNative Read(object value, string host, string[] sourceSheets, int[] sourceSlides)
        {
            var result = new StressNative { host = host };
            var warnings = new List<string>();
            dynamic document = value;
            try
            {
                if (host == "Excel")
                {
                    var sheets = new List<StressSheet>();
                    for (int n = 1; n <= (int)document.Worksheets.Count; n++)
                    {
                        object rawSheet = document.Worksheets.Item(n); dynamic sheet = rawSheet;
                        try
                        {
                            string name = Convert.ToString(sheet.Name);
                            if ((sourceSheets ?? new string[0]).Contains(name)) continue;
                            // Calculate only this run's output worksheet, not all open workbooks.
                            sheet.Calculate();
                            var output = new StressSheet { name = name, recalculated = true };
                            object rawRange = sheet.UsedRange; dynamic range = rawRange;
                            try
                            {
                                int rows = (int)range.Rows.Count, columns = (int)range.Columns.Count;
                                if ((long)rows * columns > 10000) throw new InvalidOperationException("Output exceeds 10000-cell measurement limit.");
                                int startRow = (int)range.Row, startColumn = (int)range.Column;
                                object values = range.Value2, formulas = range.Formula;
                                var cells = new List<StressCell>();
                                for (int r = 0; r < rows; r++) for (int c = 0; c < columns; c++)
                                {
                                    object cell = At(values, r, c), formula = At(formulas, r, c);
                                    if (cell == null && formula == null) continue;
                                    cells.Add(new StressCell { row = startRow + r, column = startColumn + c,
                                        value = Scribble.Office.ExcelErrorValue.Text(cell) ?? cell,
                                        formula = Convert.ToString(formula, CultureInfo.InvariantCulture) });
                                }
                                output.cells = cells.ToArray();
                            }
                            finally { Release(rawRange); }
                            var charts = new List<StressChart>();
                            dynamic chartObjects = sheet.ChartObjects();
                            for (int c = 1; c <= Math.Min((int)chartObjects.Count, 32); c++)
                                charts.Add(Chart((object)chartObjects.Item(c).Chart));
                            output.charts = charts.ToArray(); sheets.Add(output);
                        }
                        finally { Release(rawSheet); }
                    }
                    result.sheets = sheets.ToArray();
                }
                else if (host == "PowerPoint")
                {
                    result.width = Convert.ToDouble(document.PageSetup.SlideWidth);
                    result.height = Convert.ToDouble(document.PageSetup.SlideHeight);
                    var slides = new List<StressSlide>();
                    for (int n = 1; n <= (int)document.Slides.Count; n++)
                    {
                        object rawSlide = document.Slides.Item(n); dynamic slide = rawSlide;
                        try
                        {
                            if ((sourceSlides ?? new int[0]).Contains((int)slide.SlideID)) continue;
                            if (slides.Count >= 60) throw new InvalidOperationException("Output exceeds 60-slide measurement limit.");
                            var measured = new StressSlide { number = n };
                            var shapes = new List<StressShape>(); var charts = new List<StressChart>();
                            if ((int)slide.Shapes.Count > 256) throw new InvalidOperationException("Output exceeds 256 shapes per slide.");
                            for (int j = 1; j <= (int)slide.Shapes.Count; j++)
                            {
                                object rawShape = slide.Shapes.Item(j); dynamic shape = rawShape;
                                try
                                {
                                    var s = Shape(rawShape);
                                    shapes.Add(s);
                                    if (!string.IsNullOrEmpty(s.measurement_warning)) warnings.Add("Slide " + n + " " + s.name + ": " + s.measurement_warning);
                                    if ((int)shape.Type == 6)
                                        warnings.Add("Slide " + n + " group " + s.name + ": child text and geometry require native visual review.");
                                    if ((int)shape.HasTable != 0)
                                    {
                                        object rawTable = shape.Table; dynamic table = rawTable;
                                        try
                                        {
                                            int rows = (int)table.Rows.Count, columns = (int)table.Columns.Count;
                                            if ((long)rows * columns > 512) throw new InvalidOperationException("Output exceeds 512 measured table cells per shape.");
                                            var cells = new List<StressShape>();
                                            for (int r = 1; r <= rows; r++) for (int c = 1; c <= columns; c++)
                                            {
                                                object rawCell = table.Cell(r, c); dynamic cell = rawCell; object rawCellShape = null;
                                                try
                                                {
                                                    rawCellShape = cell.Shape;
                                                    // PowerPoint cell adapters have no implemented
                                                    // Shape.Name getter. The table coordinates are
                                                    // already the stable native cell identity.
                                                    var measuredCell = Shape(rawCellShape, s.name + " R" + r + "C" + c);
                                                    measuredCell.is_table_cell = true;
                                                    if (!string.IsNullOrEmpty(measuredCell.measurement_warning)) warnings.Add("Slide " + n + " " + measuredCell.name + ": " + measuredCell.measurement_warning);
                                                    // Merged cells can expose one native rectangle
                                                    // at multiple row/column addresses.
                                                    if (!cells.Any(existing => Math.Abs(existing.x - measuredCell.x) < .1 && Math.Abs(existing.y - measuredCell.y) < .1 &&
                                                        Math.Abs(existing.width - measuredCell.width) < .1 && Math.Abs(existing.height - measuredCell.height) < .1)) cells.Add(measuredCell);
                                                }
                                                finally { Release(rawCellShape); Release(rawCell); }
                                            }
                                            shapes.AddRange(cells);
                                        }
                                        finally { Release(rawTable); }
                                    }
                                    if ((int)shape.HasChart != 0) charts.Add(Chart((object)shape.Chart));
                                }
                                finally { Release(rawShape); }
                            }
                            measured.shapes = shapes.ToArray(); measured.charts = charts.ToArray(); slides.Add(measured);
                        }
                        finally { Release(rawSlide); }
                    }
                    result.slides = slides.ToArray();
                }
                else if (host == "Word")
                {
                    var tables = new List<StressTable>();
                    if ((int)document.Tables.Count > 32) throw new InvalidOperationException("Output exceeds 32 measured Word tables.");
                    for (int n = 1; n <= (int)document.Tables.Count; n++)
                    {
                        object rawTable = document.Tables.Item(n); dynamic table = rawTable;
                        try
                        {
                            int rows = (int)table.Rows.Count, columns = (int)table.Columns.Count;
                            if ((long)rows * columns > 1024) throw new InvalidOperationException("Word table exceeds 1024 measured cells.");
                            var cells = new List<StressWordCell>();
                            for (int r = 1; r <= rows; r++) for (int c = 1; c <= columns; c++)
                            {
                                object rawCell = table.Cell(r, c); dynamic cell = rawCell; object rawRange = null;
                                try
                                {
                                    rawRange = cell.Range; dynamic range = rawRange;
                                    var valueText = Convert.ToString(range.Text) ?? "";
                                    cells.Add(new StressWordCell { row = r, column = c,
                                        text = valueText.TrimEnd('\r', '\a'), bold = Convert.ToInt32(range.Font.Bold) != 0 });
                                }
                                finally { Release(rawRange); Release(rawCell); }
                            }
                            bool borders = false;
                            try { borders = Convert.ToInt32(table.Borders.Enable) != 0; } catch { }
                            tables.Add(new StressTable { number = n, rows = rows, columns = columns, borders = borders, cells = cells.ToArray() });
                        }
                        finally { Release(rawTable); }
                    }
                    result.tables = tables.ToArray();
                }
            }
            catch (Exception error) { result.error = error.GetType().Name + ": " + error.Message; }
            result.measurement_warnings = warnings.Concat(result.sheets.SelectMany(s => s.charts).Concat(result.slides.SelectMany(s => s.charts))
                .SelectMany(c => c.measurement_warnings ?? new string[0])).Distinct().ToArray();
            return result;
        }

        private static StressShape Shape(object value, string knownName = null)
        {
            dynamic shape = value;
            var result = new StressShape { name = knownName ?? Convert.ToString(shape.Name),
                x = Convert.ToDouble(shape.Left), y = Convert.ToDouble(shape.Top),
                width = Convert.ToDouble(shape.Width), height = Convert.ToDouble(shape.Height) };
            object rawFill = null;
            try { rawFill = shape.Fill; ReadFill(rawFill, out result.fill_color, out result.measurement_warning); }
            catch (Exception error) { result.measurement_warning = "Fill measurement requires visual review: " + error.GetType().Name; }
            finally { Release(rawFill); }
            if ((int)shape.HasTextFrame == 0) return result;
            object rawFrame = shape.TextFrame; dynamic frame = rawFrame;
            try
            {
                if ((int)frame.HasText == 0) return result;
                object rawText = frame.TextRange; dynamic text = rawText;
                try
                {
                    result.text = Convert.ToString(text.Text); result.font = Convert.ToString(text.Font.Name);
                    result.font_size = Convert.ToDouble(text.Font.Size);
                    result.color = Color(Convert.ToInt32(text.Font.Color.RGB));
                    result.bound_width = Convert.ToDouble(text.BoundWidth); result.bound_height = Convert.ToDouble(text.BoundHeight);
                    result.available_width = result.width - Convert.ToDouble(frame.MarginLeft) - Convert.ToDouble(frame.MarginRight);
                    result.available_height = result.height - Convert.ToDouble(frame.MarginTop) - Convert.ToDouble(frame.MarginBottom);
                }
                finally { Release(rawText); }
            }
            finally { Release(rawFrame); }
            return result;
        }

        private static void ReadFill(object raw, out string color, out string warning)
        {
            color = null; warning = null;
            dynamic fill = raw;
            if ((int)fill.Visible == 0) return;
            if ((int)fill.Type != 1 || Convert.ToDouble(fill.Transparency) > .001)
            { warning = "Non-solid or translucent fill requires visual review."; return; }
            color = Color(Convert.ToInt32(fill.ForeColor.RGB));
        }

        private static object At(object raw, int row, int column)
        {
            var array = raw as Array;
            return array == null ? raw : array.GetValue(row + array.GetLowerBound(0), column + array.GetLowerBound(1));
        }
        private static StressChart Chart(object raw)
        {
            dynamic chart = raw;
            var result = new StressChart { type = Convert.ToInt32(chart.ChartType) };
            var warnings = new List<string>();
            try
            {
                result.title = (bool)chart.HasTitle ? Convert.ToString(chart.ChartTitle.Text) : "";
                var values = new List<StressSeries>(); dynamic series = chart.SeriesCollection();
                if ((int)series.Count > 32) throw new InvalidOperationException("Chart exceeds 32 series.");
                for (int n = 1; n <= (int)series.Count; n++)
                {
                    object rawSeries = series.Item(n); dynamic item = rawSeries;
                    try
                    {
                        var measured = new StressSeries { name = Convert.ToString(item.Name), categories = Scalars((object)item.XValues), values = Scalars((object)item.Values) };
                        object rawFormat = null, rawFill = null, rawLine = null;
                        try
                        {
                            rawFormat = item.Format; dynamic format = rawFormat;
                            rawFill = format.Fill; string warning;
                            ReadFill(rawFill, out measured.fill_color, out warning);
                            if (warning != null) warnings.Add("Chart series " + measured.name + ": " + warning);
                            rawLine = format.Line; dynamic line = rawLine;
                            if ((int)line.Visible != 0) measured.line_color = Color(Convert.ToInt32(line.ForeColor.RGB));
                        }
                        catch (Exception error) { warnings.Add("Chart series " + measured.name + " colors require visual review: " + error.GetType().Name); }
                        finally { Release(rawLine); Release(rawFill); Release(rawFormat); }
                        values.Add(measured);
                    }
                    finally { Release(rawSeries); }
                }
                result.series = values.ToArray();
                try { result.minimum = Convert.ToDouble(chart.Axes(2).MinimumScale); } catch { }
            }
            catch (Exception error) { result.error = error.Message; }
            finally { Release(raw); }
            result.measurement_warnings = warnings.Distinct().ToArray();
            return result;
        }
        private static string[] Scalars(object raw)
        {
            if (raw == null || Marshal.IsComObject(raw)) return new string[0];
            var array = raw as Array;
            var values = array == null ? new[] { raw } : array.Cast<object>().ToArray();
            if (values.Length > 128) throw new InvalidOperationException("Chart exceeds 128 points.");
            return values.Select(v => Scribble.Office.ExcelErrorValue.Text(v) ?? Convert.ToString(v, CultureInfo.InvariantCulture)).ToArray();
        }
        private static string Color(int rgb)
        { return "#" + (rgb & 255).ToString("X2") + ((rgb >> 8) & 255).ToString("X2") + ((rgb >> 16) & 255).ToString("X2"); }
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    }
    public sealed class StressNative
    { public string host, error; public double width, height; public string[] measurement_warnings = new string[0]; public StressSheet[] sheets = new StressSheet[0]; public StressSlide[] slides = new StressSlide[0]; public StressTable[] tables = new StressTable[0]; }
    public sealed class StressSheet
    { public string name; public bool recalculated; public StressCell[] cells = new StressCell[0]; public StressChart[] charts = new StressChart[0]; }
    public sealed class StressCell { public int row, column; public object value; public string formula; }
    public sealed class StressSlide { public int number; public StressShape[] shapes = new StressShape[0]; public StressChart[] charts = new StressChart[0]; }
    public sealed class StressShape
    { public string name, text, font, color, fill_color, measurement_warning; public bool is_table_cell; public double x, y, width, height, bound_width, bound_height, available_width, available_height, font_size; }
    public sealed class StressChart { public int type; public string title, error; public string[] measurement_warnings = new string[0]; public double? minimum; public StressSeries[] series = new StressSeries[0]; }
    public sealed class StressSeries { public string name, fill_color, line_color; public string[] categories, values; }
    public sealed class StressTable
    { public int number, rows, columns; public bool borders; public StressWordCell[] cells = new StressWordCell[0]; }
    public sealed class StressWordCell
    { public int row, column; public string text; public bool bold; }
}
