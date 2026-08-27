using System;

namespace Scribble.Office
{
    // Maps the small model-facing chart vocabulary to XlChartType
    // codes shared by Excel and PowerPoint charts. The corporate
    // deck standard favors stacked and clustered columns, line
    // charts with markers, and 100% stacked columns for mix shifts;
    // unknown names fall back to a clustered column chart instead of
    // failing the draft. No 3-D type is reachable.
    public static class DraftChartTypes
    {
        public const int ColumnClustered = 51;
        public const int ColumnStacked = 52;
        public const int ColumnStacked100 = 53;
        public const int BarClustered = 57;
        public const int BarStacked = 58;
        public const int Line = 4;
        public const int LineMarkers = 65;
        public const int Pie = 5;
        public const int Area = 1;
        public const int Scatter = -4169;

        public static int Resolve(string name)
        {
            var kind = (name ?? string.Empty)
                .Trim()
                .ToLowerInvariant()
                .Replace("-", " ")
                .Replace("_", " ");
            switch (kind)
            {
                case "bar":
                    return BarClustered;
                case "stacked bar":
                case "bar stacked":
                    return BarStacked;
                case "stacked":
                case "stacked column":
                case "column stacked":
                    return ColumnStacked;
                case "100 stacked":
                case "100% stacked":
                case "stacked 100":
                case "column 100 stacked":
                case "100% stacked column":
                case "mix":
                    return ColumnStacked100;
                case "line":
                    return LineMarkers;
                case "line plain":
                case "line no markers":
                    return Line;
                case "pie":
                    return Pie;
                case "area":
                    return Area;
                case "scatter":
                    return Scatter;
                default:
                    return ColumnClustered;
            }
        }
    }
}
