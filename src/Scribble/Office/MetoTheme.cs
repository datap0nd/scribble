using System;

namespace Scribble.Office
{
    // The hardcoded METO corporate deck theme. Draft slides are
    // painted entirely from these tokens by PresentationDraftWriter:
    // the model supplies content only - never fonts, colors, sizes,
    // or positions - so even a small local model produces slides
    // that match the corporate executive standard. Nothing here is
    // model-adjustable; the tokens are compiled in.
    public static class MetoTheme
    {
        public const string ThemeName = "Samsung MD 1.0";

        // --- Typography -------------------------------------------
        // Brand families first; Office falls back automatically on
        // machines without the Samsung Sharp Sans family installed,
        // which is why the body and label fonts stay universal.
        public const string TitleFont = "Samsung Sharp Sans Bold";
        public const string SubtitleFont = "Arial";
        public const string BodyFont = "Arial";
        public const string LabelFont = "Arial Narrow";
        public const string FootnoteFont = "Arial Narrow";

        // --- Color tokens (hex; converted to the BGR longs the
        // Office object model expects) ------------------------------
        public const string BrandBlueHex = "#4F81BD";
        public const string TextBlackHex = "#000000";
        public const string CharcoalHex = "#404040";
        public const string GrayHex = "#7F7F7F";
        public const string LightGrayHex = "#A6A6A6";
        public const string WhiteHex = "#FFFFFF";
        public const string CardSlateHex = "#F2F2F2";
        public const string GoodGreenHex = "#E2EFDA";
        public const string BadYellowHex = "#FFF2CC";
        public const string WarningAmberHex = "#FFC000";
        public const string GrowthBlueHex = "#0000FF";
        public const string GridlineHex = "#D9D9D9";

        // --- Font sizes (pt) --------------------------------------
        public const float CoverTitleSize = 60f;
        public const float CoverSubtitleSize = 20f;
        public const float SlideTitleSize = 40f;
        public const float SectionHeaderSize = 24f;
        public const float SubtitleSize = 15f;
        public const float CardHeaderSize = 13f;
        public const float BodyBulletSize = 11f;
        public const float MetricSize = 16f;
        public const float TableHeaderSize = 11f;
        public const float TableBodySize = 10f;
        public const float ChartLabelSize = 9f;
        public const float ChartTitleSize = 10f;
        public const float FootnoteSize = 7f;

        // --- Layout templates -------------------------------------
        public const string LayoutCover = "cover";
        public const string LayoutAgenda = "agenda";
        public const string LayoutCards = "cards";
        public const string LayoutTable = "table";
        public const string LayoutChart = "chart";
        public const string LayoutBullets = "bullets";

        // --- Corporate growth markers (kept as escapes so every
        // source file of the add-in stays pure ASCII) --------------
        public const string UpArrow = "\u2191";
        public const string DownArrow = "\u2193";
        public const string RightArrow = "\u2192";
        public const string DeficitTriangle = "\u25B3";

        // Circled numbers head each card of the strategy grid.
        public static readonly string[] CircledNumbers =
            new[]
            {
                "\u2460",
                "\u2461",
                "\u2462",
                "\u2463"
            };

        // --- Selective status highlighting ------------------------
        public const int StatusNone = 0;
        public const int StatusGood = 1;
        public const int StatusBad = 2;

        // The guidelines cap highlighting at three or four cells per
        // status per slide so a table stays low-noise.
        public const int MaxHighlightsPerStatus = 4;

        // Converts "#RRGGBB" into the BGR long the Office object
        // model uses for RGB properties. Unparsable input falls back
        // to black rather than failing a draft.
        public static int Rgb(string hex)
        {
            var text = (hex ?? string.Empty).Trim();
            if (text.StartsWith("#", StringComparison.Ordinal))
            {
                text = text.Substring(1);
            }

            if (text.Length != 6)
            {
                return 0;
            }

            try
            {
                var red = Convert.ToInt32(
                    text.Substring(0, 2),
                    16);
                var green = Convert.ToInt32(
                    text.Substring(2, 2),
                    16);
                var blue = Convert.ToInt32(
                    text.Substring(4, 2),
                    16);
                return red | (green << 8) | (blue << 16);
            }
            catch
            {
                return 0;
            }
        }

        // Chart series palette: brand blue leads, then neutral greys,
        // with amber reserved for the attention series. No chart ever
        // uses more colors than it has series.
        public static int[] ChartSeriesColors()
        {
            return new[]
            {
                Rgb(BrandBlueHex),
                Rgb("#5B9BD5"),
                Rgb("#F2F2F2"),
                Rgb(WarningAmberHex),
                Rgb(CharcoalHex)
            };
        }

        // Picks the layout template for one slide. An explicit
        // request wins for the cover and agenda pages; everything
        // else is inferred from the content the model supplied, so a
        // small model never has to name a layout at all.
        public static string ResolveLayout(
            string requested,
            bool hasBullets,
            bool hasCards,
            bool hasTable,
            bool hasChart)
        {
            var kind = (requested ?? string.Empty)
                .Trim()
                .ToLowerInvariant();
            if (Array.IndexOf(SamsungSlideDesign.Layouts, kind) >= 0) return kind;
            if (kind == "cover" ||
                kind == "title" ||
                kind == "cover_slide")
            {
                return LayoutCover;
            }

            if (kind == "agenda" ||
                kind == "section" ||
                kind == "divider" ||
                kind == "contents")
            {
                return LayoutAgenda;
            }

            if (hasCards)
            {
                return LayoutCards;
            }

            if (hasTable)
            {
                return LayoutTable;
            }

            if (hasChart && !hasBullets)
            {
                return LayoutChart;
            }

            return LayoutBullets;
        }

        // Classifies one table cell for selective highlighting. The
        // corporate vocabulary already marks performance with arrows
        // and the deficit triangle, so the theme can color cells
        // without the model choosing any color: growth reads green,
        // shortfall reads yellow, everything else stays white.
        public static int CellStatus(string text)
        {
            var value = (text ?? string.Empty).Trim();
            if (value.Length == 0)
            {
                return StatusNone;
            }

            if (value.IndexOf(
                    DeficitTriangle,
                    StringComparison.Ordinal) >= 0 ||
                value.IndexOf(
                    DownArrow,
                    StringComparison.Ordinal) >= 0)
            {
                return StatusBad;
            }

            if (value.IndexOf(
                UpArrow,
                StringComparison.Ordinal) >= 0)
            {
                return StatusGood;
            }

            // A signed number is the other unambiguous marker:
            // "+12%" is growth, "-8%" is a shortfall. Bare numbers
            // and text are never highlighted.
            if (value[0] == '+' && HasDigit(value))
            {
                return StatusGood;
            }

            if ((value[0] == '-' || value[0] == '\u2212') &&
                HasDigit(value))
            {
                return StatusBad;
            }

            return StatusNone;
        }

        // Steps a headline down when the text is long so a real
        // title never overflows its band. The token size is the
        // ceiling, never a fixed value.
        public static float FitTitleSize(
            string text,
            float baseSize,
            int comfortableLength)
        {
            var length = (text ?? string.Empty).Length;
            if (comfortableLength < 1 ||
                length <= comfortableLength)
            {
                return baseSize;
            }

            var scaled = baseSize *
                ((float)comfortableLength / length);
            var floor = baseSize * 0.6f;
            return scaled < floor ? floor : scaled;
        }

        private static bool HasDigit(string value)
        {
            foreach (var character in value)
            {
                if (character >= '0' && character <= '9')
                {
                    return true;
                }
            }

            return false;
        }
    }
}
