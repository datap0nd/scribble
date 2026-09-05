using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;

namespace Scribble.Office
{
    // Versioned, host-owned geometry. Percentages come from the supplied MD;
    // conflicting/off-canvas observations are normalized in the implementation notes.
    public static class SamsungSlideDesign
    {
        public const string Version = "Samsung MD 1.0";
        public const float Width = 960f, Height = 540f;
        public const string Blue = "#4F81BD", SoftBlue = "#5B9BD5", Border = "#41719C";
        public const string Gray = "#F2F2F2", Red = "#C00000", Green = "#00B050";
        public static readonly string[] Layouts = { "cover", "divider", "agenda", "bullets", "cards", "action_list",
            "two_pane", "table", "matrix", "chart", "annotated_chart", "visual_grid", "dual_visual",
            "large_visual", "landscape", "centered_visual", "visual_comments", "roadmap", "stack", "closing" };
        public static RectangleF Percent(float x, float y, float w, float h)
        { return new RectangleF(x * Width / 100, y * Height / 100, w * Width / 100, h * Height / 100); }
        public static RectangleF Title { get { return Percent(6.1f, 5.2f, 90f, 9.6f); } }
        public static RectangleF Action { get { return Percent(3.8f, 16.1f, 92.3f, 4.5f); } }
        public static RectangleF Footer { get { return Percent(3.8f, 93.5f, 87f, 3f); } }
        public static RectangleF Page { get { return Percent(94f, 96.8f, 6f, 3.2f); } }
        public static RectangleF Takeaway { get { return Percent(16.5f, 85.1f, 60.6f, 7.6f); } }
        public static RectangleF[] Regions(string layout)
        {
            switch (layout)
            {
                case "cover": return new[] { Percent(4.6f, 28.2f, 84.4f, 34.1f) };
                case "divider": return new[] { Percent(24.1f, 36.8f, 71.9f, 48f) };
                case "closing": return new[] { Percent(7.8f, 18.6f, 63.8f, 68.7f) };
                case "two_pane": return new[] { Percent(3.8f, 25.1f, 59f, 58f), Percent(67f, 28.7f, 29.2f, 24f), Percent(67f, 60f, 29.2f, 23f) };
                case "table": case "matrix": return new[] { Percent(16.5f, 32.1f, 60.7f, 52.1f) };
                case "annotated_chart": return new[] { Percent(11.8f, 34f, 41.9f, 46f), Percent(55.8f, 38f, 36.6f, 42f) };
                case "dual_visual": return new[] { Percent(3.8f, 25f, 43f, 56f), Percent(51.1f, 25f, 45.1f, 56f) };
                case "visual_comments": return new[] { Percent(3.8f, 25f, 53.3f, 56f), Percent(60f, 25f, 36.2f, 56f) };
                case "visual_grid": return new[] { Percent(3.8f, 24f, 43f, 27f), Percent(51.1f, 24f, 45.1f, 27f), Percent(3.8f, 55f, 43f, 27f), Percent(51.1f, 55f, 45.1f, 27f) };
                case "centered_visual": return new[] { Percent(13.5f, 31.9f, 61.1f, 46.8f) };
                case "roadmap": return new[] { Percent(5.2f, 25f, 78.1f, 57f) };
                case "stack": return new[] { Percent(17.6f, 32.1f, 59.4f, 50f) };
                default: return new[] { Percent(3.8f, 25f, 92.4f, 57f) };
            }
        }
        public static bool InBounds(RectangleF rectangle)
        { return rectangle.Width > 0 && rectangle.Height > 0 && rectangle.Left >= 0 && rectangle.Top >= 0 && rectangle.Right <= Width + .01 && rectangle.Bottom <= Height + .01; }
        public static string FontFor(string text, string preferred)
        {
            if ((text ?? "").Any(c => (c >= '\uAC00' && c <= '\uD7AF') || (c >= '\u1100' && c <= '\u11FF'))) return "Malgun Gothic";
            using (var installed = new InstalledFontCollection())
                return installed.Families.Any(f => f.Name.Equals(preferred, StringComparison.OrdinalIgnoreCase)) ? preferred : "Arial";
        }
        // GDI is the deterministic preflight; native TextRange bounds are checked
        // again after PowerPoint renders, since Office font metrics can differ.
        public static float Fit(string text, string font, RectangleF area, float maximum, float minimum, bool bold = false)
        {
            using (var bitmap = new Bitmap(1, 1))
            using (var graphics = Graphics.FromImage(bitmap))
            using (var format = new StringFormat(StringFormat.GenericTypographic))
            {
                graphics.PageUnit = GraphicsUnit.Point;
                format.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces;
                for (var size = maximum; size >= minimum; size -= .5f)
                using (var type = new Font(FontFor(text, font), size, bold ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Point))
                {
                    var measured = graphics.MeasureString(text ?? "", type, new SizeF(area.Width - 4, 10000), format);
                    if (measured.Height <= area.Height - 3 && measured.Width <= area.Width) return size;
                }
            }
            throw new InvalidOperationException("SLIDE_OVERFLOW: " + (text ?? "").Substring(0, Math.Min(80, (text ?? "").Length)) +
                " [" + font + ", minimum " + minimum + "pt, box " + area.Width + " x " + area.Height + "]. Split this content across slides; no source text was dropped.");
        }
        public static int RowsPerPage(int columns, float height = 281.34f)
        { return Math.Max(1, (int)(height / 19f) - 1); }
    }
}
