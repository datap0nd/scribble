using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace Scribble.Office
{
    internal static class SamsungDeckOverview
    {
        // A validation contact sheet of actual native renders, never slide artwork.
        internal static string Montage(IEnumerable<string> images)
        {
            var items = images.Take(12).ToArray();
            if (items.Length == 0) return null;
            using (var bitmap = new Bitmap(1200, ((items.Length + 2) / 3) * 225))
            using (var graphics = Graphics.FromImage(bitmap))
            using (var output = new MemoryStream())
            {
                graphics.Clear(Color.White);
                for (var i = 0; i < items.Length; i++)
                using (var input = new MemoryStream(Convert.FromBase64String(items[i].Substring(items[i].IndexOf(',') + 1))))
                using (var image = Image.FromStream(input)) graphics.DrawImage(image, (i % 3) * 400, (i / 3) * 225, 400, 225);
                bitmap.Save(output, ImageFormat.Png);
                return "data:image/png;base64," + Convert.ToBase64String(output.ToArray());
            }
        }
        internal static object Summary(Dictionary<string, object> slide)
        {
            var fields = new[] { "id", "title", "subtitle", "purpose", "layout", "takeaway", "unit", "footnote", "sources", "content_kind", "claims" };
            return slide.Where(p => fields.Contains(p.Key)).ToDictionary(p => p.Key, p => p.Value);
        }
    }
}
