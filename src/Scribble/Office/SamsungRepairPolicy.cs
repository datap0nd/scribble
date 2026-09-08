using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;
namespace Scribble.Office
{
    internal static class SamsungRepairPolicy
    {
        internal static object Canonical(object value)
        {
            var map = value as IDictionary<string, object>;
            if (map != null) return map.OrderBy(p => p.Key, StringComparer.Ordinal).ToDictionary(p => p.Key, p => Canonical(p.Value));
            var sequence = value as IEnumerable;
            if (sequence != null && !(value is string)) return sequence.Cast<object>().Select(Canonical).ToArray();
            return value;
        }
        internal static string Serialize(object value) { return new JavaScriptSerializer { MaxJsonLength = 16000000 }.Serialize(Canonical(value)); }
        internal static void ValidateScope(object[] original, object[] corrected)
        {
            if (corrected.Length != original.Length) throw new InvalidOperationException("REVISION_REPAIR_SCOPE: Preserve the exact operation count.");
            for (var i = 0; i < original.Length; i++)
            {
                var before = SamsungAuthoringPolicy.ReadMap(original[i]); var after = SamsungAuthoringPolicy.ReadMap(corrected[i]);
                var kind = SamsungAuthoringPolicy.Text(before, "kind");
                foreach (var key in before.Keys.Union(after.Keys))
                {
                    if ((kind == "replace_text" || kind == "table_cell") && key == "text")
                    {
                        Func<string, string[]> numbers = value => System.Text.RegularExpressions.Regex.Matches(value, @"[-+]?\d+(?:[.,]\d+)*%?").Cast<System.Text.RegularExpressions.Match>().Select(m => m.Value).OrderBy(v => v, StringComparer.Ordinal).ToArray();
                        if (!numbers(SamsungAuthoringPolicy.Text(before, key)).SequenceEqual(numbers(SamsungAuthoringPolicy.Text(after, key)))) throw new InvalidOperationException("REVISION_REPAIR_VALUES: Preserve the requested numeric changes.");
                        continue;
                    }
                    if ((kind == "replace_slide" || kind == "insert") && key == "slide")
                    {
                        var a = SamsungAuthoringPolicy.ReadMap(before[key]); var b = SamsungAuthoringPolicy.ReadMap(after[key]);
                        foreach (var evidence in new[] { "id", "table", "secondary_table", "chart", "secondary_chart", "image_names", "source_spans", "evidence", "calculations", "content_kind" })
                        { object x, y; a.TryGetValue(evidence, out x); b.TryGetValue(evidence, out y); if (Serialize(x) != Serialize(y)) throw new InvalidOperationException("REVISION_REPAIR_EVIDENCE: " + evidence); }
                        continue;
                    }
                    object xValue, yValue; before.TryGetValue(key, out xValue); after.TryGetValue(key, out yValue);
                    if (Serialize(xValue) != Serialize(yValue)) throw new InvalidOperationException("REVISION_REPAIR_SCOPE: " + key);
                }
            }
        }
    }
}
