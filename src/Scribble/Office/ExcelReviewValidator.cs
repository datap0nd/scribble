using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace Scribble.Office
{
    public static class ExcelReviewValidator
    {
        // A reviewer must return source identities, not just an equally long list.
        public static IReadOnlyList<string> Validate(string json, IReadOnlyList<ExcelTaskCell> sources, out string terminology)
        {
            var result = new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.Deserialize<Dictionary<string, object>>(json);
            object raw;
            if (!result.TryGetValue("rows", out raw) || !(raw is IEnumerable)) throw new InvalidOperationException("Semantic review omitted source row identities.");
            var rows = ((IEnumerable)raw).Cast<object>().Select(value => value as IDictionary<string, object>).ToArray();
            if (rows.Length != sources.Count) throw new InvalidOperationException("Semantic review changed row coverage.");
            var values = new List<string>();
            for (var index = 0; index < rows.Length; index++)
            {
                object id, value;
                if (rows[index] == null || !rows[index].TryGetValue("id", out id) || Convert.ToString(id) != sources[index].Id ||
                    !rows[index].TryGetValue("value", out value) || !(value is string))
                    throw new InvalidOperationException("Semantic review changed source alignment at offset " + index);
                var text = (string)value;
                if (text.Length > ExcelSelectionOutputPolicy.MaxCellCharacters || (string.IsNullOrEmpty(sources[index].Value) && text.Length > 0))
                    throw new InvalidOperationException("Semantic review changed a blank row or exceeded Excel's cell length.");
                values.Add(text);
            }
            terminology = result.TryGetValue("terminology", out raw) ? Convert.ToString(raw) : "";
            return values;
        }
    }
}
