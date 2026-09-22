using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web.Script.Serialization;

namespace Scribble.Chat
{
    // Model definitions and host validation consume the same schema.
    public static class ToolContractValidator
    {
        public static IReadOnlyList<string> Validate(ChatToolCall call, ChatToolDefinition definition)
        {
            var errors = new List<string>();
            var json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            try
            {
                var args = json.DeserializeObject(call.function.arguments ?? "{}");
                var schema = json.DeserializeObject(json.Serialize(definition.function.parameters)) as IDictionary<string, object>;
                var map = args as IDictionary<string, object>;
                // This read-only inventory has no parameters. Some compatible
                // providers invent conventional paging/visibility hints, or
                // copy sheet/range arguments from read_cells, even when this
                // schema is {}. The inventory ignores all of these fields;
                // discard only known inert hints so a repeated malformed
                // call cannot strand a cross-app task.
                if (map != null && call.function.name == WorkbookToolCatalog.ListWorksheets)
                {
                    map.Remove("limit");
                    map.Remove("include_hidden");
                    map.Remove("sheet");
                    map.Remove("range");
                    map.Remove("rows");
                    map.Remove("columns");
                    map.Remove("run_in_background");
                }
                // Known compatibility case only: decode one encoded slide/plan array.
                // Some OpenAI-compatible gateways preserve a model's nested JSON
                // array as a string. Qwen can also append one structurally misplaced
                // optional field after otherwise complete slide objects. Retain only
                // independently valid, complete slide objects from that array prefix;
                // the accepted deck plan makes the model continue with any missing
                // slides in a later call. Never attempt general JSON repair.
                if (map != null && (call.function.name == "add_draft_slides" || call.function.name == "send_to_powerpoint"))
                    foreach (var key in new[] { "slides", "plan" })
                    {
                        object raw;
                        if (map.TryGetValue(key, out raw) && raw is string && ((string)raw).TrimStart().StartsWith("["))
                        {
                            try
                            {
                                var decoded = json.DeserializeObject((string)raw);
                                if (decoded is IList) map[key] = decoded;
                            }
                            catch (ArgumentException)
                            {
                                IList decodedPrefix;
                                if (key == "slides" &&
                                    (TryCloseOneTableRowsArray((string)raw, json, out decodedPrefix) ||
                                     TryDecodeCompleteObjectArrayPrefix((string)raw, json, out decodedPrefix)))
                                {
                                    map[key] = decodedPrefix;
                                }
                                else
                                {
                                    throw;
                                }
                            }
                        }
                    }
                Visit(args, schema, "$", errors);
                if (errors.Count == 0) call.function.arguments = json.Serialize(args);
            }
            catch (ArgumentException) { errors.Add("$: arguments must be valid JSON matching the tool schema."); }
            return errors;
        }

        // Qwen sometimes quotes the whole slides array and omits exactly the
        // closing bracket of a table's rows, while every row and cell is intact.
        // Repair this one structural typo only; schema, source and visual gates
        // still run on the complete decoded slide. Never synthesize content.
        private static bool TryCloseOneTableRowsArray(string raw, JavaScriptSerializer json, out IList decoded)
        {
            decoded = null;
            if (string.IsNullOrWhiteSpace(raw) ||
                (raw.IndexOf("\"table\"", StringComparison.Ordinal) < 0 &&
                 raw.IndexOf("\"secondary_table\"", StringComparison.Ordinal) < 0)) return false;
            var rows = raw.IndexOf("\"rows\"", StringComparison.Ordinal);
            if (rows < 0) return false;
            var start = raw.IndexOf('[', rows + 6);
            if (start < 0) return false;
            var depth = 0;
            var inString = false;
            var escaped = false;
            for (var i = start; i < raw.Length; i++)
            {
                var c = raw[i];
                if (inString)
                {
                    if (escaped) escaped = false;
                    else if (c == '\\') escaped = true;
                    else if (c == '"') inString = false;
                    continue;
                }
                if (c == '"') { inString = true; continue; }
                if (c == '[') { depth++; continue; }
                if (c == ']') { if (--depth == 0) return false; continue; }
                if (c != '}' || depth != 1 || i == 0 || raw[i - 1] != ']') continue;
                try
                {
                    var fixedArray = json.DeserializeObject(raw.Insert(i, "]")) as IList;
                    if (fixedArray == null || fixedArray.Count == 0 ||
                        fixedArray.Cast<object>().Any(item => !(item is IDictionary<string, object>))) return false;
                    decoded = fixedArray;
                    return true;
                }
                catch (ArgumentException) { return false; }
            }
            return false;
        }

        private static bool TryDecodeCompleteObjectArrayPrefix(
            string raw,
            JavaScriptSerializer json,
            out IList decoded)
        {
            var items = new ArrayList();
            decoded = items;
            if (string.IsNullOrWhiteSpace(raw)) return false;

            var cursor = 0;
            while (cursor < raw.Length && char.IsWhiteSpace(raw[cursor])) cursor++;
            if (cursor >= raw.Length || raw[cursor] != '[') return false;
            cursor++;

            while (cursor < raw.Length)
            {
                while (cursor < raw.Length &&
                       (char.IsWhiteSpace(raw[cursor]) || raw[cursor] == ','))
                    cursor++;
                if (cursor >= raw.Length || raw[cursor] == ']') break;
                if (raw[cursor] != '{') break;

                var start = cursor;
                var depth = 0;
                var inString = false;
                var escaped = false;
                var complete = false;
                for (; cursor < raw.Length; cursor++)
                {
                    var character = raw[cursor];
                    if (inString)
                    {
                        if (escaped) escaped = false;
                        else if (character == '\\') escaped = true;
                        else if (character == '"') inString = false;
                        continue;
                    }

                    if (character == '"') inString = true;
                    else if (character == '{') depth++;
                    else if (character == '}' && --depth == 0)
                    {
                        complete = true;
                        break;
                    }
                }

                if (!complete) break;
                object item;
                try
                {
                    item = json.DeserializeObject(
                        raw.Substring(start, cursor - start + 1));
                }
                catch (ArgumentException)
                {
                    break;
                }

                if (!(item is IDictionary<string, object>)) break;
                items.Add(item);
                cursor++;
            }

            return items.Count > 0;
        }

        private static void Visit(object value, IDictionary<string, object> schema, string path, List<string> errors)
        {
            if (schema == null || errors.Count >= 12) return;
            object raw;
            if (schema.TryGetValue("type", out raw) && raw is IList)
            {
                foreach (var option in ((IList)raw).Cast<object>().Select(Convert.ToString))
                {
                    var candidate = new Dictionary<string, object>(schema) { ["type"] = option };
                    var candidateErrors = new List<string>(); Visit(value, candidate, path, candidateErrors);
                    if (candidateErrors.Count == 0) return;
                }
                errors.Add(path + ": value does not match an allowed type."); return;
            }
            var type = schema.TryGetValue("type", out raw) ? Convert.ToString(raw) : "";
            if (type == "null") { if (value != null) errors.Add(path + ": must be null."); return; }
            if (schema.TryGetValue("enum", out raw) && raw is IEnumerable && !((IEnumerable)raw).Cast<object>().Any(v => Equals(v, value)))
                errors.Add(path + ": value is not one of the allowed choices.");
            if (type == "object")
            {
                var map = value as IDictionary<string, object>;
                if (map == null) { errors.Add(path + ": must be an object."); return; }
                if (schema.TryGetValue("required", out raw) && raw is IEnumerable)
                    foreach (var key in ((IEnumerable)raw).Cast<object>().Select(Convert.ToString))
                        if (!map.ContainsKey(key)) errors.Add(path + "." + key + ": required field missing.");
                var properties = schema.TryGetValue("properties", out raw) ? raw as IDictionary<string, object> : null;
                var forbidExtra = schema.TryGetValue("additionalProperties", out raw) && raw is bool && !(bool)raw;
                foreach (var pair in map)
                {
                    object child;
                    if (properties != null && properties.TryGetValue(pair.Key, out child)) Visit(pair.Value, child as IDictionary<string, object>, path + "." + pair.Key, errors);
                    else if (forbidExtra) errors.Add(path + "." + pair.Key + ": field is not supported.");
                }
            }
            else if (type == "array")
            {
                var list = value as IList;
                if (list == null) { errors.Add(path + ": must be an array."); return; }
                if (schema.TryGetValue("minItems", out raw) && list.Count < Convert.ToInt32(raw)) errors.Add(path + ": too few items.");
                if (schema.TryGetValue("maxItems", out raw) && list.Count > Convert.ToInt32(raw)) errors.Add(path + ": too many items; split the batch.");
                if (schema.TryGetValue("items", out raw))
                    for (var i = 0; i < list.Count && errors.Count < 12; i++) Visit(list[i], raw as IDictionary<string, object>, path + "[" + i + "]", errors);
            }
            else if (type == "string" && !(value is string)) errors.Add(path + ": must be text.");
            else if (type == "boolean" && !(value is bool)) errors.Add(path + ": must be true or false, not quoted text.");
            else if (type == "integer" || type == "number")
            {
                double number;
                if (value == null || value is string || value is bool || !double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out number) ||
                    double.IsNaN(number) || double.IsInfinity(number) || (type == "integer" && Math.Truncate(number) != number))
                { errors.Add(path + ": must be a finite " + type + "."); return; }
                if (schema.TryGetValue("minimum", out raw) && number < Convert.ToDouble(raw)) errors.Add(path + ": below the allowed minimum.");
                if (schema.TryGetValue("maximum", out raw) && number > Convert.ToDouble(raw)) errors.Add(path + ": above the allowed maximum.");
            }
        }
    }
}
