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
                // Known compatibility case only: decode one encoded slide/plan array.
                if (map != null && (call.function.name == "add_draft_slides" || call.function.name == "send_to_powerpoint"))
                    foreach (var key in new[] { "slides", "plan" })
                    {
                        object raw;
                        if (map.TryGetValue(key, out raw) && raw is string && ((string)raw).TrimStart().StartsWith("["))
                        {
                            var decoded = json.DeserializeObject((string)raw);
                            if (decoded is IList) map[key] = decoded;
                        }
                    }
                Visit(args, schema, "$", errors);
                if (errors.Count == 0) call.function.arguments = json.Serialize(args);
            }
            catch (ArgumentException) { errors.Add("$: arguments must be valid JSON matching the tool schema."); }
            return errors;
        }

        private static void Visit(object value, IDictionary<string, object> schema, string path, List<string> errors)
        {
            if (schema == null || errors.Count >= 12) return;
            object raw;
            var type = schema.TryGetValue("type", out raw) ? Convert.ToString(raw) : "";
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
