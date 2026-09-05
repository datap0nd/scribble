using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace Scribble.Chat
{
    // Interpret the result envelope, never words inside an email, page or document.
    public sealed class ToolOutcome
    {
        public bool Failed { get; private set; }
        public string ErrorCode { get; private set; }
        public bool? PermissionConsumed { get; private set; }
        public string Stage { get; private set; }

        public static ToolOutcome Parse(string content)
        {
            var result = new ToolOutcome();
            var text = (content ?? "").TrimStart();
            if (text.StartsWith("{"))
            {
                try
                {
                    var map = new JavaScriptSerializer { MaxJsonLength = int.MaxValue }
                        .Deserialize<Dictionary<string, object>>(text);
                    object value;
                    if (map.TryGetValue("error_code", out value) && value is string && !string.IsNullOrWhiteSpace((string)value))
                    { result.Failed = true; result.ErrorCode = (string)value; }
                    if (map.TryGetValue("ok", out value) && value is bool && !(bool)value) result.Failed = true;
                    if (map.TryGetValue("permission_consumed", out value) && value is bool) result.PermissionConsumed = (bool)value;
                    if (map.TryGetValue("stage", out value)) result.Stage = Convert.ToString(value);
                }
                catch (ArgumentException) { }
            }
            else
            {
                // Compatibility for existing host error envelopes only.
                var match = Regex.Match(text, @"^\[((?:BROWSER|WEB_FETCH|TASK|DRAFT|DOCUMENT|SLIDE|CROSS_APP|MAILBOX|MCP)_[A-Z0-9_]+)\]");
                if (match.Success) { result.Failed = true; result.ErrorCode = match.Groups[1].Value; }
            }
            return result;
        }
    }
}
