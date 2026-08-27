using System;

namespace Scribble.UI
{
    public static class SubjectDisplay
    {
        public static string Clean(string subject)
        {
            var value = (subject ?? string.Empty).Trim();
            while (true)
            {
                var previous = value;
                value = RemovePrefix(value, "RE:");
                value = RemovePrefix(value, "FW:");
                value = RemovePrefix(value, "FWD:");
                if (string.Equals(
                    value,
                    previous,
                    StringComparison.Ordinal))
                {
                    return value;
                }
            }
        }

        private static string RemovePrefix(
            string value,
            string prefix)
        {
            return value.StartsWith(
                prefix,
                StringComparison.OrdinalIgnoreCase)
                ? value.Substring(prefix.Length).TrimStart()
                : value;
        }
    }
}
