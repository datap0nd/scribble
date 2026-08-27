using System;
using Scribble.Security;

namespace Scribble.Chat
{
    public enum LocalSearchCommandKind
    {
        None,
        Help,
        Clear,
        Search
    }

    public sealed class LocalSearchCommand
    {
        private LocalSearchCommand(
            LocalSearchCommandKind kind,
            string query)
        {
            Kind = kind;
            Query = query ?? string.Empty;
        }

        public LocalSearchCommandKind Kind { get; }

        public string Query { get; }

        public static LocalSearchCommand Parse(string value)
        {
            var prompt = TextBoundary.PlainText(
                value,
                TextBoundary.MaxUserPromptCharacters);
            const string prefix = "/search";
            if (!prompt.StartsWith(
                    prefix,
                    StringComparison.OrdinalIgnoreCase) ||
                (prompt.Length > prefix.Length &&
                 !char.IsWhiteSpace(prompt[prefix.Length])))
            {
                return new LocalSearchCommand(
                    LocalSearchCommandKind.None,
                    string.Empty);
            }

            var argument = prompt
                .Substring(prefix.Length)
                .Trim();
            if (argument.Length == 0)
            {
                return new LocalSearchCommand(
                    LocalSearchCommandKind.Help,
                    string.Empty);
            }

            if (argument.Equals(
                "clear",
                StringComparison.OrdinalIgnoreCase))
            {
                return new LocalSearchCommand(
                    LocalSearchCommandKind.Clear,
                    string.Empty);
            }

            return new LocalSearchCommand(
                LocalSearchCommandKind.Search,
                TextBoundary.PlainText(argument, 240));
        }
    }
}
