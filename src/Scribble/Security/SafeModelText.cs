using System;
using System.Collections.Generic;
using System.Text;

namespace Scribble.Security
{
    public static class SafeModelText
    {
        public static FormattedModelText Format(
            string value,
            int maximumCharacters)
        {
            var text = TextBoundary.PlainText(
                value,
                maximumCharacters);
            var output = new StringBuilder(text.Length);
            var boldRanges = new List<TextSpan>();
            var position = 0;
            var lineStart = true;
            while (position < text.Length)
            {
                if (lineStart &&
                    position + 1 < text.Length &&
                    text[position] == '*' &&
                    text[position + 1] == ' ')
                {
                    output.Append("- ");
                    position += 2;
                    lineStart = false;
                    continue;
                }

                var markerLength = StrongMarkerLengthAt(
                    text,
                    position);
                if (markerLength > 0)
                {
                    var marker = text.Substring(
                        position,
                        markerLength);
                    var close = text.IndexOf(
                        marker,
                        position + markerLength,
                        StringComparison.Ordinal);
                    if (close > position + markerLength)
                    {
                        var content = text.Substring(
                            position + markerLength,
                            close - position - markerLength);
                        var start = output.Length;
                        output.Append(content);
                        boldRanges.Add(
                            new TextSpan(start, content.Length));
                        lineStart = EndsLine(content);
                        position = close + markerLength;
                        continue;
                    }
                }

                var character = text[position];
                if (character == '*' &&
                    !IsBetweenLettersOrDigits(text, position))
                {
                    position++;
                    continue;
                }

                output.Append(character);
                lineStart = character == '\r' || character == '\n';
                position++;
            }

            return new FormattedModelText(
                output.ToString(),
                boldRanges);
        }

        private static int StrongMarkerLengthAt(
            string text,
            int position)
        {
            if (position + 2 < text.Length &&
                text[position] == '*' &&
                text[position + 1] == '*' &&
                text[position + 2] == '*' &&
                position + 3 < text.Length &&
                !char.IsWhiteSpace(text[position + 3]))
            {
                return 3;
            }

            if (position + 1 < text.Length &&
                ((text[position] == '*' &&
                  text[position + 1] == '*') ||
                 (text[position] == '_' &&
                  text[position + 1] == '_')) &&
                position + 2 < text.Length &&
                !char.IsWhiteSpace(text[position + 2]))
            {
                return 2;
            }

            if (text[position] == '*' &&
                position + 1 < text.Length &&
                !char.IsWhiteSpace(text[position + 1]))
            {
                return 1;
            }

            return 0;
        }

        private static bool IsBetweenLettersOrDigits(
            string text,
            int position)
        {
            return position > 0 &&
                position + 1 < text.Length &&
                char.IsLetterOrDigit(text[position - 1]) &&
                char.IsLetterOrDigit(text[position + 1]);
        }

        private static bool EndsLine(string value)
        {
            return value.EndsWith(
                       "\n",
                       StringComparison.Ordinal) ||
                   value.EndsWith(
                       "\r",
                       StringComparison.Ordinal);
        }
    }

    public sealed class FormattedModelText
    {
        internal FormattedModelText(
            string plainText,
            IReadOnlyList<TextSpan> boldRanges)
        {
            PlainText = plainText ?? string.Empty;
            BoldRanges = boldRanges ?? new TextSpan[0];
        }

        public string PlainText { get; }

        public IReadOnlyList<TextSpan> BoldRanges { get; }
    }

    public sealed class TextSpan
    {
        public TextSpan(int start, int length)
        {
            if (start < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(start));
            }

            if (length < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(length));
            }

            Start = start;
            Length = length;
        }

        public int Start { get; }

        public int Length { get; }
    }
}
