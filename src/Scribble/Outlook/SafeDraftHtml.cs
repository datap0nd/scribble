using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using Scribble.Security;

namespace Scribble.Outlook
{
    public static class SafeDraftHtml
    {
        public static string Format(
            string body,
            IReadOnlyList<string> boldPhrases)
        {
            return FormatContent(body, boldPhrases).Html;
        }

        public static SafeDraftContent FormatContent(
            string body,
            IReadOnlyList<string> boldPhrases)
        {
            var formatted = SafeModelText.Format(
                body,
                TextBoundary.MaxAssistantCharacters);
            var text = formatted.PlainText;
            if (text.Length == 0)
            {
                throw new InvalidOperationException(
                    "A non-empty draft body is required.");
            }

            var ranges = FindBoldRanges(text, boldPhrases);
            foreach (var range in formatted.BoldRanges)
            {
                ranges.Add(
                    new TextRange(
                        range.Start,
                        range.Length));
            }
            ranges = MergeRanges(ranges);
            var html = new StringBuilder(text.Length + 512);
            html.Append(
                "<div style=\"font-family:Calibri,Arial,sans-serif;font-size:11pt;\">");
            var lines = SplitLines(text);
            var listKind = 0;
            var index = 0;
            while (index < lines.Count)
            {
                var consumed = TryRenderTable(
                    html,
                    text,
                    lines,
                    index,
                    ranges,
                    ref listKind);
                if (consumed > 0)
                {
                    index += consumed;
                    continue;
                }

                RenderLine(
                    html,
                    text,
                    lines[index][0],
                    lines[index][1],
                    ranges,
                    ref listKind);
                index++;
            }

            CloseList(html, ref listKind);
            html.Append("</div>");
            return new SafeDraftContent(
                text,
                html.ToString());
        }

        private static List<int[]> SplitLines(string text)
        {
            var lines = new List<int[]>();
            var position = 0;
            while (position < text.Length)
            {
                var lineStart = position;
                while (position < text.Length &&
                    text[position] != '\r' &&
                    text[position] != '\n')
                {
                    position++;
                }

                lines.Add(new[]
                {
                    lineStart,
                    position - lineStart
                });
                if (position < text.Length && text[position] == '\r')
                {
                    position++;
                }
                if (position < text.Length && text[position] == '\n')
                {
                    position++;
                }
            }

            return lines;
        }

        private static List<TextRange> FindBoldRanges(
            string text,
            IReadOnlyList<string> boldPhrases)
        {
            var candidates = new List<TextRange>();
            if (boldPhrases == null)
            {
                return candidates;
            }

            var phraseCount = Math.Min(12, boldPhrases.Count);
            for (var phraseIndex = 0;
                 phraseIndex < phraseCount;
                 phraseIndex++)
            {
                var phrase = TextBoundary.PlainText(
                    boldPhrases[phraseIndex],
                    200);
                if (phrase.Length == 0)
                {
                    continue;
                }

                var start = 0;
                while (start < text.Length)
                {
                    var found = text.IndexOf(
                        phrase,
                        start,
                        StringComparison.OrdinalIgnoreCase);
                    if (found < 0)
                    {
                        break;
                    }

                    candidates.Add(
                        new TextRange(found, phrase.Length));
                    start = found + phrase.Length;
                }
            }

            return MergeRanges(candidates);
        }

        private static List<TextRange> MergeRanges(
            List<TextRange> candidates)
        {
            candidates.Sort((left, right) =>
            {
                var startComparison = left.Start.CompareTo(right.Start);
                return startComparison != 0
                    ? startComparison
                    : right.Length.CompareTo(left.Length);
            });

            var accepted = new List<TextRange>();
            var end = -1;
            foreach (var candidate in candidates)
            {
                if (candidate.Start < end)
                {
                    continue;
                }

                accepted.Add(candidate);
                end = candidate.Start + candidate.Length;
            }

            return accepted;
        }

        private const int MaxTableRows = 30;
        private const int MaxTableColumns = 10;

        // Renders a pipe table (| a | b | header, | --- | separator,
        // data rows) as a bordered HTML table built entirely from
        // encoded text. Returns the number of consumed lines, or 0
        // when the lines at this position are not a table.
        private static int TryRenderTable(
            StringBuilder output,
            string text,
            IReadOnlyList<int[]> lines,
            int index,
            IReadOnlyList<TextRange> ranges,
            ref int listKind)
        {
            if (index + 1 >= lines.Count)
            {
                return 0;
            }

            var header = ParseTableCells(
                text,
                lines[index][0],
                lines[index][1]);
            if (header == null)
            {
                return 0;
            }

            var alignments = ParseSeparatorCells(
                text,
                lines[index + 1][0],
                lines[index + 1][1]);
            if (alignments == null)
            {
                return 0;
            }

            CloseList(output, ref listKind);
            output.Append(
                "<table style=\"border-collapse:collapse;margin:6px 0 12px 0;\">");
            output.Append("<tr>");
            RenderTableRow(
                output,
                text,
                header,
                alignments,
                ranges,
                true);
            output.Append("</tr>");
            var consumed = 2;
            var rendered = 0;
            while (index + consumed < lines.Count &&
                   rendered < MaxTableRows)
            {
                var row = ParseTableCells(
                    text,
                    lines[index + consumed][0],
                    lines[index + consumed][1]);
                if (row == null)
                {
                    break;
                }

                // A stray separator row inside the body is layout
                // punctuation, never data.
                if (ParseSeparatorCells(
                        text,
                        lines[index + consumed][0],
                        lines[index + consumed][1]) != null)
                {
                    consumed++;
                    continue;
                }

                output.Append("<tr>");
                RenderTableRow(
                    output,
                    text,
                    row,
                    alignments,
                    ranges,
                    false);
                output.Append("</tr>");
                consumed++;
                rendered++;
            }

            output.Append("</table>");
            return consumed;
        }

        private static void RenderTableRow(
            StringBuilder output,
            string text,
            IReadOnlyList<int[]> cells,
            IReadOnlyList<int> alignments,
            IReadOnlyList<TextRange> ranges,
            bool isHeader)
        {
            for (var cell = 0; cell < cells.Count; cell++)
            {
                var alignment = cell < alignments.Count
                    ? alignments[cell]
                    : 0;
                output.Append(isHeader ? "<th style=\"" : "<td style=\"");
                output.Append(
                    "border:1px solid #d0d7de;padding:4px 10px;");
                if (isHeader)
                {
                    output.Append("background:#f6f8fa;");
                }

                output.Append(
                    alignment == 2
                        ? "text-align:right;"
                        : (alignment == 1
                            ? "text-align:center;"
                            : "text-align:left;"));
                output.Append("\">");
                AppendEncodedRange(
                    output,
                    text,
                    cells[cell][0],
                    cells[cell][1],
                    ranges);
                output.Append(isHeader ? "</th>" : "</td>");
            }
        }

        // Splits one |-delimited line into trimmed cell ranges, or
        // null when the line is not shaped like a table row.
        private static List<int[]> ParseTableCells(
            string text,
            int start,
            int length)
        {
            var begin = start;
            var end = start + length;
            while (begin < end && text[begin] == ' ')
            {
                begin++;
            }

            while (end > begin && text[end - 1] == ' ')
            {
                end--;
            }

            if (end - begin < 2 ||
                text[begin] != '|' ||
                text[end - 1] != '|')
            {
                return null;
            }

            var cells = new List<int[]>();
            var cellStart = begin + 1;
            for (var position = begin + 1;
                 position < end && cells.Count < MaxTableColumns;
                 position++)
            {
                if (text[position] == '|')
                {
                    cells.Add(TrimmedRange(
                        text,
                        cellStart,
                        position - cellStart));
                    cellStart = position + 1;
                }
            }

            return cells.Count > 0 ? cells : null;
        }

        // Alignment per column from the separator row: 0 left,
        // 1 center, 2 right; null when the row is not a separator.
        private static List<int> ParseSeparatorCells(
            string text,
            int start,
            int length)
        {
            var cells = ParseTableCells(text, start, length);
            if (cells == null)
            {
                return null;
            }

            var alignments = new List<int>();
            foreach (var cell in cells)
            {
                var dashes = 0;
                var left = false;
                var right = false;
                for (var offset = 0; offset < cell[1]; offset++)
                {
                    var character = text[cell[0] + offset];
                    if (character == '-')
                    {
                        dashes++;
                    }
                    else if (character == ':' && dashes == 0)
                    {
                        left = true;
                    }
                    else if (character == ':' &&
                             offset == cell[1] - 1)
                    {
                        right = true;
                    }
                    else
                    {
                        return null;
                    }
                }

                if (dashes == 0)
                {
                    return null;
                }

                alignments.Add(
                    left && right ? 1 : (right ? 2 : 0));
            }

            return alignments;
        }

        private static int[] TrimmedRange(
            string text,
            int start,
            int length)
        {
            var begin = start;
            var end = start + length;
            while (begin < end && text[begin] == ' ')
            {
                begin++;
            }

            while (end > begin && text[end - 1] == ' ')
            {
                end--;
            }

            return new[] { begin, end - begin };
        }

        private static void RenderLine(
            StringBuilder output,
            string text,
            int start,
            int length,
            IReadOnlyList<TextRange> ranges,
            ref int listKind)
        {
            var line = text.Substring(start, length);
            if (line.Length == 0)
            {
                CloseList(output, ref listKind);
                output.Append(
                    "<div style=\"height:8px;line-height:8px;\">&nbsp;</div>");
                return;
            }

            if (line.Equals("---", StringComparison.Ordinal))
            {
                CloseList(output, ref listKind);
                output.Append(
                    "<hr style=\"border:0;border-top:1px solid #d0d7de;margin:14px 0;\">");
                return;
            }

            var prefix = NumberedPrefixLength(line);
            var requestedList = line.StartsWith("- ", StringComparison.Ordinal)
                ? 1
                : (prefix > 0 ? 2 : 0);
            if (requestedList > 0)
            {
                if (listKind != requestedList)
                {
                    CloseList(output, ref listKind);
                    listKind = requestedList;
                    output.Append(
                        requestedList == 1
                            ? "<ul style=\"margin:4px 0 10px 22px;padding:0;\">"
                            : "<ol style=\"margin:4px 0 10px 22px;padding:0;\">");
                }

                var contentStart = start +
                    (requestedList == 1 ? 2 : prefix);
                output.Append("<li style=\"margin:0 0 4px 0;\">");
                AppendEncodedRange(
                    output,
                    text,
                    contentStart,
                    length - (contentStart - start),
                    ranges);
                output.Append("</li>");
                return;
            }

            CloseList(output, ref listKind);
            if (line.StartsWith("## ", StringComparison.Ordinal))
            {
                output.Append(
                    "<h3 style=\"font-size:12pt;margin:14px 0 6px 0;color:#1f2328;\">");
                AppendEncodedRange(
                    output,
                    text,
                    start + 3,
                    length - 3,
                    ranges);
                output.Append("</h3>");
                return;
            }

            if (line.StartsWith("# ", StringComparison.Ordinal))
            {
                output.Append(
                    "<h2 style=\"font-size:16pt;margin:4px 0 10px 0;color:#005fb8;\">");
                AppendEncodedRange(
                    output,
                    text,
                    start + 2,
                    length - 2,
                    ranges);
                output.Append("</h2>");
                return;
            }

            output.Append("<div style=\"margin:0 0 8px 0;\">");
            AppendEncodedRange(output, text, start, length, ranges);
            output.Append("</div>");
        }

        private static int NumberedPrefixLength(string line)
        {
            var index = 0;
            while (index < line.Length && char.IsDigit(line[index]))
            {
                index++;
            }

            return index > 0 && index + 1 < line.Length &&
                line[index] == '.' && line[index + 1] == ' '
                ? index + 2
                : 0;
        }

        private static void CloseList(
            StringBuilder output,
            ref int listKind)
        {
            if (listKind == 1)
            {
                output.Append("</ul>");
            }
            else if (listKind == 2)
            {
                output.Append("</ol>");
            }

            listKind = 0;
        }

        private static void AppendEncodedRange(
            StringBuilder output,
            string text,
            int start,
            int length,
            IReadOnlyList<TextRange> ranges)
        {
            var end = start + length;
            var position = start;
            foreach (var range in ranges)
            {
                var rangeEnd = range.Start + range.Length;
                if (rangeEnd <= start || range.Start >= end)
                {
                    continue;
                }

                var boldStart = Math.Max(start, range.Start);
                var boldEnd = Math.Min(end, rangeEnd);
                if (boldStart > position)
                {
                    output.Append(WebUtility.HtmlEncode(
                        text.Substring(position, boldStart - position)));
                }

                output.Append("<strong>");
                output.Append(WebUtility.HtmlEncode(
                    text.Substring(boldStart, boldEnd - boldStart)));
                output.Append("</strong>");
                position = boldEnd;
            }

            if (position < end)
            {
                output.Append(WebUtility.HtmlEncode(
                    text.Substring(position, end - position)));
            }
        }

        private sealed class TextRange
        {
            public TextRange(int start, int length)
            {
                Start = start;
                Length = length;
            }

            public int Start { get; }

            public int Length { get; }
        }
    }

    public sealed class SafeDraftContent
    {
        internal SafeDraftContent(
            string plainText,
            string html)
        {
            PlainText = plainText ?? string.Empty;
            Html = html ?? string.Empty;
        }

        public string PlainText { get; }

        public string Html { get; }
    }
}
