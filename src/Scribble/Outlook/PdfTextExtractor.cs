using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace Scribble.Outlook
{
    // Local, dependency-free PDF text-layer extraction. Literal strings
    // are decoded directly; CID/Type0 hex-coded text (the common case in
    // Word, Chrome, and LibreOffice PDFs) is decoded through each font's
    // ToUnicode CMap. Scanned PDFs have no text layer and yield nothing
    // by design; the caller reports that clearly.
    public static class PdfTextExtractor
    {
        private const int MaxInflatedBytesPerStream = 4 * 1024 * 1024;
        private const int MaxCMapEntries = 200000;

        private static readonly Regex ObjectHeader = new Regex(
            "(\\d+)\\s+\\d+\\s+obj\\b",
            RegexOptions.Compiled,
            TimeSpan.FromSeconds(5));

        private static readonly Regex ContentToken = new Regex(
            "/(\\w+)\\s+[-.0-9]+\\s+Tf" +
            "|\\(((?:\\\\.|[^\\\\()])*)\\)\\s*(?:Tj|'|\")" +
            "|<([0-9A-Fa-f\\s]*)>\\s*(?:Tj|'|\")" +
            "|\\[((?:\\((?:\\\\.|[^\\\\()])*\\)|<[0-9A-Fa-f\\s]*>|[^\\]])*)\\]\\s*TJ" +
            "|(T\\*|ET)",
            RegexOptions.Compiled,
            TimeSpan.FromSeconds(5));

        private static readonly Regex ArrayElement = new Regex(
            "\\(((?:\\\\.|[^\\\\()])*)\\)" +
            "|<([0-9A-Fa-f\\s]*)>" +
            "|(-?\\d+(?:\\.\\d+)?)",
            RegexOptions.Compiled,
            TimeSpan.FromSeconds(5));

        private static readonly Regex HexToken = new Regex(
            "<([0-9A-Fa-f]+)>|(\\[)|(\\])",
            RegexOptions.Compiled,
            TimeSpan.FromSeconds(5));

        private sealed class PdfObject
        {
            public string DictionaryText = string.Empty;

            public byte[] StreamData;
        }

        private sealed class PdfFontMap
        {
            public readonly Dictionary<int, string> Codes =
                new Dictionary<int, string>();

            public int CodeBytes = 2;
        }

        private sealed class PdfPage
        {
            public readonly List<int> ContentRefs = new List<int>();

            public readonly Dictionary<string, int> Fonts =
                new Dictionary<string, int>(StringComparer.Ordinal);
        }

        public static string Extract(byte[] bytes, int maxCharacters)
        {
            try
            {
                return ExtractCore(
                    bytes ?? new byte[0],
                    Math.Max(1, maxCharacters));
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string ExtractCore(byte[] bytes, int maxCharacters)
        {
            var objects = IndexObjects(bytes);
            ExpandObjectStreams(objects);
            var fontMaps = BuildFontMaps(objects);
            var pages = FindPages(objects);
            var builder = new StringBuilder();

            foreach (var page in pages)
            {
                foreach (var contentRef in page.ContentRefs)
                {
                    PdfObject contentObject;
                    if (!objects.TryGetValue(
                        contentRef,
                        out contentObject))
                    {
                        continue;
                    }

                    var data = GetStreamData(contentObject);
                    if (data == null)
                    {
                        continue;
                    }

                    DecodeContentStream(
                        Latin1(data),
                        page.Fonts,
                        fontMaps,
                        null,
                        builder,
                        maxCharacters);
                }

                builder.Append('\n');
                if (builder.Length >= maxCharacters)
                {
                    break;
                }
            }

            if (CountLettersAndDigits(builder) < 40)
            {
                builder.Length = 0;
                FallbackScan(
                    bytes,
                    objects,
                    fontMaps,
                    builder,
                    maxCharacters);
            }

            return builder.ToString();
        }

        private static Dictionary<int, PdfObject> IndexObjects(
            byte[] bytes)
        {
            var objects = new Dictionary<int, PdfObject>();
            var latin = Latin1(bytes);
            var position = 0;
            while (position < latin.Length)
            {
                var header = ObjectHeader.Match(latin, position);
                if (!header.Success)
                {
                    break;
                }

                int objectNumber;
                if (!int.TryParse(
                    header.Groups[1].Value,
                    out objectNumber))
                {
                    position = header.Index + header.Length;
                    continue;
                }

                var bodyStart = header.Index + header.Length;
                var streamIndex = latin.IndexOf(
                    "stream",
                    bodyStart,
                    StringComparison.Ordinal);
                var endObjIndex = latin.IndexOf(
                    "endobj",
                    bodyStart,
                    StringComparison.Ordinal);
                var entry = new PdfObject();
                if (streamIndex >= 0 &&
                    (endObjIndex < 0 || streamIndex < endObjIndex))
                {
                    entry.DictionaryText = latin.Substring(
                        bodyStart,
                        streamIndex - bodyStart);
                    var dataStart = streamIndex + 6;
                    if (dataStart < latin.Length &&
                        latin[dataStart] == '\r')
                    {
                        dataStart++;
                    }

                    if (dataStart < latin.Length &&
                        latin[dataStart] == '\n')
                    {
                        dataStart++;
                    }

                    var streamEnd = latin.IndexOf(
                        "endstream",
                        dataStart,
                        StringComparison.Ordinal);
                    if (streamEnd < 0)
                    {
                        break;
                    }

                    entry.StreamData = new byte[streamEnd - dataStart];
                    Array.Copy(
                        bytes,
                        dataStart,
                        entry.StreamData,
                        0,
                        entry.StreamData.Length);
                    position = streamEnd + 9;
                    var closing = latin.IndexOf(
                        "endobj",
                        position,
                        StringComparison.Ordinal);
                    if (closing >= 0 && closing - position < 40)
                    {
                        position = closing + 6;
                    }
                }
                else if (endObjIndex >= 0)
                {
                    entry.DictionaryText = latin.Substring(
                        bodyStart,
                        endObjIndex - bodyStart);
                    position = endObjIndex + 6;
                }
                else
                {
                    entry.DictionaryText = latin.Substring(bodyStart);
                    position = latin.Length;
                }

                if (!objects.ContainsKey(objectNumber))
                {
                    objects[objectNumber] = entry;
                }
            }

            return objects;
        }

        private static void ExpandObjectStreams(
            Dictionary<int, PdfObject> objects)
        {
            var members = new Dictionary<int, PdfObject>();
            foreach (var pair in objects)
            {
                if (pair.Value.DictionaryText.IndexOf(
                        "/ObjStm",
                        StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                var data = GetStreamData(pair.Value);
                if (data == null)
                {
                    continue;
                }

                var text = Latin1(data);
                var count = ReadDictionaryInteger(
                    pair.Value.DictionaryText,
                    "/N");
                var first = ReadDictionaryInteger(
                    pair.Value.DictionaryText,
                    "/First");
                if (count <= 0 || first <= 0 || first > text.Length)
                {
                    continue;
                }

                var numbers = new List<int>();
                foreach (Match match in Regex.Matches(
                    text.Substring(0, first),
                    "\\d+",
                    RegexOptions.None,
                    TimeSpan.FromSeconds(2)))
                {
                    int value;
                    if (int.TryParse(match.Value, out value))
                    {
                        numbers.Add(value);
                    }
                }

                for (var index = 0;
                     index < count &&
                     index * 2 + 1 < numbers.Count;
                     index++)
                {
                    var objectNumber = numbers[index * 2];
                    var start = first + numbers[index * 2 + 1];
                    var end = index * 2 + 3 < numbers.Count &&
                              index + 1 < count
                        ? first + numbers[index * 2 + 3]
                        : text.Length;
                    if (start < 0 ||
                        start >= text.Length ||
                        end <= start)
                    {
                        continue;
                    }

                    end = Math.Min(end, text.Length);
                    if (!objects.ContainsKey(objectNumber) &&
                        !members.ContainsKey(objectNumber))
                    {
                        members[objectNumber] = new PdfObject
                        {
                            DictionaryText = text.Substring(
                                start,
                                end - start)
                        };
                    }
                }
            }

            foreach (var pair in members)
            {
                objects[pair.Key] = pair.Value;
            }
        }

        private static int ReadDictionaryInteger(
            string dictionary,
            string key)
        {
            var match = Regex.Match(
                dictionary,
                Regex.Escape(key) + "\\s+(\\d+)",
                RegexOptions.None,
                TimeSpan.FromSeconds(2));
            int value;
            return match.Success &&
                   int.TryParse(match.Groups[1].Value, out value)
                ? value
                : -1;
        }

        private static Dictionary<int, PdfFontMap> BuildFontMaps(
            Dictionary<int, PdfObject> objects)
        {
            var maps = new Dictionary<int, PdfFontMap>();
            foreach (var pair in objects)
            {
                var match = Regex.Match(
                    pair.Value.DictionaryText,
                    "/ToUnicode\\s+(\\d+)\\s+\\d+\\s+R",
                    RegexOptions.None,
                    TimeSpan.FromSeconds(2));
                if (!match.Success)
                {
                    continue;
                }

                int cmapNumber;
                PdfObject cmapObject;
                if (!int.TryParse(
                        match.Groups[1].Value,
                        out cmapNumber) ||
                    !objects.TryGetValue(cmapNumber, out cmapObject))
                {
                    continue;
                }

                var data = GetStreamData(cmapObject);
                if (data == null)
                {
                    continue;
                }

                var map = ParseCMap(Latin1(data));
                if (map != null && map.Codes.Count > 0)
                {
                    maps[pair.Key] = map;
                }
            }

            return maps;
        }

        private static PdfFontMap ParseCMap(string text)
        {
            var map = new PdfFontMap();
            var codespace = Regex.Match(
                text,
                "begincodespacerange\\s*<([0-9A-Fa-f]+)>",
                RegexOptions.None,
                TimeSpan.FromSeconds(2));
            if (codespace.Success)
            {
                map.CodeBytes = Math.Max(
                    1,
                    codespace.Groups[1].Value.Length / 2);
            }

            foreach (Match block in Regex.Matches(
                text,
                "beginbfchar(.*?)endbfchar",
                RegexOptions.Singleline,
                TimeSpan.FromSeconds(2)))
            {
                foreach (Match entry in Regex.Matches(
                    block.Groups[1].Value,
                    "<([0-9A-Fa-f]+)>\\s*<([0-9A-Fa-f]+)>",
                    RegexOptions.None,
                    TimeSpan.FromSeconds(2)))
                {
                    var code = ParseHex(entry.Groups[1].Value);
                    if (code >= 0 &&
                        map.Codes.Count < MaxCMapEntries)
                    {
                        map.Codes[code] = HexToUnicode(
                            entry.Groups[2].Value);
                    }
                }
            }

            foreach (Match block in Regex.Matches(
                text,
                "beginbfrange(.*?)endbfrange",
                RegexOptions.Singleline,
                TimeSpan.FromSeconds(2)))
            {
                ParseBfRange(block.Groups[1].Value, map);
            }

            return map;
        }

        private static void ParseBfRange(string block, PdfFontMap map)
        {
            var tokens = HexToken.Matches(block);
            var index = 0;
            while (index + 2 < tokens.Count + 1 &&
                   index + 1 < tokens.Count)
            {
                if (!tokens[index].Groups[1].Success ||
                    !tokens[index + 1].Groups[1].Success)
                {
                    index++;
                    continue;
                }

                var low = ParseHex(
                    tokens[index].Groups[1].Value);
                var high = ParseHex(
                    tokens[index + 1].Groups[1].Value);
                if (low < 0 || high < low)
                {
                    index += 2;
                    continue;
                }

                if (index + 2 >= tokens.Count)
                {
                    break;
                }

                if (tokens[index + 2].Groups[2].Success)
                {
                    // Bracketed list: one destination per code.
                    var cursor = index + 3;
                    var code = low;
                    while (cursor < tokens.Count &&
                           tokens[cursor].Groups[1].Success)
                    {
                        if (code <= high &&
                            map.Codes.Count < MaxCMapEntries)
                        {
                            map.Codes[code] = HexToUnicode(
                                tokens[cursor].Groups[1].Value);
                        }

                        code++;
                        cursor++;
                    }

                    if (cursor < tokens.Count &&
                        tokens[cursor].Groups[3].Success)
                    {
                        cursor++;
                    }

                    index = cursor;
                    continue;
                }

                if (!tokens[index + 2].Groups[1].Success)
                {
                    index += 2;
                    continue;
                }

                var destination = HexToUnicode(
                    tokens[index + 2].Groups[1].Value);
                var span = Math.Min(high - low, 65535);
                for (var offset = 0;
                     offset <= span &&
                     map.Codes.Count < MaxCMapEntries;
                     offset++)
                {
                    map.Codes[low + offset] = IncrementLastUnit(
                        destination,
                        offset);
                }

                index += 3;
            }
        }

        private static string IncrementLastUnit(
            string destination,
            int offset)
        {
            if (offset == 0 || destination.Length == 0)
            {
                return destination;
            }

            var last = destination[destination.Length - 1];
            return destination.Substring(
                       0,
                       destination.Length - 1) +
                   (char)(last + offset);
        }

        private static List<PdfPage> FindPages(
            Dictionary<int, PdfObject> objects)
        {
            var pages = new List<PdfPage>();
            foreach (var pair in objects)
            {
                var dictionary = pair.Value.DictionaryText;
                if (!Regex.IsMatch(
                    dictionary,
                    "/Type\\s*/Page\\b",
                    RegexOptions.None,
                    TimeSpan.FromSeconds(2)))
                {
                    continue;
                }

                var page = new PdfPage();
                var single = Regex.Match(
                    dictionary,
                    "/Contents\\s+(\\d+)\\s+\\d+\\s+R",
                    RegexOptions.None,
                    TimeSpan.FromSeconds(2));
                if (single.Success)
                {
                    AddInt(page.ContentRefs, single.Groups[1].Value);
                }
                else
                {
                    var array = Regex.Match(
                        dictionary,
                        "/Contents\\s*\\[([^\\]]*)\\]",
                        RegexOptions.None,
                        TimeSpan.FromSeconds(2));
                    if (array.Success)
                    {
                        foreach (Match reference in Regex.Matches(
                            array.Groups[1].Value,
                            "(\\d+)\\s+\\d+\\s+R",
                            RegexOptions.None,
                            TimeSpan.FromSeconds(2)))
                        {
                            AddInt(
                                page.ContentRefs,
                                reference.Groups[1].Value);
                        }
                    }
                }

                var resources = ResolveResources(
                    dictionary,
                    objects);
                if (resources.Length > 0)
                {
                    var fontRegion = ExtractDictionaryAfter(
                        resources,
                        "/Font",
                        objects);
                    foreach (Match font in Regex.Matches(
                        fontRegion,
                        "/(\\w+)\\s+(\\d+)\\s+\\d+\\s+R",
                        RegexOptions.None,
                        TimeSpan.FromSeconds(2)))
                    {
                        int fontNumber;
                        if (int.TryParse(
                            font.Groups[2].Value,
                            out fontNumber))
                        {
                            page.Fonts[font.Groups[1].Value] =
                                fontNumber;
                        }
                    }
                }

                if (page.ContentRefs.Count > 0)
                {
                    pages.Add(page);
                }
            }

            return pages;
        }

        private static void AddInt(List<int> target, string value)
        {
            int parsed;
            if (int.TryParse(value, out parsed))
            {
                target.Add(parsed);
            }
        }

        private static string ResolveResources(
            string dictionary,
            Dictionary<int, PdfObject> objects)
        {
            var reference = Regex.Match(
                dictionary,
                "/Resources\\s+(\\d+)\\s+\\d+\\s+R",
                RegexOptions.None,
                TimeSpan.FromSeconds(2));
            if (reference.Success)
            {
                int resourceNumber;
                PdfObject resourceObject;
                if (int.TryParse(
                        reference.Groups[1].Value,
                        out resourceNumber) &&
                    objects.TryGetValue(
                        resourceNumber,
                        out resourceObject))
                {
                    return resourceObject.DictionaryText;
                }

                return string.Empty;
            }

            var inline = dictionary.IndexOf(
                "/Resources",
                StringComparison.Ordinal);
            return inline >= 0
                ? BalancedDictionary(dictionary, inline)
                : string.Empty;
        }

        private static string ExtractDictionaryAfter(
            string text,
            string key,
            Dictionary<int, PdfObject> objects)
        {
            var keyIndex = text.IndexOf(
                key,
                StringComparison.Ordinal);
            if (keyIndex < 0)
            {
                return string.Empty;
            }

            var reference = Regex.Match(
                text.Substring(keyIndex),
                "^" + Regex.Escape(key) + "\\s+(\\d+)\\s+\\d+\\s+R",
                RegexOptions.None,
                TimeSpan.FromSeconds(2));
            if (reference.Success)
            {
                int number;
                PdfObject target;
                if (int.TryParse(
                        reference.Groups[1].Value,
                        out number) &&
                    objects.TryGetValue(number, out target))
                {
                    return target.DictionaryText;
                }

                return string.Empty;
            }

            return BalancedDictionary(text, keyIndex);
        }

        private static string BalancedDictionary(
            string text,
            int fromIndex)
        {
            var open = text.IndexOf(
                "<<",
                fromIndex,
                StringComparison.Ordinal);
            if (open < 0)
            {
                return string.Empty;
            }

            var depth = 0;
            for (var index = open;
                 index < text.Length - 1;
                 index++)
            {
                if (text[index] == '<' && text[index + 1] == '<')
                {
                    depth++;
                    index++;
                }
                else if (text[index] == '>' &&
                         text[index + 1] == '>')
                {
                    depth--;
                    index++;
                    if (depth == 0)
                    {
                        return text.Substring(
                            open,
                            index + 1 - open);
                    }
                }
            }

            return string.Empty;
        }

        private static byte[] GetStreamData(PdfObject entry)
        {
            if (entry?.StreamData == null)
            {
                return null;
            }

            var hasFilter = entry.DictionaryText.IndexOf(
                "/Filter",
                StringComparison.Ordinal) >= 0;
            if (!hasFilter)
            {
                return entry.StreamData;
            }

            if (entry.DictionaryText.IndexOf(
                    "/FlateDecode",
                    StringComparison.Ordinal) < 0)
            {
                return null;
            }

            return TryInflate(entry.StreamData);
        }

        private static byte[] TryInflate(byte[] data)
        {
            try
            {
                var offset = data.Length > 2 && data[0] == 0x78
                    ? 2
                    : 0;
                using (var input = new MemoryStream(
                    data,
                    offset,
                    data.Length - offset))
                using (var deflate = new DeflateStream(
                    input,
                    CompressionMode.Decompress))
                using (var output = new MemoryStream())
                {
                    var buffer = new byte[8192];
                    var total = 0;
                    while (true)
                    {
                        var read = deflate.Read(
                            buffer,
                            0,
                            buffer.Length);
                        if (read <= 0)
                        {
                            break;
                        }

                        total += read;
                        if (total > MaxInflatedBytesPerStream)
                        {
                            break;
                        }

                        output.Write(buffer, 0, read);
                    }

                    var inflated = output.ToArray();
                    return inflated.Length > 0 ? inflated : null;
                }
            }
            catch
            {
                return null;
            }
        }

        private static void DecodeContentStream(
            string text,
            Dictionary<string, int> pageFonts,
            Dictionary<int, PdfFontMap> fontMaps,
            PdfFontMap defaultMap,
            StringBuilder builder,
            int maxCharacters)
        {
            var current = defaultMap;
            try
            {
                foreach (Match token in ContentToken.Matches(text))
                {
                    if (builder.Length >= maxCharacters)
                    {
                        return;
                    }

                    if (token.Groups[1].Success)
                    {
                        current = defaultMap;
                        int fontNumber;
                        PdfFontMap map;
                        if (pageFonts != null &&
                            pageFonts.TryGetValue(
                                token.Groups[1].Value,
                                out fontNumber) &&
                            fontMaps.TryGetValue(
                                fontNumber,
                                out map))
                        {
                            current = map;
                        }
                    }
                    else if (token.Groups[2].Success)
                    {
                        AppendLiteral(
                            token.Groups[2].Value,
                            current,
                            builder);
                    }
                    else if (token.Groups[3].Success)
                    {
                        AppendHex(
                            token.Groups[3].Value,
                            current,
                            builder);
                    }
                    else if (token.Groups[4].Success)
                    {
                        AppendArray(
                            token.Groups[4].Value,
                            current,
                            builder);
                    }
                    else if (token.Groups[5].Success)
                    {
                        AppendSeparator(builder, '\n');
                    }
                }
            }
            catch (RegexMatchTimeoutException)
            {
            }
        }

        private static void AppendArray(
            string array,
            PdfFontMap current,
            StringBuilder builder)
        {
            foreach (Match element in ArrayElement.Matches(array))
            {
                if (element.Groups[1].Success)
                {
                    AppendLiteral(
                        element.Groups[1].Value,
                        current,
                        builder);
                }
                else if (element.Groups[2].Success)
                {
                    AppendHex(
                        element.Groups[2].Value,
                        current,
                        builder);
                }
                else if (element.Groups[3].Success)
                {
                    double adjustment;
                    if (double.TryParse(
                            element.Groups[3].Value,
                            System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo
                                .InvariantCulture,
                            out adjustment) &&
                        adjustment < -150)
                    {
                        AppendSeparator(builder, ' ');
                    }
                }
            }

            AppendSeparator(builder, ' ');
        }

        private static void AppendLiteral(
            string escaped,
            PdfFontMap current,
            StringBuilder builder)
        {
            var raw = UnescapeLiteral(escaped);
            if (raw.Length == 0)
            {
                return;
            }

            if (current == null)
            {
                builder.Append(raw);
                AppendSeparator(builder, ' ');
                return;
            }

            AppendCodes(raw, current, builder);
        }

        private static void AppendHex(
            string hex,
            PdfFontMap current,
            StringBuilder builder)
        {
            var digits = new StringBuilder(hex.Length);
            foreach (var character in hex)
            {
                if (!char.IsWhiteSpace(character))
                {
                    digits.Append(character);
                }
            }

            if (digits.Length == 0 || current == null)
            {
                // Hex codes without a ToUnicode map are glyph ids,
                // not characters; decoding them would emit garbage.
                return;
            }

            if (digits.Length % 2 == 1)
            {
                digits.Append('0');
            }

            var raw = new StringBuilder(digits.Length / 2);
            for (var index = 0; index < digits.Length; index += 2)
            {
                var value = ParseHex(
                    digits.ToString(index, 2));
                if (value >= 0)
                {
                    raw.Append((char)value);
                }
            }

            AppendCodes(raw.ToString(), current, builder);
        }

        private static void AppendCodes(
            string raw,
            PdfFontMap map,
            StringBuilder builder)
        {
            var step = Math.Max(1, map.CodeBytes);
            var appended = false;
            for (var index = 0;
                 index + step <= raw.Length;
                 index += step)
            {
                var code = 0;
                for (var offset = 0; offset < step; offset++)
                {
                    code = (code << 8) | raw[index + offset];
                }

                string mapped;
                if (map.Codes.TryGetValue(code, out mapped))
                {
                    builder.Append(mapped);
                    appended = true;
                }
            }

            if (appended)
            {
                AppendSeparator(builder, ' ');
            }
        }

        private static void AppendSeparator(
            StringBuilder builder,
            char separator)
        {
            if (builder.Length > 0 &&
                builder[builder.Length - 1] != separator)
            {
                builder.Append(separator);
            }
        }

        private static string UnescapeLiteral(string escaped)
        {
            var text = new StringBuilder(escaped.Length);
            for (var index = 0; index < escaped.Length; index++)
            {
                var character = escaped[index];
                if (character != '\\')
                {
                    text.Append(character);
                    continue;
                }

                index++;
                if (index >= escaped.Length)
                {
                    break;
                }

                var next = escaped[index];
                if (next == 'n' || next == 'r')
                {
                    text.Append('\n');
                }
                else if (next == 't')
                {
                    text.Append('\t');
                }
                else if (next >= '0' && next <= '7')
                {
                    var octal = 0;
                    var digits = 0;
                    while (digits < 3 &&
                           index < escaped.Length &&
                           escaped[index] >= '0' &&
                           escaped[index] <= '7')
                    {
                        octal = octal * 8 +
                            (escaped[index] - '0');
                        index++;
                        digits++;
                    }

                    index--;
                    text.Append((char)octal);
                }
                else
                {
                    text.Append(next);
                }
            }

            return text.ToString();
        }

        private static void FallbackScan(
            byte[] bytes,
            Dictionary<int, PdfObject> objects,
            Dictionary<int, PdfFontMap> fontMaps,
            StringBuilder builder,
            int maxCharacters)
        {
            // With exactly one mapped font in the file there is no
            // ambiguity, so hex-coded text can still be decoded even
            // when the page-to-font linkage could not be resolved.
            PdfFontMap defaultMap = null;
            if (fontMaps.Count == 1)
            {
                foreach (var pair in fontMaps)
                {
                    defaultMap = pair.Value;
                }
            }

            if (objects.Count == 0)
            {
                DecodeContentStream(
                    Latin1(bytes),
                    null,
                    fontMaps,
                    defaultMap,
                    builder,
                    maxCharacters);
                return;
            }

            foreach (var pair in objects)
            {
                if (builder.Length >= maxCharacters)
                {
                    return;
                }

                var data = GetStreamData(pair.Value);
                if (data == null)
                {
                    continue;
                }

                var text = Latin1(data);
                if (text.IndexOf("BT", StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                DecodeContentStream(
                    text,
                    null,
                    fontMaps,
                    defaultMap,
                    builder,
                    maxCharacters);
            }
        }

        private static int ParseHex(string hex)
        {
            int value;
            return int.TryParse(
                hex,
                System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture,
                out value)
                ? value
                : -1;
        }

        private static string HexToUnicode(string hex)
        {
            var padded = hex.Length % 4 == 0
                ? hex
                : hex.PadLeft(
                    hex.Length + (4 - hex.Length % 4),
                    '0');
            var builder = new StringBuilder(padded.Length / 4);
            for (var index = 0;
                 index + 4 <= padded.Length;
                 index += 4)
            {
                var value = ParseHex(
                    padded.Substring(index, 4));
                if (value > 0)
                {
                    builder.Append((char)value);
                }
            }

            return builder.ToString();
        }

        private static int CountLettersAndDigits(
            StringBuilder builder)
        {
            var count = 0;
            for (var index = 0; index < builder.Length; index++)
            {
                if (char.IsLetterOrDigit(builder[index]))
                {
                    count++;
                }
            }

            return count;
        }

        private static string Latin1(byte[] bytes)
        {
            return Encoding.GetEncoding("ISO-8859-1")
                .GetString(bytes);
        }
    }
}
