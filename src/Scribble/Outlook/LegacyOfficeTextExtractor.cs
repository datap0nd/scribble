using System;
using System.Collections.Generic;
using System.Text;

namespace Scribble.Outlook
{
    // Local, dependency-free text extraction for legacy binary Office
    // formats (.doc, .ppt, .xls live inside OLE compound files) and RTF.
    // Everything is best-effort and bounded; failures return empty text
    // so the caller can add a visible note instead of dropping the file.
    public static class LegacyOfficeTextExtractor
    {
        private const uint EndOfChain = 0xFFFFFFFE;
        private const uint FreeSector = 0xFFFFFFFF;
        private const int MaxChainLength = 65536;
        private const int MaxOutputCharacters = 16000;

        public static string ExtractPptText(byte[] fileBytes)
        {
            try
            {
                var stream = ReadCompoundStream(
                    fileBytes,
                    "PowerPoint Document");
                if (stream == null)
                {
                    return string.Empty;
                }

                var builder = new StringBuilder();
                var position = 0;
                while (position + 8 <= stream.Length &&
                       builder.Length < MaxOutputCharacters)
                {
                    var verInstance = ReadUInt16(stream, position);
                    var recordType = ReadUInt16(stream, position + 2);
                    var recordLength = (int)Math.Min(
                        ReadUInt32(stream, position + 4),
                        (uint)(stream.Length - position - 8));
                    if ((verInstance & 0x000F) == 0x000F)
                    {
                        // Container record: descend into its payload.
                        position += 8;
                        continue;
                    }

                    if (recordType == 0x0FA0)
                    {
                        // TextCharsAtom: UTF-16LE text.
                        AppendClean(
                            DecodeUtf16(
                                stream,
                                position + 8,
                                recordLength),
                            builder);
                    }
                    else if (recordType == 0x0FA8)
                    {
                        // TextBytesAtom: single-byte text.
                        AppendClean(
                            DecodeLatin(
                                stream,
                                position + 8,
                                recordLength),
                            builder);
                    }

                    position += 8 + recordLength;
                }

                return builder.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        public static string ExtractDocText(byte[] fileBytes)
        {
            try
            {
                var stream = ReadCompoundStream(
                    fileBytes,
                    "WordDocument");
                if (stream == null)
                {
                    return string.Empty;
                }

                var text = string.Empty;
                if (stream.Length > 0x20)
                {
                    var flags = ReadUInt16(stream, 0x0A);
                    var complex = (flags & 0x0004) != 0;
                    var fcMin = (int)Math.Min(
                        ReadUInt32(stream, 0x18),
                        (uint)stream.Length);
                    var fcMac = (int)Math.Min(
                        ReadUInt32(stream, 0x1C),
                        (uint)stream.Length);
                    if (!complex && fcMac > fcMin)
                    {
                        text = DecodeTextSlice(
                            stream,
                            fcMin,
                            fcMac - fcMin);
                    }
                }

                if (CountLettersAndDigits(text) < 40)
                {
                    text = ScanPrintableRuns(stream);
                }

                return text;
            }
            catch
            {
                return string.Empty;
            }
        }

        public static string ExtractXlsText(byte[] fileBytes)
        {
            try
            {
                var stream =
                    ReadCompoundStream(fileBytes, "Workbook") ??
                    ReadCompoundStream(fileBytes, "Book");
                if (stream == null)
                {
                    return string.Empty;
                }

                var builder = new StringBuilder();
                var position = 0;
                while (position + 4 <= stream.Length &&
                       builder.Length < MaxOutputCharacters)
                {
                    var opcode = ReadUInt16(stream, position);
                    var size = ReadUInt16(stream, position + 2);
                    var payload = position + 4;
                    if (payload + size > stream.Length)
                    {
                        break;
                    }

                    if (opcode == 0x00FC)
                    {
                        // Shared string table: cell text without
                        // positions, split across Continue records
                        // is skipped on boundary overrun.
                        ParseSharedStrings(
                            stream,
                            payload,
                            size,
                            builder);
                    }

                    position = payload + size;
                }

                return builder.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        public static string ExtractRtfText(byte[] fileBytes)
        {
            try
            {
                var text = Encoding.GetEncoding("ISO-8859-1")
                    .GetString(fileBytes);
                var builder = new StringBuilder();
                var skipDepth = 0;
                var depth = 0;
                var index = 0;
                while (index < text.Length &&
                       builder.Length < MaxOutputCharacters)
                {
                    var character = text[index];
                    if (character == '{')
                    {
                        depth++;
                        index++;
                        continue;
                    }

                    if (character == '}')
                    {
                        if (skipDepth > 0 && depth <= skipDepth)
                        {
                            skipDepth = 0;
                        }

                        depth = Math.Max(0, depth - 1);
                        index++;
                        continue;
                    }

                    if (character != '\\')
                    {
                        if (skipDepth == 0 &&
                            character != '\r' &&
                            character != '\n')
                        {
                            builder.Append(character);
                        }

                        index++;
                        continue;
                    }

                    index++;
                    if (index >= text.Length)
                    {
                        break;
                    }

                    var next = text[index];
                    if (next == '\'')
                    {
                        if (index + 2 < text.Length &&
                            skipDepth == 0)
                        {
                            var value = HexValue(text[index + 1]) * 16 +
                                HexValue(text[index + 2]);
                            if (value >= 32)
                            {
                                builder.Append((char)value);
                            }
                        }

                        index += 3;
                        continue;
                    }

                    if (next == '*')
                    {
                        // Optional destination: skip its group.
                        if (skipDepth == 0)
                        {
                            skipDepth = depth;
                        }

                        index++;
                        continue;
                    }

                    if (!char.IsLetter(next))
                    {
                        if (skipDepth == 0 &&
                            (next == '\\' ||
                             next == '{' ||
                             next == '}'))
                        {
                            builder.Append(next);
                        }

                        index++;
                        continue;
                    }

                    var wordStart = index;
                    while (index < text.Length &&
                           char.IsLetter(text[index]))
                    {
                        index++;
                    }

                    var word = text.Substring(
                        wordStart,
                        index - wordStart);
                    var negative = false;
                    if (index < text.Length && text[index] == '-')
                    {
                        negative = true;
                        index++;
                    }

                    var numberStart = index;
                    while (index < text.Length &&
                           char.IsDigit(text[index]))
                    {
                        index++;
                    }

                    var number = index > numberStart
                        ? text.Substring(
                            numberStart,
                            index - numberStart)
                        : string.Empty;
                    if (index < text.Length && text[index] == ' ')
                    {
                        index++;
                    }

                    if (skipDepth > 0)
                    {
                        continue;
                    }

                    if (IsRtfSkipDestination(word))
                    {
                        skipDepth = depth;
                    }
                    else if (word == "par" ||
                             word == "line" ||
                             word == "sect" ||
                             word == "page")
                    {
                        builder.Append('\n');
                    }
                    else if (word == "tab" || word == "cell")
                    {
                        builder.Append('\t');
                    }
                    else if (word == "u" && number.Length > 0)
                    {
                        int code;
                        if (int.TryParse(number, out code) &&
                            !negative &&
                            code >= 32)
                        {
                            builder.Append((char)code);
                        }

                        // Skip the fallback character after \uN.
                        if (index < text.Length &&
                            text[index] != '\\' &&
                            text[index] != '{' &&
                            text[index] != '}')
                        {
                            index++;
                        }
                    }
                }

                return builder.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        // Outlook .msg / .oft items are compound files whose MAPI
        // string properties live in named substreams: 001F suffixes
        // are UTF-16, 001E are the ANSI code page. Subject, sender,
        // recipients, and the plain-text body are enough to make an
        // attached message readable; nested attachments inside the
        // .msg are not expanded.
        public static string ExtractMsgText(byte[] fileBytes)
        {
            try
            {
                var builder = new StringBuilder();
                var subject = ReadMsgString(fileBytes, "0037");
                var sender = ReadMsgString(fileBytes, "0C1A");
                var recipients = ReadMsgString(fileBytes, "0E04");
                var body = ReadMsgString(fileBytes, "1000");
                if (subject.Length > 0)
                {
                    builder.AppendLine("Subject: " + subject);
                }

                if (sender.Length > 0)
                {
                    builder.AppendLine("From: " + sender);
                }

                if (recipients.Length > 0)
                {
                    builder.AppendLine("To: " + recipients);
                }

                if (builder.Length > 0 && body.Length > 0)
                {
                    builder.AppendLine();
                }

                builder.Append(body);
                return builder.ToString().Trim();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string ReadMsgString(
            byte[] fileBytes,
            string propertyId)
        {
            var unicode = ReadCompoundStream(
                fileBytes,
                "__substg1.0_" + propertyId + "001F");
            if (unicode != null && unicode.Length > 1)
            {
                return Encoding.Unicode
                    .GetString(unicode)
                    .Trim('\0')
                    .Trim();
            }

            var ansi = ReadCompoundStream(
                fileBytes,
                "__substg1.0_" + propertyId + "001E");
            if (ansi != null && ansi.Length > 0)
            {
                return Encoding.Default
                    .GetString(ansi)
                    .Trim('\0')
                    .Trim();
            }

            return string.Empty;
        }

        public static bool LooksLikeCompoundFile(byte[] bytes)
        {
            return bytes != null &&
                   bytes.Length > 8 &&
                   bytes[0] == 0xD0 &&
                   bytes[1] == 0xCF &&
                   bytes[2] == 0x11 &&
                   bytes[3] == 0xE0;
        }

        public static bool CompoundStreamExists(
            byte[] fileBytes,
            string streamName)
        {
            try
            {
                return ReadCompoundStream(
                    fileBytes,
                    streamName) != null;
            }
            catch
            {
                return false;
            }
        }

        public static byte[] ReadCompoundStream(
            byte[] bytes,
            string streamName)
        {
            if (!LooksLikeCompoundFile(bytes) ||
                bytes.Length < 512)
            {
                return null;
            }

            var sectorShift = ReadUInt16(bytes, 30);
            if (sectorShift < 7 || sectorShift > 12)
            {
                return null;
            }

            var sectorSize = 1 << sectorShift;
            var miniShift = ReadUInt16(bytes, 32);
            var miniSectorSize = 1 << Math.Max(
                4,
                Math.Min(12, (int)miniShift));
            var firstDirSector = ReadUInt32(bytes, 48);
            var miniCutoff = ReadUInt32(bytes, 56);
            var firstMiniFatSector = ReadUInt32(bytes, 60);
            var firstDifatSector = ReadUInt32(bytes, 68);

            var fatSectors = new List<uint>();
            for (var index = 0; index < 109; index++)
            {
                var value = ReadUInt32(bytes, 76 + index * 4);
                if (value != FreeSector && value != EndOfChain)
                {
                    fatSectors.Add(value);
                }
            }

            var difatSector = firstDifatSector;
            var difatGuard = 0;
            while (difatSector != EndOfChain &&
                   difatSector != FreeSector &&
                   difatGuard++ < MaxChainLength)
            {
                var offset = SectorOffset(
                    difatSector,
                    sectorSize);
                if (offset + sectorSize > bytes.Length)
                {
                    break;
                }

                var perSector = sectorSize / 4 - 1;
                for (var index = 0; index < perSector; index++)
                {
                    var value = ReadUInt32(
                        bytes,
                        offset + index * 4);
                    if (value != FreeSector &&
                        value != EndOfChain)
                    {
                        fatSectors.Add(value);
                    }
                }

                difatSector = ReadUInt32(
                    bytes,
                    offset + sectorSize - 4);
            }

            var fat = new List<uint>();
            foreach (var fatSector in fatSectors)
            {
                var offset = SectorOffset(fatSector, sectorSize);
                if (offset + sectorSize > bytes.Length)
                {
                    continue;
                }

                for (var index = 0;
                     index < sectorSize / 4;
                     index++)
                {
                    fat.Add(ReadUInt32(
                        bytes,
                        offset + index * 4));
                }
            }

            var directory = ReadChain(
                bytes,
                fat,
                firstDirSector,
                sectorSize,
                long.MaxValue);
            if (directory == null)
            {
                return null;
            }

            byte[] rootChainData = null;
            uint targetStart = 0;
            long targetSize = -1;
            uint rootStart = 0;
            long rootSize = 0;
            for (var entry = 0;
                 entry + 128 <= directory.Length;
                 entry += 128)
            {
                var nameLength = ReadUInt16(
                    directory,
                    entry + 64);
                if (nameLength < 2 || nameLength > 64)
                {
                    continue;
                }

                var name = Encoding.Unicode.GetString(
                    directory,
                    entry,
                    nameLength - 2);
                var objectType = directory[entry + 66];
                var start = ReadUInt32(directory, entry + 116);
                var size = (long)ReadUInt32(
                    directory,
                    entry + 120);
                if (objectType == 5)
                {
                    rootStart = start;
                    rootSize = size;
                }
                else if (objectType == 2 &&
                         string.Equals(
                             name,
                             streamName,
                             StringComparison.OrdinalIgnoreCase))
                {
                    targetStart = start;
                    targetSize = size;
                }
            }

            if (targetSize < 0)
            {
                return null;
            }

            if (targetSize >= miniCutoff)
            {
                return ReadChain(
                    bytes,
                    fat,
                    targetStart,
                    sectorSize,
                    targetSize);
            }

            rootChainData = ReadChain(
                bytes,
                fat,
                rootStart,
                sectorSize,
                rootSize);
            var miniFat = new List<uint>();
            var miniFatData = ReadChain(
                bytes,
                fat,
                firstMiniFatSector,
                sectorSize,
                long.MaxValue);
            if (rootChainData == null || miniFatData == null)
            {
                return null;
            }

            for (var index = 0;
                 index + 4 <= miniFatData.Length;
                 index += 4)
            {
                miniFat.Add(ReadUInt32(miniFatData, index));
            }

            using (var output = new System.IO.MemoryStream())
            {
                var sector = targetStart;
                var guard = 0;
                var remaining = targetSize;
                while (sector != EndOfChain &&
                       sector != FreeSector &&
                       remaining > 0 &&
                       guard++ < MaxChainLength)
                {
                    var offset = (long)sector * miniSectorSize;
                    if (offset + miniSectorSize >
                        rootChainData.Length)
                    {
                        break;
                    }

                    var take = (int)Math.Min(
                        miniSectorSize,
                        remaining);
                    output.Write(
                        rootChainData,
                        (int)offset,
                        take);
                    remaining -= take;
                    sector = sector < miniFat.Count
                        ? miniFat[(int)sector]
                        : EndOfChain;
                }

                return output.ToArray();
            }
        }

        private static byte[] ReadChain(
            byte[] bytes,
            List<uint> fat,
            uint startSector,
            int sectorSize,
            long maximumBytes)
        {
            if (startSector == EndOfChain ||
                startSector == FreeSector)
            {
                return null;
            }

            using (var output = new System.IO.MemoryStream())
            {
                var sector = startSector;
                var guard = 0;
                while (sector != EndOfChain &&
                       sector != FreeSector &&
                       guard++ < MaxChainLength)
                {
                    var offset = SectorOffset(sector, sectorSize);
                    if (offset + sectorSize > bytes.Length)
                    {
                        break;
                    }

                    output.Write(bytes, offset, sectorSize);
                    if (output.Length >= maximumBytes)
                    {
                        break;
                    }

                    sector = sector < fat.Count
                        ? fat[(int)sector]
                        : EndOfChain;
                }

                var result = output.ToArray();
                if (maximumBytes < result.Length)
                {
                    var bounded = new byte[maximumBytes];
                    Array.Copy(
                        result,
                        bounded,
                        (int)maximumBytes);
                    return bounded;
                }

                return result;
            }
        }

        private static void ParseSharedStrings(
            byte[] stream,
            int payload,
            int size,
            StringBuilder builder)
        {
            var position = payload + 8;
            var end = payload + size;
            while (position + 3 <= end &&
                   builder.Length < MaxOutputCharacters)
            {
                var length = ReadUInt16(stream, position);
                var grbit = stream[position + 2];
                position += 3;
                var runs = 0;
                var extension = 0;
                if ((grbit & 0x08) != 0)
                {
                    if (position + 2 > end)
                    {
                        return;
                    }

                    runs = ReadUInt16(stream, position);
                    position += 2;
                }

                if ((grbit & 0x04) != 0)
                {
                    if (position + 4 > end)
                    {
                        return;
                    }

                    extension = (int)ReadUInt32(
                        stream,
                        position);
                    position += 4;
                }

                var wide = (grbit & 0x01) != 0;
                var byteCount = wide ? length * 2 : length;
                if (position + byteCount > end)
                {
                    return;
                }

                AppendClean(
                    wide
                        ? DecodeUtf16(
                            stream,
                            position,
                            byteCount)
                        : DecodeLatin(
                            stream,
                            position,
                            byteCount),
                    builder);
                position += byteCount + runs * 4 + extension;
            }
        }

        private static string DecodeTextSlice(
            byte[] stream,
            int offset,
            int length)
        {
            var zeroAtOdd = 0;
            var samples = 0;
            for (var index = offset + 1;
                 index < offset + length &&
                 samples < 512;
                 index += 2)
            {
                if (stream[index] == 0)
                {
                    zeroAtOdd++;
                }

                samples++;
            }

            var utf16 = samples > 0 &&
                zeroAtOdd * 10 >= samples * 4;
            return utf16
                ? DecodeUtf16(stream, offset, length)
                : DecodeLatin(stream, offset, length);
        }

        private static string ScanPrintableRuns(byte[] stream)
        {
            var builder = new StringBuilder();
            var run = new StringBuilder();
            for (var index = 0;
                 index < stream.Length &&
                 builder.Length < MaxOutputCharacters;
                 index++)
            {
                var isWide = index + 1 < stream.Length &&
                    stream[index + 1] == 0 &&
                    IsPrintable(stream[index]);
                if (isWide)
                {
                    run.Append((char)stream[index]);
                    index++;
                    continue;
                }

                if (IsPrintable(stream[index]))
                {
                    run.Append((char)stream[index]);
                    continue;
                }

                FlushRun(run, builder);
            }

            FlushRun(run, builder);
            return builder.ToString();
        }

        private static void FlushRun(
            StringBuilder run,
            StringBuilder builder)
        {
            if (run.Length >= 8)
            {
                builder.Append(run);
                builder.Append('\n');
            }

            run.Length = 0;
        }

        private static bool IsPrintable(byte value)
        {
            return value == 9 ||
                   (value >= 32 && value < 127);
        }

        private static bool IsRtfSkipDestination(string word)
        {
            switch (word)
            {
                case "fonttbl":
                case "colortbl":
                case "stylesheet":
                case "info":
                case "pict":
                case "object":
                case "themedata":
                case "colorschememapping":
                case "datastore":
                case "latentstyles":
                case "listtable":
                case "listoverridetable":
                case "generator":
                    return true;
                default:
                    return false;
            }
        }

        private static void AppendClean(
            string text,
            StringBuilder builder)
        {
            var appended = false;
            foreach (var character in text)
            {
                if (character == '\r' || character == '\v')
                {
                    builder.Append('\n');
                    appended = true;
                }
                else if (character == '\n' ||
                         character == '\t' ||
                         !char.IsControl(character))
                {
                    builder.Append(character);
                    appended = true;
                }
            }

            if (appended)
            {
                builder.Append('\n');
            }
        }

        private static string DecodeUtf16(
            byte[] bytes,
            int offset,
            int length)
        {
            var bounded = Math.Max(
                0,
                Math.Min(length, bytes.Length - offset));
            return bounded > 1
                ? Encoding.Unicode.GetString(
                    bytes,
                    offset,
                    bounded - bounded % 2)
                : string.Empty;
        }

        private static string DecodeLatin(
            byte[] bytes,
            int offset,
            int length)
        {
            var bounded = Math.Max(
                0,
                Math.Min(length, bytes.Length - offset));
            return bounded > 0
                ? Encoding.GetEncoding("ISO-8859-1").GetString(
                    bytes,
                    offset,
                    bounded)
                : string.Empty;
        }

        private static int SectorOffset(uint sector, int sectorSize)
        {
            return (int)Math.Min(
                int.MaxValue,
                ((long)sector + 1) * sectorSize);
        }

        private static int HexValue(char character)
        {
            if (character >= '0' && character <= '9')
            {
                return character - '0';
            }

            if (character >= 'a' && character <= 'f')
            {
                return character - 'a' + 10;
            }

            if (character >= 'A' && character <= 'F')
            {
                return character - 'A' + 10;
            }

            return 0;
        }

        private static int CountLettersAndDigits(string text)
        {
            var count = 0;
            foreach (var character in text ?? string.Empty)
            {
                if (char.IsLetterOrDigit(character))
                {
                    count++;
                }
            }

            return count;
        }

        private static ushort ReadUInt16(byte[] bytes, int offset)
        {
            return offset + 2 <= bytes.Length
                ? (ushort)(bytes[offset] |
                           (bytes[offset + 1] << 8))
                : (ushort)0;
        }

        private static uint ReadUInt32(byte[] bytes, int offset)
        {
            return offset + 4 <= bytes.Length
                ? (uint)(bytes[offset] |
                         (bytes[offset + 1] << 8) |
                         (bytes[offset + 2] << 16) |
                         (bytes[offset + 3] << 24))
                : 0U;
        }
    }
}
