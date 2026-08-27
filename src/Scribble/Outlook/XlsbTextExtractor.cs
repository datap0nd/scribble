using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace Scribble.Outlook
{
    // Zero-dependency text extractor for Excel Binary Workbooks
    // (.xlsb). The format is a ZIP container holding BIFF12 record
    // streams (MS-XLSB): each record is a 7-bit variable-length id
    // (one or two bytes) followed by a 7-bit variable-length payload
    // size (one to four bytes). Only the shared-string table and the
    // cell-value records are decoded; everything else is skipped by
    // record length, so unknown records never derail extraction.
    public static class XlsbTextExtractor
    {
        // Cell and worksheet record ids from MS-XLSB.
        private const int BrtRowHdr = 0;
        private const int BrtCellRk = 2;
        private const int BrtCellBool = 4;
        private const int BrtCellReal = 5;
        private const int BrtCellSt = 6;
        private const int BrtCellIsst = 7;
        private const int BrtFmlaString = 8;
        private const int BrtFmlaNum = 9;
        private const int BrtFmlaBool = 10;
        private const int BrtSstItem = 19;

        // A 25 MB workbook can decompress far larger; each part is
        // read up to this cap so a hostile archive cannot balloon
        // memory. A truncated final record simply ends extraction.
        private const int MaxPartBytes = 32 * 1024 * 1024;
        private const int MaxSharedStrings = 500000;
        private const int MaxSharedStringCharacters = 16 * 1024 * 1024;

        public static string Extract(byte[] bytes, int maxCharacters)
        {
            try
            {
                return ExtractCore(
                    bytes,
                    Math.Max(1, maxCharacters));
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string ExtractCore(
            byte[] bytes,
            int maxCharacters)
        {
            using (var archive = new ZipArchive(
                new MemoryStream(bytes),
                ZipArchiveMode.Read))
            {
                var sharedStrings = ReadSharedStrings(archive);
                var sheets = archive.Entries
                    .Where(entry =>
                        entry.FullName.StartsWith(
                            "xl/worksheets/sheet",
                            StringComparison.OrdinalIgnoreCase) &&
                        entry.FullName.EndsWith(
                            ".bin",
                            StringComparison.OrdinalIgnoreCase))
                    .OrderBy(entry => SheetNumber(entry.FullName))
                    .ToList();
                var builder = new StringBuilder();
                foreach (var sheet in sheets)
                {
                    if (sheets.Count > 1)
                    {
                        builder.AppendLine(
                            "[Sheet " +
                            SheetNumber(sheet.FullName).ToString() +
                            "]");
                    }

                    ExtractSheet(
                        ReadPart(sheet),
                        sharedStrings,
                        builder,
                        maxCharacters);
                    if (builder.Length >= maxCharacters)
                    {
                        break;
                    }
                }

                return builder.ToString();
            }
        }

        private static void ExtractSheet(
            byte[] data,
            IList<string> sharedStrings,
            StringBuilder builder,
            int maxCharacters)
        {
            var position = 0;
            var rowValues = new List<string>();
            int id;
            int start;
            int length;
            while (TryReadRecord(
                data,
                ref position,
                out id,
                out start,
                out length))
            {
                if (id == BrtRowHdr)
                {
                    FlushRow(rowValues, builder);
                    if (builder.Length >= maxCharacters)
                    {
                        return;
                    }

                    continue;
                }

                // All cell records begin with a 4-byte column index
                // and a 4-byte style field; the value follows.
                if (length < 8)
                {
                    continue;
                }

                var valueOffset = start + 8;
                var valueLength = length - 8;
                switch (id)
                {
                    case BrtCellRk:
                        if (valueLength >= 4)
                        {
                            rowValues.Add(RkToString(
                                ReadUInt32(data, valueOffset)));
                        }

                        break;
                    case BrtCellBool:
                    case BrtFmlaBool:
                        if (valueLength >= 1)
                        {
                            rowValues.Add(
                                data[valueOffset] != 0
                                    ? "TRUE"
                                    : "FALSE");
                        }

                        break;
                    case BrtCellReal:
                    case BrtFmlaNum:
                        if (valueLength >= 8)
                        {
                            rowValues.Add(NumberToString(
                                BitConverter.ToDouble(
                                    data,
                                    valueOffset)));
                        }

                        break;
                    case BrtCellSt:
                    case BrtFmlaString:
                        var inline = ReadWideString(
                            data,
                            valueOffset,
                            valueOffset + valueLength);
                        if (inline.Length > 0)
                        {
                            rowValues.Add(inline);
                        }

                        break;
                    case BrtCellIsst:
                        if (valueLength >= 4)
                        {
                            var index = (int)ReadUInt32(
                                data,
                                valueOffset);
                            if (index >= 0 &&
                                index < sharedStrings.Count &&
                                sharedStrings[index].Length > 0)
                            {
                                rowValues.Add(
                                    sharedStrings[index]);
                            }
                        }

                        break;
                }
            }

            FlushRow(rowValues, builder);
        }

        private static void FlushRow(
            List<string> rowValues,
            StringBuilder builder)
        {
            if (rowValues.Count > 0)
            {
                builder.AppendLine(
                    string.Join("\t", rowValues));
                rowValues.Clear();
            }
        }

        private static IList<string> ReadSharedStrings(
            ZipArchive archive)
        {
            var entry = archive.GetEntry("xl/sharedStrings.bin");
            if (entry == null)
            {
                return new string[0];
            }

            var data = ReadPart(entry);
            var strings = new List<string>();
            var totalCharacters = 0L;
            var position = 0;
            int id;
            int start;
            int length;
            while (TryReadRecord(
                       data,
                       ref position,
                       out id,
                       out start,
                       out length) &&
                   strings.Count < MaxSharedStrings &&
                   totalCharacters < MaxSharedStringCharacters)
            {
                if (id != BrtSstItem || length < 5)
                {
                    continue;
                }

                // BrtSSTItem: one flags byte, then an XLWideString
                // (4-byte character count plus UTF-16 code units).
                // Rich-text runs that may follow are ignored.
                var value = ReadWideString(
                    data,
                    start + 1,
                    start + length);
                strings.Add(value);
                totalCharacters += value.Length;
            }

            return strings;
        }

        private static string ReadWideString(
            byte[] data,
            int offset,
            int limit)
        {
            if (offset + 4 > limit)
            {
                return string.Empty;
            }

            var count = ReadUInt32(data, offset);
            if (count > int.MaxValue / 2)
            {
                return string.Empty;
            }

            var byteCount = (int)count * 2;
            if (offset + 4 + byteCount > limit)
            {
                return string.Empty;
            }

            return Encoding.Unicode.GetString(
                data,
                offset + 4,
                byteCount);
        }

        private static bool TryReadRecord(
            byte[] data,
            ref int position,
            out int id,
            out int start,
            out int length)
        {
            id = 0;
            start = 0;
            length = 0;
            if (position >= data.Length)
            {
                return false;
            }

            var first = data[position++];
            id = first & 0x7F;
            if ((first & 0x80) != 0)
            {
                if (position >= data.Length)
                {
                    return false;
                }

                id |= (data[position++] & 0x7F) << 7;
            }

            var shift = 0;
            for (var index = 0; index < 4; index++)
            {
                if (position >= data.Length)
                {
                    return false;
                }

                var value = data[position++];
                length |= (value & 0x7F) << shift;
                shift += 7;
                if ((value & 0x80) == 0)
                {
                    break;
                }
            }

            start = position;
            if (length < 0 || start + length > data.Length)
            {
                return false;
            }

            position = start + length;
            return true;
        }

        private static uint ReadUInt32(byte[] data, int offset)
        {
            return (uint)(data[offset] |
                (data[offset + 1] << 8) |
                (data[offset + 2] << 16) |
                (data[offset + 3] << 24));
        }

        private static string RkToString(uint rk)
        {
            double value;
            if ((rk & 2) != 0)
            {
                value = (int)rk >> 2;
            }
            else
            {
                var bits = (long)((ulong)(rk & 0xFFFFFFFCu) << 32);
                value = BitConverter.Int64BitsToDouble(bits);
            }

            if ((rk & 1) != 0)
            {
                value /= 100.0;
            }

            return NumberToString(value);
        }

        private static string NumberToString(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                return string.Empty;
            }

            if (value == Math.Floor(value) &&
                Math.Abs(value) < 1e15)
            {
                return ((long)value).ToString(
                    CultureInfo.InvariantCulture);
            }

            return value.ToString(
                "G15",
                CultureInfo.InvariantCulture);
        }

        private static byte[] ReadPart(ZipArchiveEntry entry)
        {
            using (var stream = entry.Open())
            using (var buffer = new MemoryStream())
            {
                var chunk = new byte[81920];
                while (buffer.Length < MaxPartBytes)
                {
                    var read = stream.Read(
                        chunk,
                        0,
                        (int)Math.Min(
                            chunk.Length,
                            MaxPartBytes - buffer.Length));
                    if (read <= 0)
                    {
                        break;
                    }

                    buffer.Write(chunk, 0, read);
                }

                return buffer.ToArray();
            }
        }

        private static int SheetNumber(string entryName)
        {
            var digits = new StringBuilder();
            foreach (var character in entryName)
            {
                if (char.IsDigit(character))
                {
                    digits.Append(character);
                }
                else if (digits.Length > 0)
                {
                    break;
                }
            }

            int number;
            return int.TryParse(digits.ToString(), out number)
                ? number
                : 0;
        }
    }
}
