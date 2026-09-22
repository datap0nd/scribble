using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace Scribble.Office
{
    public static class OpenXmlWorkbookSnapshotReader
    {
        public const int MaxCells = 200000;
        public const long MaxPartBytes = 32L * 1024L * 1024L;
        public const long MaxExpandedBytes = 256L * 1024L * 1024L;

        public static SourceSnapshot Capture(
            string path,
            string sourceInstanceId,
            string captureRevision,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                throw new FileNotFoundException(
                    "The workbook source is unavailable.", path);
            var extension = Path.GetExtension(path);
            if (!string.Equals(extension, ".xlsx",
                    StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(extension, ".xlsm",
                    StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(extension, ".xltx",
                    StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(extension, ".xltm",
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "ANALYSIS_WORKBOOK_FORMAT_UNSUPPORTED: Typed package capture currently supports OpenXML workbooks.");

            using (var archive = ZipFile.OpenRead(path))
            {
                if (archive.Entries.Count > 10000 ||
                    archive.Entries.Sum(entry => entry.Length) >
                        MaxExpandedBytes)
                    throw new InvalidOperationException(
                        "ANALYSIS_WORKBOOK_RESOURCE_LIMIT: The workbook package exceeds the typed-capture boundary.");
                var shared = SharedStrings(archive, cancellationToken);
                var formats = Formats(archive, cancellationToken);
                var sheets = Sheets(archive, cancellationToken);
                var tables = new List<TableDataset>();
                var locators = new List<SourceLocator>();
                var formulaCount = 0;
                var cells = 0;
                foreach (var sheet in sheets)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var entry = archive.GetEntry(sheet.Path);
                    if (entry == null)
                        throw new InvalidOperationException(
                            "ANALYSIS_WORKBOOK_METADATA_INVALID: A declared worksheet part is missing.");
                    var table = ReadSheet(entry, sheet.Name, shared, formats,
                        ref formulaCount, ref cells, cancellationToken);
                    tables.Add(table);
                    locators.Add(new SourceLocator
                    {
                        Kind = "excel_range",
                        SourceInstanceId = sourceInstanceId,
                        WorksheetIdentity = sheet.Name,
                        Range = table.Rows == 0 || table.Columns == 0
                            ? string.Empty
                            : "A1:" + Address(table.Rows, table.Columns)
                    });
                }
                if (tables.Count == 0)
                    throw new InvalidOperationException(
                        "ANALYSIS_WORKBOOK_EMPTY: The package has no readable worksheet parts.");
                return AnalysisContract.CreateSnapshot(sourceInstanceId,
                    "excel_openxml", captureRevision,
                    "package_cached_values_formulas_formats",
                    formulaCount > 0
                        ? "cached_formula_values_unverified"
                        : "literal_values",
                    locators, tables);
            }
        }

        private static TableDataset ReadSheet(
            ZipArchiveEntry entry,
            string name,
            IList<string> shared,
            IList<string> formats,
            ref int formulaCount,
            ref int totalCells,
            CancellationToken cancellationToken)
        {
            var document = ReadXml(entry, MaxPartBytes, cancellationToken);
            var table = new TableDataset
            {
                TableId = AnalysisContract.HostId("table",
                    entry.FullName + "|" + name),
                Name = name
            };
            foreach (var element in document.Descendants().Where(item =>
                item.Name.LocalName == "c"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                totalCells++;
                if (totalCells > MaxCells)
                    throw new InvalidOperationException(
                        "ANALYSIS_WORKBOOK_RESOURCE_LIMIT: Typed capture exceeded " +
                        MaxCells.ToString(CultureInfo.InvariantCulture) +
                        " explicit cells.");
                int row;
                int column;
                var reference = (string)element.Attribute("r") ?? string.Empty;
                if (!TryAddress(reference, out row, out column))
                    throw new InvalidOperationException(
                        "ANALYSIS_WORKBOOK_CELL_REFERENCE_INVALID: A worksheet cell has a missing or out-of-range address.");
                table.Rows = Math.Max(table.Rows, row);
                table.Columns = Math.Max(table.Columns, column);
                var type = (string)element.Attribute("t") ?? string.Empty;
                var style = ParseInt((string)element.Attribute("s"), 0);
                var format = style >= 0 && style < formats.Count
                    ? formats[style]
                    : string.Empty;
                var formula = string.Concat(element.Descendants().Where(item =>
                    item.Name.LocalName == "f").Select(item => item.Value));
                if (formula.Length > 0)
                {
                    formula = formula.StartsWith("=", StringComparison.Ordinal)
                        ? formula
                        : "=" + formula;
                    formulaCount++;
                }
                var raw = string.Concat(element.Descendants().Where(item =>
                    item.Name.LocalName == "v").Select(item => item.Value));
                var inline = string.Concat(element.Descendants().Where(item =>
                    item.Name.LocalName == "t").Select(item => item.Value));
                object value = null;
                string display = string.Empty;
                if (type == "s")
                {
                    var index = ParseInt(raw, -1);
                    if (index < 0 || index >= shared.Count)
                    {
                        AddError(table, row, column, reference, type, raw,
                            formula, format);
                        continue;
                    }
                    value = shared[index];
                    display = Convert.ToString(value,
                        CultureInfo.InvariantCulture);
                }
                else if (type == "inlineStr" || type == "str")
                {
                    value = inline.Length > 0 ? inline : raw;
                    display = Convert.ToString(value,
                        CultureInfo.InvariantCulture);
                }
                else if (type == "b")
                {
                    if (raw != "0" && raw != "1" &&
                        !string.Equals(raw, "true",
                            StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(raw, "false",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        AddError(table, row, column, reference, type, raw,
                            formula, format);
                        continue;
                    }
                    value = raw == "1" || string.Equals(raw, "true",
                        StringComparison.OrdinalIgnoreCase);
                    display = (bool)value ? "TRUE" : "FALSE";
                }
                else if (type == "d")
                {
                    DateTime date;
                    if (!DateTime.TryParse(raw, CultureInfo.InvariantCulture,
                        DateTimeStyles.RoundtripKind, out date))
                    {
                        AddError(table, row, column, reference, type, raw,
                            formula, format);
                        continue;
                    }
                    value = date;
                    display = raw;
                }
                else if (type == "e")
                {
                    table.Cells.Add(new DatasetCell
                    {
                        Row = row - 1,
                        Column = column - 1,
                        Reference = reference,
                        ValueType = AnalysisContract.ErrorValue,
                        RawCellType = "e",
                        RawValue = raw,
                        Value = string.Empty,
                        DisplayText = raw,
                        Formula = formula,
                        NumberFormat = format,
                        Status = AnalysisContract.Unresolved
                    });
                    continue;
                }
                else if (type.Length == 0 || type == "n")
                {
                    double number;
                    if (double.TryParse(raw, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out number))
                        value = number;
                    else if (raw.Length > 0)
                    {
                        AddError(table, row, column, reference, type, raw,
                            formula, format);
                        continue;
                    }
                }
                else
                {
                    AddError(table, row, column, reference, type, raw,
                        formula, format);
                    continue;
                }
                var captured = WorkbookTypedCapture.Capture("cell", name,
                    value, formula, format, display, 1, 1, row, column)
                    .Cells[0];
                captured.Row = row - 1;
                captured.Column = column - 1;
                captured.Reference = reference;
                captured.RawCellType = type.Length == 0 ? "n" : type;
                captured.RawValue = type == "inlineStr"
                    ? inline
                    : raw;
                if (formula.Length > 0)
                    captured.Status = AnalysisContract.Unresolved;
                table.Cells.Add(captured);
            }
            return table;
        }

        private static void AddError(
            TableDataset table,
            int row,
            int column,
            string reference,
            string rawCellType,
            string raw,
            string formula,
            string format)
        {
            table.Cells.Add(new DatasetCell
            {
                Row = row - 1,
                Column = column - 1,
                Reference = reference,
                ValueType = AnalysisContract.ErrorValue,
                RawCellType = string.IsNullOrEmpty(rawCellType)
                    ? "n"
                    : rawCellType,
                RawValue = raw,
                Value = string.Empty,
                DisplayText = raw,
                Formula = formula,
                NumberFormat = format,
                Status = AnalysisContract.Unresolved
            });
        }

        private static IList<string> SharedStrings(
            ZipArchive archive,
            CancellationToken cancellationToken)
        {
            var entry = archive.GetEntry("xl/sharedStrings.xml");
            if (entry == null) return new string[0];
            var document = ReadXml(entry, MaxPartBytes,
                cancellationToken);
            return document.Descendants().Where(item =>
                item.Name.LocalName == "si").Select(item =>
                    string.Concat(item.Descendants().Where(text =>
                        text.Name.LocalName == "t").Select(text =>
                            text.Value))).ToList();
        }

        private static IList<string> Formats(
            ZipArchive archive,
            CancellationToken cancellationToken)
        {
            var entry = archive.GetEntry("xl/styles.xml");
            if (entry == null) return new[] { "General" };
            var document = ReadXml(entry, 8 * 1024 * 1024,
                cancellationToken);
            var custom = document.Descendants().Where(item =>
                item.Name.LocalName == "numFmt").ToDictionary(item =>
                    ParseInt((string)item.Attribute("numFmtId"), -1),
                    item => (string)item.Attribute("formatCode") ??
                        string.Empty);
            var result = new List<string>();
            var cellXfs = document.Descendants().FirstOrDefault(item =>
                item.Name.LocalName == "cellXfs");
            foreach (var xf in cellXfs == null
                ? new XElement[0]
                : cellXfs.Elements().Where(item => item.Name.LocalName == "xf"))
            {
                var id = ParseInt((string)xf.Attribute("numFmtId"), 0);
                string value;
                result.Add(custom.TryGetValue(id, out value)
                    ? value
                    : BuiltInFormat(id));
            }
            if (result.Count == 0) result.Add("General");
            return result;
        }

        private static IList<SheetPart> Sheets(
            ZipArchive archive,
            CancellationToken cancellationToken)
        {
            var result = new List<SheetPart>();
            var workbook = archive.GetEntry("xl/workbook.xml");
            var relationships = archive.GetEntry(
                "xl/_rels/workbook.xml.rels");
            if (workbook == null || relationships == null) return result;
            var relationDocument = ReadXml(relationships,
                4 * 1024 * 1024, cancellationToken);
            var targets = new Dictionary<string, string>(
                StringComparer.Ordinal);
            foreach (var relationship in relationDocument.Descendants().Where(
                item => item.Name.LocalName == "Relationship" &&
                    ((string)item.Attribute("Type") ?? string.Empty).EndsWith(
                        "/worksheet", StringComparison.OrdinalIgnoreCase)))
            {
                var id = (string)relationship.Attribute("Id") ?? string.Empty;
                var target = NormalizeWorksheetPath(
                    (string)relationship.Attribute("Target") ?? string.Empty);
                if (id.Length == 0 || target.Length == 0 ||
                    targets.ContainsKey(id))
                    throw new InvalidOperationException(
                        "ANALYSIS_WORKBOOK_METADATA_INVALID: Worksheet relationships are incomplete or duplicated.");
                targets.Add(id, target);
            }
            var workbookDocument = ReadXml(workbook,
                4 * 1024 * 1024, cancellationToken);
            foreach (var sheet in workbookDocument.Descendants().Where(item =>
                item.Name.LocalName == "sheet"))
            {
                var relation = sheet.Attributes().FirstOrDefault(item =>
                    item.Name.LocalName == "id");
                string target;
                if (relation == null || !targets.TryGetValue(relation.Value,
                    out target))
                    throw new InvalidOperationException(
                        "ANALYSIS_WORKBOOK_METADATA_INVALID: A workbook sheet has no worksheet relationship.");
                result.Add(new SheetPart
                {
                    Name = (string)sheet.Attribute("name") ?? target,
                    Path = target
                });
            }
            return result;
        }

        private static string NormalizeWorksheetPath(string target)
        {
            var text = (target ?? string.Empty).Replace('\\', '/');
            var parts = new List<string>();
            if (!text.StartsWith("/", StringComparison.Ordinal) &&
                !text.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                parts.Add("xl");
            foreach (var part in text.Split('/'))
            {
                if (part.Length == 0 || part == ".") continue;
                if (part == "..")
                {
                    if (parts.Count == 0) return string.Empty;
                    parts.RemoveAt(parts.Count - 1);
                }
                else parts.Add(part);
            }
            var path = string.Join("/", parts);
            return path.StartsWith("xl/worksheets/",
                       StringComparison.OrdinalIgnoreCase) &&
                   path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                ? path
                : string.Empty;
        }

        private static XDocument ReadXml(
            ZipArchiveEntry entry,
            long limit,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.Length < 0 || entry.Length > limit)
                throw new InvalidOperationException(
                    "ANALYSIS_WORKBOOK_RESOURCE_LIMIT: Workbook XML part exceeds its bound.");
            using (var stream = entry.Open())
            using (var reader = XmlReader.Create(stream,
                new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    MaxCharactersInDocument = limit,
                    IgnoreComments = true
                }))
                return XDocument.Load(reader, LoadOptions.None);
        }

        private static bool TryAddress(
            string reference,
            out int row,
            out int column)
        {
            row = 0;
            column = 0;
            var index = 0;
            try
            {
                while (index < reference.Length &&
                    char.IsLetter(reference[index]))
                {
                    column = checked(column * 26 +
                        char.ToUpperInvariant(reference[index]) - 'A' + 1);
                    index++;
                }
            }
            catch (OverflowException) { return false; }
            return column > 0 && column <= 16384 &&
                index < reference.Length &&
                int.TryParse(reference.Substring(index),
                    NumberStyles.None, CultureInfo.InvariantCulture,
                    out row) && row > 0 && row <= 1048576;
        }

        private static string Address(int row, int column)
        {
            var letters = string.Empty;
            while (column > 0)
            {
                column--;
                letters = (char)('A' + column % 26) + letters;
                column /= 26;
            }
            return letters + row.ToString(CultureInfo.InvariantCulture);
        }

        private static int ParseInt(string value, int fallback)
        {
            int parsed;
            return int.TryParse(value, NumberStyles.Integer,
                CultureInfo.InvariantCulture, out parsed)
                    ? parsed
                    : fallback;
        }

        private static string BuiltInFormat(int id)
        {
            if (id == 14) return "m/d/yy";
            if (id == 15) return "d-mmm-yy";
            if (id == 16) return "d-mmm";
            if (id == 17) return "mmm-yy";
            if (id == 18) return "h:mm AM/PM";
            if (id == 19) return "h:mm:ss AM/PM";
            if (id == 20) return "h:mm";
            if (id == 21) return "h:mm:ss";
            if (id == 22) return "m/d/yy h:mm";
            if (id == 9) return "0%";
            if (id == 10) return "0.00%";
            return id == 0 ? "General" : "builtin:" +
                id.ToString(CultureInfo.InvariantCulture);
        }

        private sealed class SheetPart
        {
            public string Name { get; set; }
            public string Path { get; set; }
        }
    }
}
