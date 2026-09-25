using System.Collections.Generic;

namespace Scribble.Office
{
    // Persisted before-image for a bounded ordinary write_cells operation.
    // The evidence store encrypts this payload; HostData retains only its hash.
    internal sealed class ExcelGridWriteReceipt
    {
        public int Version { get; set; } = 1;
        public string CallId { get; set; }
        public string WorkbookName { get; set; }
        public string WorkbookFullName { get; set; }
        public string SheetName { get; set; }
        public string StartCell { get; set; }
        public List<ExcelGridCellReceipt> Cells { get; set; } =
            new List<ExcelGridCellReceipt>();
    }

    internal sealed class ExcelGridCellReceipt
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public bool HasFormula { get; set; }
        public string Formula { get; set; }
        public string ValueKind { get; set; }
        public string Value { get; set; }
        public string NumberFormat { get; set; }
        public string Planned { get; set; }
    }
}
