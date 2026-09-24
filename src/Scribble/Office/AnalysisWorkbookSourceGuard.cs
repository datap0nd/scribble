using System;
using System.Globalization;
using System.Linq;
using Scribble.Chat;

namespace Scribble.Office
{
    // Re-read the exact bound Excel range before a pilot write. A saved-file
    // timestamp or an unsaved-workbook name cannot establish cell freshness.
    public static class AnalysisWorkbookSourceGuard
    {
        public static void Validate(object excelApplication,
            AnalysisArtifact artifact)
        {
            AnalysisContract.Serialize(artifact);
            if (excelApplication == null || artifact.Snapshots.Count != 1 ||
                artifact.Snapshots[0].Tables.Count != 1)
                throw new InvalidOperationException(
                    "ANALYSIS_SOURCE_UNAVAILABLE");
            var snapshot = artifact.Snapshots[0];
            var table = snapshot.Tables[0];
            var locator = snapshot.Locators.SingleOrDefault(item =>
                item.Kind == "excel_range" &&
                item.SourceInstanceId == snapshot.SourceInstanceId &&
                item.WorksheetIdentity == table.Name);
            if (locator == null || string.IsNullOrWhiteSpace(locator.Range))
                throw new InvalidOperationException(
                    "ANALYSIS_SOURCE_UNAVAILABLE");
            try
            {
                var binding = OfficeTaskBinding.Capture("excel",
                    excelApplication);
                if (binding == null ||
                    binding.Id != snapshot.SourceInstanceId)
                    throw new InvalidOperationException(
                        "ANALYSIS_SOURCE_CHANGED: workbook binding");
                dynamic application = excelApplication;
                dynamic workbook = application.ActiveWorkbook;
                dynamic sheet = workbook.Worksheets[table.Name];
                dynamic range = sheet.Range(locator.Range);
                var address = Convert.ToString(
                    range.Address(false, false),
                    CultureInfo.InvariantCulture);
                if (!string.Equals(address, locator.Range,
                        StringComparison.OrdinalIgnoreCase) ||
                    (int)range.Rows.Count != table.Rows ||
                    (int)range.Columns.Count != table.Columns ||
                    snapshot.CaptureRevision != binding.Fingerprint + "|" +
                        table.Name + "|" + address)
                    throw new InvalidOperationException(
                        "ANALYSIS_SOURCE_CHANGED: range or revision");
                object values = range.Value2;
                object formulas = range.Formula;
                object formats = WorkbookTypedCapture
                    .ResolveMixedNumberFormats((object)range.NumberFormat,
                        table.Rows, table.Columns,
                        column => (object)range.Columns[column + 1]
                            .NumberFormat,
                        (row, column) => (object)range.Cells[row + 1,
                            column + 1].NumberFormat,
                        WorkbookToolHost.MaxTypedMetadataCells);
                if (formats == null)
                    throw new InvalidOperationException(
                        "ANALYSIS_SOURCE_CHANGED: mixed formats incomplete");
                var current = WorkbookTypedCapture.Capture(table.TableId,
                    table.Name, values, formulas, formats, null,
                    table.Rows, table.Columns, (int)range.Row,
                    (int)range.Column);
                var metadata = current.Cells.Where(cell =>
                    !string.IsNullOrEmpty(cell.Formula) ||
                    (!string.IsNullOrEmpty(cell.NumberFormat) &&
                     !string.Equals(cell.NumberFormat, "General",
                         StringComparison.OrdinalIgnoreCase)) ||
                    cell.ValueType == AnalysisContract.DateValue ||
                    cell.ValueType == AnalysisContract.BooleanValue ||
                    cell.ValueType == AnalysisContract.ErrorValue).ToArray();
                if (metadata.Length > WorkbookToolHost.MaxTypedMetadataCells)
                    throw new InvalidOperationException(
                        "ANALYSIS_SOURCE_CHANGED: metadata bound exceeded");
                foreach (var cell in current.Cells.Where(item =>
                    !string.IsNullOrEmpty(item.Formula)))
                    cell.Status = AnalysisContract.Unresolved;
                foreach (var cell in metadata)
                    cell.DisplayText = Convert.ToString(
                        range.Cells[cell.Row + 1, cell.Column + 1].Text,
                        CultureInfo.InvariantCulture) ?? cell.DisplayText;
                var fresh = AnalysisContract.CreateSnapshot(
                    snapshot.SourceInstanceId, snapshot.SourceType,
                    snapshot.CaptureRevision, snapshot.Coverage,
                    snapshot.CalculationState, snapshot.Locators,
                    new[] { current });
                if (fresh.ContentHash != snapshot.ContentHash)
                    throw new InvalidOperationException(
                        "ANALYSIS_SOURCE_CHANGED: cell content or format");
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "ANALYSIS_SOURCE_UNAVAILABLE: The bound Excel range could not be re-read.",
                    exception);
            }
        }
    }
}
