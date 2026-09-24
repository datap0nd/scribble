using System;

namespace Scribble.Office
{
    // A normal Excel draft keeps the workbook/worksheet that was active when
    // the user submitted the request. Later focus changes cannot redirect it.
    internal sealed class ExcelDraftBinding
    {
        private readonly object _workbook;
        private readonly object _sheet;
        private readonly string _workbookName;
        private readonly string _workbookFullName;
        private readonly string _sheetName;

        private ExcelDraftBinding(object workbook, object sheet,
            string workbookName, string workbookFullName, string sheetName)
        {
            _workbook = workbook;
            _sheet = sheet;
            _workbookName = workbookName;
            _workbookFullName = workbookFullName;
            _sheetName = sheetName;
        }

        internal static ExcelDraftBinding Capture(object application)
        {
            dynamic app = application;
            dynamic workbook = app.ActiveWorkbook;
            if (workbook == null)
                return new ExcelDraftBinding(null, null, null, null, null);
            dynamic sheet = workbook.ActiveSheet;
            return new ExcelDraftBinding((object)workbook, (object)sheet,
                Convert.ToString(workbook.Name),
                Convert.ToString(workbook.FullName),
                sheet == null ? null : Convert.ToString(sheet.Name));
        }

        internal object Workbook()
        {
            if (_workbook == null) return null;
            try
            {
                dynamic workbook = _workbook;
                if (Convert.ToString(workbook.Name) != _workbookName ||
                    Convert.ToString(workbook.FullName) !=
                        _workbookFullName)
                    throw new InvalidOperationException(
                        "EXCEL_DRAFT_TARGET_CHANGED: The captured workbook was renamed or saved under another path.");
                return _workbook;
            }
            catch (System.Runtime.InteropServices.COMException error)
            {
                throw new InvalidOperationException(
                    "EXCEL_DRAFT_TARGET_UNAVAILABLE: The captured workbook is no longer open.",
                    error);
            }
        }

        internal object Sheet()
        {
            if (_sheet == null || Workbook() == null)
                throw new InvalidOperationException(
                    "EXCEL_DRAFT_SHEET_UNAVAILABLE: The request began without an active worksheet.");
            try
            {
                dynamic sheet = _sheet;
                if (Convert.ToString(sheet.Name) != _sheetName ||
                    Convert.ToString(sheet.Parent.FullName) !=
                        _workbookFullName)
                    throw new InvalidOperationException(
                        "EXCEL_DRAFT_TARGET_CHANGED: The captured worksheet changed.");
                return _sheet;
            }
            catch (System.Runtime.InteropServices.COMException error)
            {
                throw new InvalidOperationException(
                    "EXCEL_DRAFT_SHEET_UNAVAILABLE: The captured worksheet is no longer open.",
                    error);
            }
        }
    }
}
