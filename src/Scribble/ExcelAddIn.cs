using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Scribble.Interop;
using Scribble.Office;
using Scribble.UI;
using Scribble.Utilities;

namespace Scribble
{
    // Scribble for Excel: the same restrained chat pane, hosted by
    // Excel. The add-in opens the sidebar and can capture an
    // immutable, bounded selection snapshot; every read and write
    // boundary lives in the pane and its tool hosts.
    [ComVisible(true)]
    [Guid("C0ABFA36-9854-434D-A542-DD834938737F")]
    [ProgId("Scribble.ExcelAddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class ExcelAddIn :
        IDTExtensibility2,
        IRibbonExtensibility,
        ICustomTaskPaneConsumer
    {
        private object _excelApplication;
        private object _ctpFactory;
        private readonly TaskPaneRegistry _panes =
            new TaskPaneRegistry("excel");

        public void OnConnection(
            object application,
            ExtConnectMode connectMode,
            object addInInstance,
            ref Array custom)
        {
            _excelApplication = application;
        }

        public void OnDisconnection(
            ExtDisconnectMode removeMode,
            ref Array custom)
        {
            CloseTaskPane();
            _excelApplication = null;
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
            CloseTaskPane();
        }

        public void CTPFactoryAvailable(object ctpFactory)
        {
            _ctpFactory = ctpFactory;
        }

        public string GetCustomUI(string ribbonId)
        {
            // This add-in is registered only under Excel, so any
            // ribbon id it receives is Excel's; exact-id matching
            // would silently hide the button on a casing change.
            if (string.IsNullOrEmpty(ribbonId))
            {
                return null;
            }

            var contextMenus =
                "<contextMenus>" +
                "<contextMenu idMso=\"ContextMenuCell\">" +
                "<button id=\"Scribble.Excel.SendCell\" " +
                "label=\"Send to Scribble\" imageMso=\"ResearchPane\" " +
                "onAction=\"OnSendToScribble\"/>" +
                "</contextMenu>" +
                "<contextMenu idMso=\"ContextMenuRow\">" +
                "<button id=\"Scribble.Excel.SendRow\" " +
                "label=\"Send to Scribble\" imageMso=\"ResearchPane\" " +
                "onAction=\"OnSendToScribble\"/>" +
                "</contextMenu>" +
                "<contextMenu idMso=\"ContextMenuColumn\">" +
                "<button id=\"Scribble.Excel.SendColumn\" " +
                "label=\"Send to Scribble\" imageMso=\"ResearchPane\" " +
                "onAction=\"OnSendToScribble\"/>" +
                "</contextMenu>" +
                "</contextMenus>";

            return
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
                "<ribbon><tabs><tab idMso=\"TabHome\">" +
                "<group id=\"Scribble.Excel.Group\" label=\"Scribble\">" +
                "<button id=\"Scribble.Excel.Open\" label=\"Scribble\" " +
                "size=\"large\" imageMso=\"ResearchPane\" onAction=\"OnOpenChat\" " +
                "screentip=\"Open Scribble\" " +
                "supertip=\"Chat with your workbook. Scribble never saves, deletes, or sends anything.\"/>" +
                "</group></tab></tabs></ribbon>" +
                contextMenus +
                "</customUI>";
        }

        public void OnSendToScribble(object control)
        {
            try
            {
                if (_excelApplication == null || _ctpFactory == null)
                {
                    OnOpenChat(control);
                    return;
                }

                // Capture before CreateCTP: creating a pane pumps
                // messages and can move focus away from the range
                // whose context menu invoked this callback.
                var snapshot = new WorkbookToolHost(
                    _excelApplication).CaptureSelection();
                var pane = _panes.ShowForActiveWindow(
                    _ctpFactory,
                    _excelApplication);
                pane?.AddExcelSelection(snapshot);
            }
            catch (Exception exception)
            {
                Log.Error("ExcelSendToScribble", exception);
                MessageBox.Show(
                    DiagnosticDetails.ForException(
                        exception,
                        "SELECTION_CONTEXT_FAILED"),
                    "Scribble",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        public void OnOpenChat(object control)
        {
            try
            {
                if (_excelApplication == null)
                {
                    MessageBox.Show(
                        "Excel is not ready. Restart Excel and try again.",
                        "Scribble",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (_ctpFactory == null)
                {
                    MessageBox.Show(
                        "Excel has not made the sidebar service available yet. " +
                        "Wait a moment and try again.",
                        "Scribble",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // One pane per Excel window: workbooks opened later
                // get their own pane instead of silently un-hiding
                // the first window's.
                _panes.ShowForActiveWindow(
                    _ctpFactory,
                    _excelApplication);
            }
            catch (Exception exception)
            {
                Log.Error("ExcelOpenChat", exception);
                MessageBox.Show(
                    DiagnosticDetails.ForException(
                        exception,
                        "SIDEBAR_OPEN_FAILED"),
                    "Scribble",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void CloseTaskPane()
        {
            try
            {
                _panes.CloseAll();
            }
            catch (Exception exception)
            {
                Log.Error("ExcelCloseTaskPane", exception);
            }

            Release(_ctpFactory);
            _ctpFactory = null;
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }
    }
}
