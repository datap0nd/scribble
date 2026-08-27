using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Scribble.Interop;
using Scribble.UI;
using Scribble.Utilities;

namespace Scribble
{
    // Scribble for PowerPoint: the same restrained chat pane, hosted by
    // PowerPoint. The add-in itself holds no capability beyond
    // opening the sidebar; every read and draft boundary lives in
    // the pane and its tool hosts.
    [ComVisible(true)]
    [Guid("69FAE812-274F-43F8-8F45-1B4EB22B5248")]
    [ProgId("Scribble.PowerPointAddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class PowerPointAddIn :
        IDTExtensibility2,
        IRibbonExtensibility,
        ICustomTaskPaneConsumer
    {
        private object _powerPointApplication;
        private object _ctpFactory;
        private readonly TaskPaneRegistry _panes =
            new TaskPaneRegistry("powerpoint");

        public void OnConnection(
            object application,
            ExtConnectMode connectMode,
            object addInInstance,
            ref Array custom)
        {
            _powerPointApplication = application;
        }

        public void OnDisconnection(
            ExtDisconnectMode removeMode,
            ref Array custom)
        {
            CloseTaskPane();
            _powerPointApplication = null;
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
            // This add-in is registered only under PowerPoint, so
            // any ribbon id it receives is PowerPoint's. The id's
            // exact casing has varied across Office versions
            // ("Microsoft.Powerpoint.Presentation" vs
            // "Microsoft.PowerPoint.Presentation"), and a mismatch
            // silently hides the button.
            if (string.IsNullOrEmpty(ribbonId))
            {
                return null;
            }

            return
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
                "<ribbon><tabs><tab idMso=\"TabHome\">" +
                "<group id=\"Scribble.PowerPoint.Group\" label=\"Scribble\">" +
                "<button id=\"Scribble.PowerPoint.Open\" label=\"Scribble\" " +
                "size=\"large\" imageMso=\"ResearchPane\" onAction=\"OnOpenChat\" " +
                "screentip=\"Open Scribble\" " +
                "supertip=\"Chat with your presentation. Scribble never saves, deletes, or sends anything.\"/>" +
                "</group></tab></tabs></ribbon>" +
                "</customUI>";
        }

        public void OnOpenChat(object control)
        {
            try
            {
                if (_powerPointApplication == null)
                {
                    MessageBox.Show(
                        "PowerPoint is not ready. Restart PowerPoint and try again.",
                        "Scribble",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (_ctpFactory == null)
                {
                    MessageBox.Show(
                        "PowerPoint has not made the sidebar service available yet. " +
                        "Wait a moment and try again.",
                        "Scribble",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // One pane per PowerPoint window: presentations
                // opened later get their own pane instead of
                // silently un-hiding the first window's.
                _panes.ShowForActiveWindow(
                    _ctpFactory,
                    _powerPointApplication);
            }
            catch (Exception exception)
            {
                Log.Error("PowerPointOpenChat", exception);
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
                Log.Error("PowerPointCloseTaskPane", exception);
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
