using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Scribble.Interop;
using Scribble.UI;
using Scribble.Utilities;

namespace Scribble
{
    // Scribble for Word: the same restrained chat pane, hosted by
    // Word. The add-in itself holds no capability beyond opening
    // the sidebar; every read and draft boundary lives in the pane
    // and its tool hosts.
    [ComVisible(true)]
    [Guid("B49E9DB7-0C40-46A8-80A3-547626FE5331")]
    [ProgId("Scribble.WordAddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class WordAddIn :
        IDTExtensibility2,
        IRibbonExtensibility,
        ICustomTaskPaneConsumer
    {
        private object _wordApplication;
        private object _ctpFactory;
        private readonly TaskPaneRegistry _panes =
            new TaskPaneRegistry("word");

        public void OnConnection(
            object application,
            ExtConnectMode connectMode,
            object addInInstance,
            ref Array custom)
        {
            _wordApplication = application;
        }

        public void OnDisconnection(
            ExtDisconnectMode removeMode,
            ref Array custom)
        {
            CloseTaskPane();
            _wordApplication = null;
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
            // This add-in is registered only under Word, so any
            // ribbon id it receives is Word's; exact-id matching
            // would silently hide the button on a casing change.
            if (string.IsNullOrEmpty(ribbonId))
            {
                return null;
            }

            return
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
                "<ribbon><tabs><tab idMso=\"TabHome\">" +
                "<group id=\"Scribble.Word.Group\" label=\"Scribble\">" +
                "<button id=\"Scribble.Word.Open\" label=\"Scribble\" " +
                "size=\"large\" imageMso=\"ResearchPane\" onAction=\"OnOpenChat\" " +
                "screentip=\"Open Scribble\" " +
                "supertip=\"Chat with your document. Scribble never saves, deletes, or sends anything.\"/>" +
                "</group></tab></tabs></ribbon>" +
                "</customUI>";
        }

        public void OnOpenChat(object control)
        {
            try
            {
                if (_wordApplication == null)
                {
                    MessageBox.Show(
                        "Word is not ready. Restart Word and try again.",
                        "Scribble",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (_ctpFactory == null)
                {
                    MessageBox.Show(
                        "Word has not made the sidebar service available yet. " +
                        "Wait a moment and try again.",
                        "Scribble",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // One pane per Word window: documents opened or
                // created later (like a fresh draft) get their own
                // pane instead of silently un-hiding the first
                // window's.
                _panes.ShowForActiveWindow(
                    _ctpFactory,
                    _wordApplication);
            }
            catch (Exception exception)
            {
                Log.Error("WordOpenChat", exception);
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
                Log.Error("WordCloseTaskPane", exception);
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
