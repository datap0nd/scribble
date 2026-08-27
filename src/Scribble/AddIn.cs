using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Scribble.Interop;
using Scribble.UI;
using Scribble.Utilities;

namespace Scribble
{
    [ComVisible(true)]
    [Guid("0D6E56F9-BE2D-4B94-B5E4-4C2DB0FD13E7")]
    [ProgId("Scribble.AddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class AddIn :
        IDTExtensibility2,
        IRibbonExtensibility,
        ICustomTaskPaneConsumer
    {
        private object _outlookApplication;
        private object _ctpFactory;
        private object _taskPane;
        private ChatPane _chatPane;

        public void OnConnection(
            object application,
            ExtConnectMode connectMode,
            object addInInstance,
            ref Array custom)
        {
            _outlookApplication = application;
        }

        public void OnDisconnection(
            ExtDisconnectMode removeMode,
            ref Array custom)
        {
            CloseTaskPane();
            _outlookApplication = null;
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
            var tabId = GetTabId(ribbonId);
            if (tabId == null)
            {
                return null;
            }

            var contextMenu = string.Equals(
                ribbonId,
                "Microsoft.Outlook.Explorer",
                StringComparison.Ordinal)
                ? "<contextMenus>" +
                  "<contextMenu idMso=\"ContextMenuMailItem\">" +
                  "<button id=\"Scribble.SendToAi\" " +
                  "label=\"Send to Scribble\" imageMso=\"ResearchPane\" " +
                  "onAction=\"OnSendToAi\"/>" +
                  "<button id=\"Scribble.SuggestReply\" " +
                  "label=\"Scribble - Suggest a response\" " +
                  "imageMso=\"Reply\" " +
                  "onAction=\"OnSuggestReply\"/>" +
                  "</contextMenu></contextMenus>"
                : string.Empty;

            return
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
                "<ribbon><tabs><tab idMso=\"" + tabId + "\">" +
                "<group id=\"Scribble.Group\" label=\"Scribble\">" +
                "<button id=\"Scribble.Open\" label=\"Scribble\" " +
                "size=\"large\" imageMso=\"ResearchPane\" onAction=\"OnOpenChat\" " +
                "screentip=\"Open Scribble\" " +
                "supertip=\"Chat with your mailbox and refine one unsent Outlook draft.\"/>" +
                "</group></tab></tabs></ribbon>" +
                contextMenu +
                "</customUI>";
        }

        public void OnSendToAi(object control)
        {
            OpenChat(control, false);
            _chatPane?.AddActiveSelection();
        }

        public void OnSuggestReply(object control)
        {
            OpenChat(control, true);
            _chatPane?.BeginSuggestResponse();
        }

        public void OnOpenChat(object control)
        {
            // Opening the pane never captures whatever happens to be
            // selected: context is always something the user adds
            // deliberately, by dragging emails onto the pane, using
            // Add email, or right-clicking Send to Scribble.
            OpenChat(control, false);
        }

        private void OpenChat(
            object control,
            bool refreshSelection)
        {
            try
            {
                if (_outlookApplication == null)
                {
                    MessageBox.Show(
                        "Outlook is not ready. Restart Outlook and try again.",
                        "Scribble",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (_ctpFactory == null)
                {
                    MessageBox.Show(
                        "Outlook has not made the sidebar service available yet. " +
                        "Wait a moment and try again.",
                        "Scribble",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                if (_taskPane == null)
                {
                    object parentWindow = GetTaskPaneParent(control);
                    dynamic factory = _ctpFactory;
                    _taskPane = factory.CreateCTP(
                        "Scribble.ChatPane",
                        "Scribble",
                        parentWindow ?? Type.Missing);

                    dynamic pane = _taskPane;
                    pane.DockPosition = 2;
                    pane.Width = 380;

                    _chatPane = pane.ContentControl as ChatPane ??
                        ChatPane.LastCreated;
                    if (_chatPane == null)
                    {
                        throw new InvalidOperationException(
                            "Outlook created the sidebar but its chat control was unavailable.");
                    }

                    _chatPane.Initialize(
                        _outlookApplication,
                        refreshSelection);
                    pane.Visible = true;
                }
                else
                {
                    if (refreshSelection)
                    {
                        _chatPane?.RefreshSelectedMessage();
                    }
                    dynamic pane = _taskPane;
                    pane.Visible = true;
                }
            }
            catch (Exception exception)
            {
                Log.Error("OnOpenChat", exception);
                MessageBox.Show(
                    DiagnosticDetails.ForException(
                        exception,
                        "SIDEBAR_OPEN_FAILED"),
                    "Scribble",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private static string GetTabId(string ribbonId)
        {
            if (string.Equals(
                ribbonId,
                "Microsoft.Outlook.Explorer",
                StringComparison.Ordinal))
            {
                return "TabMail";
            }

            if (string.Equals(
                ribbonId,
                "Microsoft.Outlook.Mail.Read",
                StringComparison.Ordinal))
            {
                return "TabReadMessage";
            }

            return null;
        }

        private static object GetRibbonContext(object control)
        {
            try
            {
                dynamic ribbonControl = control;
                return ribbonControl?.Context;
            }
            catch
            {
                return null;
            }
        }

        private object GetTaskPaneParent(object control)
        {
            var context = GetRibbonContext(control);
            try
            {
                dynamic window = context;
                object windowHandle = window.HWND;
                return windowHandle == null
                    ? null
                    : context;
            }
            catch
            {
                try
                {
                    dynamic application = _outlookApplication;
                    return application.ActiveExplorer();
                }
                catch
                {
                    return null;
                }
            }
        }

        private void CloseTaskPane()
        {
            try
            {
                _chatPane?.Shutdown();
                if (_taskPane != null)
                {
                    dynamic pane = _taskPane;
                    pane.Visible = false;
                }
            }
            catch (Exception exception)
            {
                Log.Error("CloseTaskPane", exception);
            }

            Release(_taskPane);
            Release(_ctpFactory);
            _taskPane = null;
            _ctpFactory = null;
            _chatPane = null;
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
