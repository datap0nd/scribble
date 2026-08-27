using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Scribble.Utilities;

namespace Scribble.UI
{
    // Office has been single-document since 2013: every workbook,
    // presentation, and document opens its own top-level window with
    // its own set of custom task panes, and CreateCTP always targets
    // the ACTIVE window. A single cached pane therefore only ever
    // exists in the first window it was created for - the ribbon
    // button in any newer window (a fresh draft document, say) would
    // silently un-hide the other window's pane and appear to do
    // nothing. This registry keeps one pane per host window instead;
    // PaneMemory keeps the chat history shared per host, so every
    // window shows the same conversation.
    internal sealed class TaskPaneRegistry
    {
        private sealed class Entry
        {
            internal int WindowHandle;
            internal object TaskPane;
            internal OfficeChatPane ChatPane;
        }

        private readonly string _hostKind;
        private readonly List<Entry> _entries = new List<Entry>();

        internal TaskPaneRegistry(string hostKind)
        {
            _hostKind = hostKind ?? string.Empty;
        }

        // Shows the pane belonging to the active window, creating
        // one on that window's first use.
        internal void ShowForActiveWindow(
            object ctpFactory,
            object hostApplication)
        {
            var handle = ActiveWindowHandle(hostApplication);
            for (var index = _entries.Count - 1;
                 index >= 0;
                 index--)
            {
                var entry = _entries[index];
                if (entry.WindowHandle != handle)
                {
                    continue;
                }

                try
                {
                    dynamic existing = entry.TaskPane;
                    existing.Visible = true;
                    return;
                }
                catch
                {
                    // The window was closed and took its pane with
                    // it; drop the stale entry and create a fresh
                    // pane below.
                    Drop(entry);
                    _entries.RemoveAt(index);
                }
            }

            dynamic factory = ctpFactory;
            object taskPane = factory.CreateCTP(
                "Scribble.OfficePane",
                "Scribble",
                Type.Missing);
            dynamic pane = taskPane;
            pane.DockPosition = 2;
            pane.Width = 380;

            var chatPane = pane.ContentControl as OfficeChatPane ??
                OfficeChatPane.LastCreated;
            if (chatPane == null)
            {
                throw new InvalidOperationException(
                    "The sidebar was created but its chat control " +
                    "was unavailable.");
            }

            chatPane.Initialize(_hostKind, hostApplication);
            pane.Visible = true;
            _entries.Add(new Entry
            {
                WindowHandle = handle,
                TaskPane = taskPane,
                ChatPane = chatPane
            });
        }

        internal void CloseAll()
        {
            foreach (var entry in _entries)
            {
                Drop(entry);
            }

            _entries.Clear();
        }

        private void Drop(Entry entry)
        {
            try
            {
                entry.ChatPane?.Shutdown();
                if (entry.TaskPane != null)
                {
                    dynamic pane = entry.TaskPane;
                    pane.Visible = false;
                }
            }
            catch (Exception exception)
            {
                Log.Error(
                    "TaskPaneRegistry." + _hostKind,
                    exception);
            }

            if (entry.TaskPane != null &&
                Marshal.IsComObject(entry.TaskPane))
            {
                Marshal.ReleaseComObject(entry.TaskPane);
            }

            entry.TaskPane = null;
            entry.ChatPane = null;
        }

        // Window handle of the host's active window; 0 when it
        // cannot be read, which degrades to one shared pane.
        private static int ActiveWindowHandle(object hostApplication)
        {
            try
            {
                dynamic application = hostApplication;
                dynamic window = application.ActiveWindow;
                if (window == null)
                {
                    return 0;
                }

                // IDispatch name lookup is case-insensitive, so
                // Hwnd also resolves PowerPoint's HWND property.
                return (int)window.Hwnd;
            }
            catch
            {
                return 0;
            }
        }
    }
}
