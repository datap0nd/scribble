using System;
using System.Threading;
using System.Windows.Forms;
using Scribble.Configuration;
using Scribble.UI;

namespace Scribble.BrowserHost
{
    // Shows the shared Scribble settings window from the native
    // host so the extension offers the same Settings surface as the
    // Office add-ins. The dialog runs on its own STA thread while
    // the native-messaging call waits; the browser panel refreshes
    // its connection state from the response after the dialog
    // closes.
    internal static class SettingsLauncher
    {
        internal static void ShowSettingsDialog()
        {
            Exception failure = null;
            var thread = new Thread(() =>
            {
                try
                {
                    Application.EnableVisualStyles();
                    var store = new SettingsStore();
                    using (var window = new SettingsWindow(
                        store,
                        store.Load()))
                    {
                        window.StartPosition =
                            FormStartPosition.CenterScreen;
                        window.ShowInTaskbar = true;
                        window.ShowDialog();
                    }
                }
                catch (Exception exception)
                {
                    failure = exception;
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (failure != null)
            {
                throw new InvalidOperationException(
                    "Scribble Settings could not be opened (" +
                    failure.GetType().Name + ").");
            }
        }
    }
}
