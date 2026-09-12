using System;

namespace Scribble.BrowserHost
{
    internal static class Program
    {
        internal const string AllowedOrigin =
            "chrome-extension://olkepladbgkfkhlglooilnmalckpdada/";

        [STAThread]
        private static int Main(string[] args)
        {
            // Explicit local operator launch, separate from native messaging and Office lifetimes.
            if (args != null &&
                (args.Length == 1 || args.Length == 2) &&
                args[0] == "--test-lab-suite") {
                System.Windows.Forms.Application.EnableVisualStyles();
                var launchNonce = args.Length == 2
                    ? args[1]
                    : Guid.NewGuid().ToString("N");
                try { using (var lab = new Scribble.Testing.TestLabSuiteWindow(launchNonce)) lab.ShowDialog(); return 0; }
                catch (Exception error) {
                    System.Windows.Forms.MessageBox.Show("Test Lab could not start.\r\n\r\n" + error, "Scribble Test Lab startup error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return 1;
                }
            }
            if (args != null &&
                args.Length > 0 &&
                string.Equals(
                    args[0],
                    "--setup",
                    StringComparison.OrdinalIgnoreCase))
            {
                return BrowserSetup.Run(
                    args.Length > 1
                        ? args[1]
                        : "auto");
            }

            // Chrome already enforces allowed_origins from
            // the native-host manifest. Check the supplied origin a
            // second time so a differently registered extension can
            // never use this executable by accident.
            if (args == null ||
                args.Length == 0 ||
                !string.Equals(
                    args[0],
                    AllowedOrigin,
                    StringComparison.Ordinal))
            {
                Console.Error.WriteLine(
                    "Scribble refused an unrecognized extension origin.");
                return 2;
            }

            return NativeMessageProtocol.Run(
                Console.OpenStandardInput(),
                Console.OpenStandardOutput());
        }
    }
}
