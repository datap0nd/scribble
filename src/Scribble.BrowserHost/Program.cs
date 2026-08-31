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
