using System;
using System.Globalization;
using System.Net;
using System.Runtime.InteropServices;
using Scribble.Chat;
using Scribble.Security;

namespace Scribble.Utilities
{
    internal static class DiagnosticDetails
    {
        public static string ForException(
            Exception exception,
            string fallbackCode = "UNEXPECTED_ERROR")
        {
            var endpointException = exception as AiEndpointException;
            if (endpointException != null)
            {
                return endpointException.ToDiagnosticText();
            }

            var comException = exception as COMException;
            if (comException != null)
            {
                return
                    "[OUTLOOK_COM_" +
                    HResult(comException.HResult) +
                    "] " +
                    TextBoundary.PlainText(comException.Message, 1200);
            }

            var webException = FindWebException(exception);
            if (webException != null)
            {
                return
                    "[NETWORK_" +
                    webException.Status.ToString().ToUpperInvariant() +
                    "] " +
                    TextBoundary.PlainText(webException.Message, 1200) +
                    " HRESULT: " +
                    HResult(webException.HResult) +
                    ".";
            }

            return
                "[" + fallbackCode + "] " +
                TextBoundary.PlainText(exception?.Message, 1200) +
                " HRESULT: " +
                HResult(exception?.HResult ?? 0) +
                ".";
        }

        public static string CodeForLog(Exception exception)
        {
            var endpointException = exception as AiEndpointException;
            if (endpointException != null)
            {
                return endpointException.Code;
            }

            var comException = exception as COMException;
            if (comException != null)
            {
                return "OUTLOOK_COM_" + HResult(comException.HResult);
            }

            var webException = FindWebException(exception);
            if (webException != null)
            {
                return "NETWORK_" +
                    webException.Status.ToString().ToUpperInvariant();
            }

            return "HRESULT_" + HResult(exception?.HResult ?? 0);
        }

        private static WebException FindWebException(Exception exception)
        {
            var current = exception;
            while (current != null)
            {
                var webException = current as WebException;
                if (webException != null)
                {
                    return webException;
                }

                current = current.InnerException;
            }

            return null;
        }

        private static string HResult(int value)
        {
            return "0x" +
                value.ToString("X8", CultureInfo.InvariantCulture);
        }
    }
}
