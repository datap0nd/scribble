using System;
using System.Runtime.InteropServices;
using Scribble.Security;

namespace Scribble.Outlook
{
    // Opens one unsent Outlook draft window from outside the Outlook
    // add-in process (the browser native host). The draft is only
    // ever displayed for the user's review; this class contains no
    // send capability and must never gain one.
    public static class OutlookDraftLauncher
    {
        public const int MaxRecipientCharacters = 1000;
        public const int MaxSubjectCharacters = 500;

        public static string OpenDraft(
            string to,
            string cc,
            string subject,
            string body,
            Func<string, object> applicationResolver = null)
        {
            var safeBody = TextBoundary.PlainText(
                body,
                TextBoundary.MaxMessageBodyCharacters);
            if (safeBody.Length == 0)
            {
                throw new InvalidOperationException(
                    "The draft body is empty.");
            }

            var safeTo = TextBoundary.SingleLine(
                to,
                MaxRecipientCharacters);
            var safeCc = TextBoundary.SingleLine(
                cc,
                MaxRecipientCharacters);
            var safeSubject = TextBoundary.SingleLine(
                subject,
                MaxSubjectCharacters);

            object application = null;
            object draftItem = null;
            try
            {
                application = applicationResolver == null ? ResolveOutlookApplication() : applicationResolver("Outlook.Application");
                dynamic outlook = application;
                draftItem = outlook.CreateItem(0);
                dynamic mail = draftItem;
                if (safeTo.Length > 0)
                {
                    mail.To = safeTo;
                }

                if (safeCc.Length > 0)
                {
                    mail.CC = safeCc;
                }

                mail.Subject = safeSubject;
                mail.Body = safeBody;
                mail.Display(false);
                return
                    "An unsent Outlook draft window is now open for the " +
                    "user's review. It was not sent and cannot be sent by " +
                    "this tool.";
            }
            finally
            {
                Release(draftItem);
                Release(application);
            }
        }

        private static object ResolveOutlookApplication()
        {
            try
            {
                return Marshal.GetActiveObject("Outlook.Application");
            }
            catch (COMException)
            {
                // Outlook is not running; start a user-visible
                // instance so the draft window has a host.
            }

            var outlookType = Type.GetTypeFromProgID(
                "Outlook.Application");
            if (outlookType == null)
            {
                throw new InvalidOperationException(
                    "Outlook is not installed on this machine.");
            }

            return Activator.CreateInstance(outlookType);
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                try
                {
                    Marshal.ReleaseComObject(value);
                }
                catch
                {
                    // Releasing a COM wrapper twice must never mask
                    // the draft outcome.
                }
            }
        }
    }
}
