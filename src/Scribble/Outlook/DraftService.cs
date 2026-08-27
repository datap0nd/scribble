using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Scribble.Security;

namespace Scribble.Outlook
{
    internal sealed class DraftService
    {
        private readonly object _outlookApplication;

        internal DraftService(object outlookApplication)
        {
            _outlookApplication = outlookApplication ??
                throw new ArgumentNullException(nameof(outlookApplication));
        }

        internal DraftSession CreateReplyDraft(
            MessageSnapshot source,
            string body,
            IReadOnlyList<string> boldPhrases)
        {
            if (source == null || !source.CanReply)
            {
                throw new InvalidOperationException(
                    "The resolved source email is not available for reply.");
            }

            object session = null;
            object original = null;
            object reply = null;

            try
            {
                dynamic application = _outlookApplication;
                session = application.Session;
                dynamic outlookSession = session;
                original = source.StoreId.Length > 0
                    ? outlookSession.GetItemFromID(
                        source.EntryId,
                        source.StoreId)
                    : outlookSession.GetItemFromID(source.EntryId);
                dynamic originalMail = original;
                reply = CreateReply(
                    application,
                    originalMail,
                    source);
                dynamic replyMail = reply;
                var quotedHtml = SafeString(
                    () => replyMail.HTMLBody);
                var draft = new DraftSession(
                    reply,
                    "reply",
                    quotedHtml);
                draft.Update(
                    body,
                    boldPhrases,
                    null,
                    null,
                    null);
                reply = null;
                return draft;
            }
            finally
            {
                Release(reply);
                Release(original);
                Release(session);
            }
        }

        internal DraftSession CreateNewDraft(
            string body,
            IReadOnlyList<string> boldPhrases,
            string subject,
            string to,
            string cc)
        {
            object draftItem = null;
            try
            {
                dynamic application = _outlookApplication;
                draftItem = application.CreateItem(0);
                var draft = new DraftSession(
                    draftItem,
                    "new",
                    string.Empty);
                draft.Update(
                    body,
                    boldPhrases,
                    subject,
                    to,
                    cc);
                draftItem = null;
                return draft;
            }
            finally
            {
                Release(draftItem);
            }
        }

        // Outlook refuses MailItem.Reply() while a reply to the same
        // message sits docked inline in the reading pane ("This
        // method can't be used with an inline response mail item").
        // In that case the docked inline reply is popped out into
        // its own window and becomes the linked draft itself, so the
        // request succeeds instead of failing.
        private static object CreateReply(
            dynamic application,
            dynamic originalMail,
            MessageSnapshot source)
        {
            try
            {
                return originalMail.Reply();
            }
            catch (System.Runtime.InteropServices
                .COMException exception)
            {
                var takeover = TryTakeOverInlineResponse(
                    application,
                    source);
                if (takeover != null)
                {
                    return takeover;
                }

                throw new InvalidOperationException(
                    "Outlook is showing an inline reply in the " +
                    "reading pane, which blocks automated reply " +
                    "drafts. Pop that reply out into its own " +
                    "window (or discard it), or turn on 'Open " +
                    "replies and forwards in a new window' under " +
                    "File > Options > Mail, then ask again.",
                    exception);
            }
        }

        // Returns the reading-pane inline reply as a normal draft
        // when one is open for the same conversation: Display() pops
        // it out into an inspector window, after which it behaves as
        // any unsent reply draft. Returns null when there is no
        // matching inline response to take over.
        private static object TryTakeOverInlineResponse(
            dynamic application,
            MessageSnapshot source)
        {
            try
            {
                dynamic explorer = application.ActiveExplorer();
                if (explorer == null)
                {
                    return null;
                }

                dynamic inline = explorer.ActiveInlineResponse;
                if (inline == null)
                {
                    return null;
                }

                var topic = Convert.ToString(
                    inline.ConversationTopic) ?? string.Empty;
                var subject = source?.Subject ?? string.Empty;
                if (topic.Length > 0 &&
                    subject.Length > 0 &&
                    subject.IndexOf(
                        topic,
                        StringComparison.OrdinalIgnoreCase) < 0 &&
                    topic.IndexOf(
                        subject,
                        StringComparison.OrdinalIgnoreCase) < 0)
                {
                    // The inline reply belongs to a different
                    // conversation; leave it alone.
                    return null;
                }

                inline.Display();
                return inline;
            }
            catch
            {
                return null;
            }
        }

        private static string SafeString(Func<object> reader)
        {
            try
            {
                return Convert.ToString(reader()) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }
    }

    internal sealed class DraftSession : IDisposable
    {
        private object _mailItem;
        private readonly string _kind;
        private readonly string _quotedHtml;
        private string _body = string.Empty;

        internal DraftSession(
            object mailItem,
            string kind,
            string quotedHtml)
        {
            _mailItem = mailItem ??
                throw new ArgumentNullException(nameof(mailItem));
            _kind = kind ?? string.Empty;
            _quotedHtml = quotedHtml ?? string.Empty;
        }

        internal DraftReference Reference
        {
            get
            {
                EnsureAvailable();
                dynamic mail = _mailItem;
                return new DraftReference(
                    _kind,
                    SafeString(() => mail.Subject),
                    SafeString(() => mail.To),
                    SafeString(() => mail.CC),
                    _body);
            }
        }

        internal void Update(
            string body,
            IReadOnlyList<string> boldPhrases,
            string subject,
            string to,
            string cc)
        {
            EnsureAvailable();
            var boundedBody = TextBoundary.PlainText(
                body,
                TextBoundary.MaxAssistantCharacters);
            var content = SafeDraftHtml.FormatContent(
                boundedBody,
                boldPhrases);
            dynamic mail = _mailItem;

            if (subject != null)
            {
                var formattedSubject = SafeModelText.Format(
                    subject,
                    255);
                mail.Subject = TextBoundary.SingleLine(
                    formattedSubject.PlainText,
                    255);
            }

            if (to != null)
            {
                mail.To = TextBoundary.SingleLine(to, 2000);
            }

            if (cc != null)
            {
                mail.CC = TextBoundary.SingleLine(cc, 2000);
            }

            mail.HTMLBody = _kind == "reply" &&
                _quotedHtml.Length > 0
                ? content.Html + "<br><br>" + _quotedHtml
                : content.Html;
            mail.Save();
            mail.Display(false);
            _body = content.PlainText;
        }

        // Attaches an existing file the user already has on disk to
        // the unsent draft. Reading the file for attachment never
        // modifies it; the draft still only opens for review.
        internal void AttachFile(string path)
        {
            EnsureAvailable();
            if (string.IsNullOrWhiteSpace(path) ||
                !File.Exists(path))
            {
                return;
            }

            dynamic mail = _mailItem;
            mail.Attachments.Add(path);
            mail.Save();
        }

        public void Dispose()
        {
            var item = _mailItem;
            _mailItem = null;
            if (item != null && Marshal.IsComObject(item))
            {
                Marshal.ReleaseComObject(item);
            }
        }

        private void EnsureAvailable()
        {
            if (_mailItem == null)
            {
                throw new InvalidOperationException(
                    "The linked Outlook draft is no longer available.");
            }
        }

        private static string SafeString(Func<object> reader)
        {
            try
            {
                return Convert.ToString(reader()) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
