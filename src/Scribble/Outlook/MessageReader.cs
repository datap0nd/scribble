using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Scribble.Security;

namespace Scribble.Outlook
{
    public sealed class MessageReader
    {
        private readonly object _outlookApplication;

        public MessageReader(object outlookApplication)
        {
            _outlookApplication = outlookApplication ??
                throw new ArgumentNullException(nameof(outlookApplication));
        }

        public MessageSnapshot CaptureCurrent()
        {
            object item = null;
            object inspector = null;
            object explorer = null;
            object selection = null;

            try
            {
                dynamic application = _outlookApplication;
                inspector = application.ActiveInspector();
                if (inspector != null)
                {
                    dynamic activeInspector = inspector;
                    item = activeInspector.CurrentItem;
                }

                if (item == null)
                {
                    explorer = application.ActiveExplorer();
                    if (explorer != null)
                    {
                        dynamic activeExplorer = explorer;
                        selection = activeExplorer.Selection;
                        dynamic currentSelection = selection;
                        if (currentSelection != null && currentSelection.Count > 0)
                        {
                            item = currentSelection.Item(1);
                        }
                    }
                }

                if (item == null)
                {
                    throw new InvalidOperationException(
                        "Select or open an email in Outlook first.");
                }

                dynamic mail = item;
                var messageClass = SafeString(() => mail.MessageClass);
                if (!IsReadableItemClass(messageClass))
                {
                    throw new InvalidOperationException(
                        "The selected Outlook item is not an email, meeting invite, or appointment.");
                }

                return CaptureItem(item);
            }
            finally
            {
                Release(selection);
                Release(explorer);
                Release(inspector);
                Release(item);
            }
        }

        public MessageSnapshot CaptureById(
            string entryId,
            string storeId)
        {
            if (string.IsNullOrWhiteSpace(entryId))
            {
                throw new ArgumentException(
                    "A message entry ID is required.",
                    nameof(entryId));
            }

            object session = null;
            object item = null;
            try
            {
                dynamic application = _outlookApplication;
                session = application.Session;
                dynamic outlookSession = session;
                item = string.IsNullOrWhiteSpace(storeId)
                    ? outlookSession.GetItemFromID(entryId)
                    : outlookSession.GetItemFromID(entryId, storeId);
                return CaptureItem(item);
            }
            finally
            {
                Release(item);
                Release(session);
            }
        }

        public IReadOnlyList<MessageSnapshot> CaptureActiveSelectionMany()
        {
            object explorer = null;
            object selection = null;
            try
            {
                dynamic application = _outlookApplication;
                explorer = application.ActiveExplorer();
                if (explorer == null)
                {
                    throw new InvalidOperationException(
                        "Open an Outlook mailbox view and select one to ten emails.");
                }

                dynamic activeExplorer = explorer;
                selection = activeExplorer.Selection;
                return CaptureSelectionMany(selection);
            }
            finally
            {
                Release(selection);
                Release(explorer);
            }
        }

        public MessageSnapshot CaptureSelection(object selection)
        {
            var messages = CaptureSelectionMany(selection);
            if (messages.Count != 1)
            {
                throw new InvalidOperationException(
                    "Select exactly one email before using this action.");
            }

            return messages[0];
        }

        public IReadOnlyList<MessageSnapshot> CaptureSelectionMany(
            object context)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            object selection = context;
            var releaseSelection = false;
            try
            {
                int count;
                try
                {
                    dynamic selectedItems = selection;
                    count = Convert.ToInt32(selectedItems.Count);
                }
                catch
                {
                    dynamic window = context;
                    selection = window.Selection;
                    releaseSelection = !ReferenceEquals(
                        selection,
                        context);
                    dynamic selectedItems = selection;
                    count = Convert.ToInt32(selectedItems.Count);
                }

                if (count < 1)
                {
                    throw new InvalidOperationException(
                        "Select at least one email before using Send to Scribble.");
                }

                if (count > MailboxWorkingSet.MaxMessages)
                {
                    throw new InvalidOperationException(
                        "Select no more than ten emails before using Send to Scribble.");
                }

                var messages = new List<MessageSnapshot>(count);
                dynamic items = selection;
                for (var index = 1; index <= count; index++)
                {
                    object item = null;
                    try
                    {
                        item = items.Item(index);
                        messages.Add(CaptureItem(item));
                    }
                    finally
                    {
                        Release(item);
                    }
                }

                return MailboxWorkingSet.Normalize(messages);
            }
            finally
            {
                if (releaseSelection)
                {
                    Release(selection);
                }
            }
        }

        internal static MessageSnapshot CaptureItem(object item)
        {
            if (item == null)
            {
                throw new ArgumentNullException(nameof(item));
            }

            object parent = null;
            try
            {
                dynamic mail = item;
                var messageClass = SafeString(() => mail.MessageClass);
                if (!IsReadableItemClass(messageClass))
                {
                    throw new InvalidOperationException(
                        "The Outlook item is not an email, meeting invite, or appointment.");
                }

                var isAppointment = messageClass.StartsWith(
                    "IPM.Appointment",
                    StringComparison.OrdinalIgnoreCase);
                var isMeeting = messageClass.StartsWith(
                    "IPM.Schedule.Meeting",
                    StringComparison.OrdinalIgnoreCase);

                parent = SafeObject(() => mail.Parent);
                var storeId = string.Empty;
                if (parent != null)
                {
                    dynamic folder = parent;
                    storeId = SafeString(() => folder.StoreID);
                }

                var receivedAt =
                    SafeDateTime(() => mail.ReceivedTime) ??
                    SafeDateTime(() => mail.SentOn) ??
                    SafeDateTime(() => mail.Start) ??
                    SafeDateTime(() => mail.CreationTime);

                var sender = isAppointment
                    ? SafeString(() => mail.Organizer)
                    : BuildSender(mail);
                var meetingDetails = isMeeting || isAppointment
                    ? BuildMeetingDetails(mail, isMeeting)
                    : string.Empty;

                return new MessageSnapshot(
                    SafeString(() => mail.EntryID),
                    storeId,
                    TextBoundary.PlainText(
                        SafeString(() => mail.Subject),
                        1000),
                    TextBoundary.PlainText(sender, 1000),
                    TextBoundary.PlainText(
                        BuildRecipients(mail, isAppointment),
                        2000),
                    receivedAt,
                    TextBoundary.PlainText(
                        meetingDetails +
                        SafeString(() => mail.Body),
                        ContextScale.Scaled(
                            TextBoundary
                                .MaxMessageBodyCharacters)),
                    CaptureAttachmentNames(mail),
                    CountRemoteImages(
                        SafeString(() => mail.HTMLBody)));
            }
            finally
            {
                Release(parent);
            }
        }

        private static bool IsReadableItemClass(string messageClass)
        {
            var value = messageClass ?? string.Empty;
            return value.StartsWith(
                       "IPM.Note",
                       StringComparison.OrdinalIgnoreCase) ||
                   value.StartsWith(
                       "IPM.Schedule.Meeting",
                       StringComparison.OrdinalIgnoreCase) ||
                   value.StartsWith(
                       "IPM.Appointment",
                       StringComparison.OrdinalIgnoreCase);
        }

        internal static bool IsReadableItem(object item)
        {
            if (item == null)
            {
                return false;
            }

            dynamic mail = item;
            return IsReadableItemClass(
                SafeString(() => mail.MessageClass));
        }

        private static string BuildRecipients(
            dynamic mail,
            bool isAppointment)
        {
            if (isAppointment)
            {
                var required = SafeString(
                    () => mail.RequiredAttendees);
                var optional = SafeString(
                    () => mail.OptionalAttendees);
                var joined = required +
                    (required.Length > 0 && optional.Length > 0
                        ? "; "
                        : string.Empty) +
                    optional;
                if (joined.Trim().Length > 0)
                {
                    return joined;
                }
            }

            var to = SafeString(() => mail.To);
            if (to.Length > 0)
            {
                return to;
            }

            // Meeting invites expose attendees only through the
            // Recipients collection.
            object recipients = null;
            try
            {
                recipients = SafeObject(() => mail.Recipients);
                if (recipients == null)
                {
                    return string.Empty;
                }

                dynamic collection = recipients;
                var count = Math.Min(
                    Convert.ToInt32(collection.Count),
                    20);
                var names = new List<string>(count);
                for (var index = 1; index <= count; index++)
                {
                    object recipient = null;
                    try
                    {
                        recipient = collection.Item(index);
                        dynamic outlookRecipient = recipient;
                        var name = SafeString(
                            () => outlookRecipient.Name);
                        if (name.Length > 0)
                        {
                            names.Add(name);
                        }
                    }
                    finally
                    {
                        Release(recipient);
                    }
                }

                return string.Join("; ", names);
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                Release(recipients);
            }
        }

        private static string BuildMeetingDetails(
            dynamic mail,
            bool viaAssociatedAppointment)
        {
            object appointment = null;
            try
            {
                dynamic source = mail;
                if (viaAssociatedAppointment)
                {
                    // false: read the existing appointment only, never
                    // add the meeting to the calendar.
                    appointment = SafeObject(
                        () => mail.GetAssociatedAppointment(false));
                    if (appointment == null)
                    {
                        return string.Empty;
                    }

                    source = appointment;
                }

                dynamic details = source;
                var start = SafeDateTime(() => details.Start);
                var end = SafeDateTime(() => details.End);
                var location = SafeString(() => details.Location);
                if (start == null &&
                    end == null &&
                    location.Length == 0)
                {
                    return string.Empty;
                }

                return "[Calendar item" +
                    (start != null
                        ? " | Start: " + start.Value.ToString(
                            "yyyy-MM-dd HH:mm")
                        : string.Empty) +
                    (end != null
                        ? " | End: " + end.Value.ToString(
                            "yyyy-MM-dd HH:mm")
                        : string.Empty) +
                    (location.Length > 0
                        ? " | Location: " + location
                        : string.Empty) +
                    "]\n";
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                Release(appointment);
            }
        }

        internal static int CountRemoteImages(string htmlBody)
        {
            if (string.IsNullOrEmpty(htmlBody))
            {
                return 0;
            }

            // Web-hosted images exist only as URLs; the message stores no
            // image bytes, so Scribble can never read them. Embedded images
            // use cid: sources and arrive through the Attachments path.
            var bounded = htmlBody.Length > 512 * 1024
                ? htmlBody.Substring(0, 512 * 1024)
                : htmlBody;
            try
            {
                return System.Text.RegularExpressions.Regex.Matches(
                    bounded,
                    "<img\\b[^>]*?src\\s*=\\s*[\"']?https?://",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase,
                    TimeSpan.FromSeconds(2)).Count;
            }
            catch (System.Text.RegularExpressions.RegexMatchTimeoutException)
            {
                return 0;
            }
        }

        private static string[] CaptureAttachmentNames(dynamic mail)
        {
            object attachments = null;
            try
            {
                attachments = mail.Attachments;
                if (attachments == null)
                {
                    return new string[0];
                }

                dynamic outlookAttachments = attachments;
                var count = Math.Min(
                    Convert.ToInt32(outlookAttachments.Count),
                    EmailAttachmentReader.MaxAttachments);
                var names = new List<string>(count);
                for (var index = 1; index <= count; index++)
                {
                    object attachment = null;
                    try
                    {
                        attachment = outlookAttachments.Item(index);
                        dynamic outlookAttachment = attachment;
                        var fileName = SafeString(
                            () => outlookAttachment.FileName);
                        if (fileName.Length == 0)
                        {
                            continue;
                        }

                        long attachmentSize;
                        try
                        {
                            attachmentSize = Convert.ToInt64(
                                outlookAttachment.Size);
                        }
                        catch
                        {
                            attachmentSize = 0;
                        }

                        if (EmailAttachmentReader
                            .IsLikelySignatureImage(
                                attachment,
                                System.IO.Path.GetExtension(
                                    fileName),
                                attachmentSize))
                        {
                            continue;
                        }

                        // Every other attachment is listed;
                        // read_messages notes the ones that cannot
                        // be converted.
                        names.Add(
                            TextBoundary.SingleLine(fileName, 180));
                    }
                    finally
                    {
                        Release(attachment);
                    }
                }

                return names
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
            }
            catch
            {
                return new string[0];
            }
            finally
            {
                Release(attachments);
            }
        }

        // Strict email check, kept separate from IsReadableItem so the
        // writing-style sampler never learns from meeting responses.
        internal static bool IsMailItem(object item)
        {
            if (item == null)
            {
                return false;
            }

            dynamic mail = item;
            return SafeString(() => mail.MessageClass).StartsWith(
                "IPM.Note",
                StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildSender(dynamic mail)
        {
            var name = SafeString(() => mail.SenderName);
            var address = SafeString(() => mail.SenderEmailAddress);
            if (name.Length == 0)
            {
                return address;
            }

            if (address.Length == 0 ||
                name.Equals(address, StringComparison.OrdinalIgnoreCase))
            {
                return name;
            }

            return name + " <" + address + ">";
        }

        private static object SafeObject(Func<object> reader)
        {
            try
            {
                return reader();
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

        private static DateTime? SafeDateTime(Func<object> reader)
        {
            try
            {
                var value = reader();
                if (value is DateTime)
                {
                    return (DateTime)value;
                }
            }
            catch
            {
            }

            return null;
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
