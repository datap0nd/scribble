using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;

namespace Scribble.Outlook
{
    public sealed class MailboxAttachmentPage
    {
        public string FileName { get; set; }
        public string Fingerprint { get; set; }
        public string Kind { get; set; }
        public string Text { get; set; }
        public string ImageDataUrl { get; set; }
        public int Offset { get; set; }
        public int? NextOffset { get; set; }
    }

    internal static class MailboxAttachmentPages
    {
        internal static int Count(object application, MessageSnapshot source)
        {
            object session = null, item = null, attachments = null;
            try
            {
                dynamic app = application;
                session = app.Session;
                dynamic ns = session;
                item = ns.GetItemFromID(source.EntryId, source.StoreId);
                dynamic mail = item;
                attachments = mail.Attachments;
                dynamic collection = attachments;
                return Convert.ToInt32(collection.Count);
            }
            finally { Release(attachments); Release(item); Release(session); }
        }

        internal static async Task<MailboxAttachmentPage> ReadAsync(object application, MessageSnapshot source,
            int index, int offset, CancellationToken token)
        {
            object session = null, item = null, attachments = null, attachment = null;
            string temporary = null;
            string name = null;
            try
            {
                // Only this short capture touches Outlook COM, on its owning context.
                dynamic app = application;
                session = app.Session;
                dynamic ns = session;
                item = ns.GetItemFromID(source.EntryId, source.StoreId);
                dynamic mail = item;
                attachments = mail.Attachments;
                dynamic collection = attachments;
                if (index < 1 || index > Convert.ToInt32(collection.Count)) throw new ArgumentException("Attachment index is outside the captured message.");
                attachment = collection.Item(index);
                dynamic file = attachment;
                name = Convert.ToString(file.FileName);
                var warning = AttachmentIntakePolicy.ValidateFile(Convert.ToInt64(file.Size));
                if (warning.Length > 0) throw new InvalidOperationException(warning);
                temporary = Path.Combine(Path.GetTempPath(), "scribble-page-" + Guid.NewGuid().ToString("N") + Path.GetExtension(name));
                file.SaveAsFile(temporary);
            }
            catch
            {
                if (temporary != null && File.Exists(temporary)) File.Delete(temporary);
                throw;
            }
            finally { Release(attachment); Release(attachments); Release(item); Release(session); }
            try
            {
                return await Task.Run(() =>
                {
                    token.ThrowIfCancellationRequested();
                    var page = EmailAttachmentReader.LoadLocalPage(temporary, offset, 6000, token);
                    page.FileName = name;
                    using (var stream = File.OpenRead(temporary))
                    using (var hash = SHA256.Create())
                        page.Fingerprint = BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "");
                    return page;
                }, token).ConfigureAwait(true);
            }
            finally { if (temporary != null && File.Exists(temporary)) File.Delete(temporary); }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }
}
