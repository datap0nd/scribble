using System;
using System.IO;
using System.Collections.Generic;
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
        public bool CacheHit { get; set; }
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
                item = Scribble.Testing.TestLabMail.OpenItem((object)ns, source.EntryId, source.StoreId);
                dynamic mail = item;
                attachments = mail.Attachments;
                dynamic collection = attachments;
                return Convert.ToInt32(collection.Count);
            }
            finally { Release(attachments); Release(item); Release(session); }
        }

        internal static async Task<MailboxAttachmentPage> ReadAsync(object application, MessageSnapshot source,
            int index, int offset, CancellationToken token, IDictionary<string, MailboxAttachmentPage> cache = null)
        {
            object session = null, item = null, attachments = null, attachment = null;
            string temporary = null;
            string name = null;
            var ownsTemporary = false;
            try
            {
                // The isolated test mailbox is a verified native projection of
                // manifest-backed files. Reading those immutable bytes avoids
                // Outlook providers that can block indefinitely while opening
                // PR_ATTACH_DATA_BIN. Production mail still uses Outlook COM.
                if (!Scribble.Testing.TestLabMailbox.TryResolveAttachment(
                    source.EntryId, source.StoreId, index, out temporary, out name))
                {
                    // Only this short capture touches Outlook COM, on its owning context.
                    dynamic app = application;
                    session = app.Session;
                    dynamic ns = session;
                    item = Scribble.Testing.TestLabMail.OpenItem((object)ns, source.EntryId, source.StoreId);
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
                    ownsTemporary = true;
                    if (!TrySaveByValue(file, temporary))
                        file.SaveAsFile(temporary);
                }
            }
            catch
            {
                if (ownsTemporary && temporary != null && File.Exists(temporary)) File.Delete(temporary);
                throw;
            }
            finally { Release(attachment); Release(attachments); Release(item); Release(session); }
            try
            {
                return await Task.Run(() =>
                {
                    token.ThrowIfCancellationRequested();
                    string fingerprint;
                    using (var stream = File.OpenRead(temporary))
                    using (var hash = SHA256.Create())
                        fingerprint = BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "");
                    // Hash the current bytes on every read. Reusing extracted pages
                    // must never hide a changed attachment, even at the same index.
                    var key = fingerprint + ":" + Path.GetExtension(name).ToLowerInvariant() + ":" + offset;
                    MailboxAttachmentPage page;
                    if (cache != null && cache.TryGetValue(key, out page))
                        return new MailboxAttachmentPage { FileName = name, Fingerprint = fingerprint,
                            Kind = page.Kind, Text = page.Text, ImageDataUrl = page.ImageDataUrl,
                            Offset = page.Offset, NextOffset = page.NextOffset, CacheHit = true };
                    page = EmailAttachmentReader.LoadLocalPage(temporary, offset, 6000, token);
                    page.FileName = name;
                    page.Fingerprint = fingerprint;
                    if (cache != null && string.IsNullOrEmpty(page.ImageDataUrl))
                    {
                        if (cache.Count >= 128) cache.Clear();
                        cache[key] = page;
                    }
                    return page;
                }, token).ConfigureAwait(true);
            }
            finally { if (ownsTemporary && temporary != null && File.Exists(temporary)) File.Delete(temporary); }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }

        private static bool TrySaveByValue(dynamic attachment, string path)
        {
            object accessor = null;
            try
            {
                accessor = attachment.PropertyAccessor;
                dynamic properties = accessor;
                var bytes = properties.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x37010102")
                    as byte[];
                if (bytes == null || bytes.Length == 0) return false;
                File.WriteAllBytes(path, bytes);
                return true;
            }
            catch
            {
                if (File.Exists(path)) File.Delete(path);
                return false;
            }
            finally
            {
                Release(accessor);
            }
        }
    }
}
