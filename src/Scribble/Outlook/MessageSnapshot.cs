using System;
using System.Collections.Generic;

namespace Scribble.Outlook
{
    public sealed class MessageSnapshot
    {
        private static readonly string[] EmptyAttachmentNames =
            new string[0];

        public MessageSnapshot(
            string entryId,
            string storeId,
            string subject,
            string sender,
            string recipients,
            DateTime? receivedAt,
            string body,
            IReadOnlyList<string> attachmentNames = null,
            int remoteImageCount = 0)
        {
            EntryId = entryId ?? string.Empty;
            StoreId = storeId ?? string.Empty;
            Subject = subject ?? string.Empty;
            Sender = sender ?? string.Empty;
            Recipients = recipients ?? string.Empty;
            ReceivedAt = receivedAt;
            Body = body ?? string.Empty;
            AttachmentNames = attachmentNames ??
                EmptyAttachmentNames;
            RemoteImageCount = Math.Max(0, remoteImageCount);
        }

        public string EntryId { get; }

        public string StoreId { get; }

        public string Subject { get; }

        public string Sender { get; }

        public string Recipients { get; }

        public DateTime? ReceivedAt { get; }

        public string Body { get; }

        public IReadOnlyList<string> AttachmentNames { get; }

        public int RemoteImageCount { get; }

        public bool CanReply
        {
            get { return EntryId.Length > 0; }
        }
    }
}
