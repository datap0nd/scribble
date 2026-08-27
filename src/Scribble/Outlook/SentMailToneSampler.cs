using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Scribble.Chat;
using Scribble.Security;

namespace Scribble.Outlook
{
    public sealed class SentMailToneSampler
    {
        private const int SentItemsFolder = 5;
        private readonly object _outlookApplication;

        public SentMailToneSampler(object outlookApplication)
        {
            _outlookApplication = outlookApplication ??
                throw new ArgumentNullException(nameof(outlookApplication));
        }

        public IReadOnlyList<MessageSnapshot> CaptureRecent()
        {
            var samples = new List<MessageSnapshot>();
            object session = null;
            object folder = null;
            object items = null;
            try
            {
                dynamic application = _outlookApplication;
                session = application.Session;
                dynamic outlookSession = session;
                folder = outlookSession.GetDefaultFolder(
                    SentItemsFolder);
                dynamic sentFolder = folder;
                items = sentFolder.Items;
                dynamic sentItems = items;
                sentItems.Sort("[SentOn]", true);
                var count = Math.Min(
                    Convert.ToInt32(sentItems.Count),
                    120);
                for (var index = 1;
                     index <= count &&
                     samples.Count < ToneProfileRequestFactory.MaxSamples;
                     index++)
                {
                    object item = null;
                    try
                    {
                        item = sentItems.Item(index);
                        if (!MessageReader.IsMailItem(item))
                        {
                            continue;
                        }

                        var snapshot = MessageReader.CaptureItem(item);
                        var cleanedBody = CleanBody(snapshot.Body);
                        if (cleanedBody.Length < 40)
                        {
                            continue;
                        }

                        samples.Add(
                            new MessageSnapshot(
                                snapshot.EntryId,
                                snapshot.StoreId,
                                snapshot.Subject,
                                string.Empty,
                                string.Empty,
                                snapshot.ReceivedAt,
                                cleanedBody));
                    }
                    finally
                    {
                        Release(item);
                    }
                }
            }
            finally
            {
                Release(items);
                Release(folder);
                Release(session);
            }

            return samples;
        }

        public static string CleanBody(string body)
        {
            var text = TextBoundary.PlainText(
                body,
                TextBoundary.MaxMessageBodyCharacters);
            var markers = new[]
            {
                "-----Original Message-----",
                "\nFrom:",
                "\r\nFrom:",
                "\nSent:",
                "\r\nSent:"
            };
            var cut = text.Length;
            foreach (var marker in markers)
            {
                var index = text.IndexOf(
                    marker,
                    StringComparison.OrdinalIgnoreCase);
                if (index >= 0)
                {
                    cut = Math.Min(cut, index);
                }
            }

            return TextBoundary.PlainText(
                text.Substring(0, cut),
                ToneProfileRequestFactory.MaxSampleCharacters);
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
