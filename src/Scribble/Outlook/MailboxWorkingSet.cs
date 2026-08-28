using System;
using System.Collections.Generic;
using System.Globalization;
using Scribble.Security;

namespace Scribble.Outlook
{
    public static class MailboxWorkingSet
    {
        // Fixed end-user default of ten. The legacy override wrapper
        // is reset to this value whenever settings are applied.
        public const int RecommendedMaxMessages = 10;

        public static int MaxMessages
        {
            get { return LimitOverrides.WorkingSetMessages; }
        }

        public static string HandleAt(int index)
        {
            if (index < 0 || index >= MaxMessages)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            return "context" +
                (index + 1).ToString(CultureInfo.InvariantCulture);
        }

        public static IReadOnlyList<MessageSnapshot> Normalize(
            IEnumerable<MessageSnapshot> messages)
        {
            var result = new List<MessageSnapshot>();
            if (messages == null)
            {
                return result;
            }

            var identities = new HashSet<string>(
                StringComparer.Ordinal);
            foreach (var message in messages)
            {
                if (message == null || message.EntryId.Length == 0)
                {
                    continue;
                }

                var identity = message.EntryId +
                    "\n" +
                    message.StoreId;
                if (!identities.Add(identity))
                {
                    continue;
                }

                result.Add(message);
                if (result.Count == MaxMessages)
                {
                    break;
                }
            }

            return result;
        }
    }
}
