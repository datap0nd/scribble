using System.Collections.Generic;
using Scribble.Chat;

namespace Scribble.Utilities
{
    // In-process memory for each host's chat pane. Office task panes
    // can be torn down and recreated while the host application
    // keeps running; parking the conversation here lets a reopened
    // pane restore its transcript and history. Nothing is written to
    // disk - closing the Office application forgets the chat, which
    // keeps mailbox and document text off the file system.
    public static class PaneMemory
    {
        public sealed class Slot
        {
            public List<ChatTurn> History { get; } =
                new List<ChatTurn>();

            public List<string> Transcript { get; } =
                new List<string>();

            public string LastAnswer { get; set; } =
                string.Empty;
        }

        private static readonly object Gate = new object();
        private static readonly Dictionary<string, Slot> Slots =
            new Dictionary<string, Slot>();

        public static Slot For(string hostKind)
        {
            lock (Gate)
            {
                Slot slot;
                var key = hostKind ?? string.Empty;
                if (!Slots.TryGetValue(key, out slot))
                {
                    slot = new Slot();
                    Slots[key] = slot;
                }

                return slot;
            }
        }
    }
}
