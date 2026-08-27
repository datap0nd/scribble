using System;
using System.Threading;

namespace Scribble.Security
{
    public sealed class OneShotDraftAuthorization
    {
        private int _available;
        private int _emailAvailable;
        private int _consumed;
        private int _created;
        private int _updated;

        public OneShotDraftAuthorization(bool authorized)
            : this(authorized, false)
        {
        }

        public OneShotDraftAuthorization(
            bool allowCreate,
            bool allowUpdate)
            : this(allowCreate, allowUpdate, SingleCall)
        {
        }

        // One user request authorizes ONE deliverable, which a
        // model may need several bounded calls to build: a dense
        // deck is written slide-batch by slide-batch because no
        // small local model can emit a whole executive deck in one
        // JSON payload. The budget caps those calls; it never
        // widens what a call may do, and email drafts stay strictly
        // single-shot (enforced by the draft hosts).
        public OneShotDraftAuthorization(
            bool allowCreate,
            bool allowUpdate,
            int callBudget)
        {
            CanCreate = allowCreate;
            CanUpdate = allowUpdate;
            WasAuthorized = allowCreate || allowUpdate;
            var bounded = callBudget < SingleCall
                ? SingleCall
                : (callBudget > MaxCallBudget
                    ? MaxCallBudget
                    : callBudget);
            CallBudget = WasAuthorized ? bounded : 0;
            _available = CallBudget;
            _emailAvailable = WasAuthorized ? SingleCall : 0;
        }

        public const int SingleCall = 1;

        public const int MaxCallBudget = 6;

        // How many draft calls this request was granted.
        public int CallBudget { get; }

        // How many remain unused.
        public int RemainingCalls
        {
            get
            {
                return Volatile.Read(ref _available);
            }
        }

        public bool WasAuthorized { get; }

        public bool CanCreate { get; }

        public bool CanUpdate { get; }

        public bool IsConsumed
        {
            get
            {
                return Volatile.Read(ref _consumed) == 1;
            }
        }

        public bool IsCreated
        {
            get
            {
                return Volatile.Read(ref _created) == 1;
            }
        }

        public bool IsUpdated
        {
            get
            {
                return Volatile.Read(ref _updated) == 1;
            }
        }

        public bool TryConsume()
        {
            while (true)
            {
                var current = Volatile.Read(ref _available);
                if (current <= 0)
                {
                    return false;
                }

                if (Interlocked.CompareExchange(
                    ref _available,
                    current - 1,
                    current) == current)
                {
                    Volatile.Write(ref _consumed, 1);
                    return true;
                }
            }
        }

        // Email drafts are never batched: whatever the call
        // budget is, one request may open at most one unsent email
        // draft, because recipients are the sensitive surface.
        public bool TryConsumeEmailDraft()
        {
            if (Interlocked.CompareExchange(
                ref _emailAvailable,
                0,
                1) != 1)
            {
                return false;
            }

            Volatile.Write(ref _consumed, 1);
            return true;
        }

        internal void MarkCreated()
        {
            if (!IsConsumed)
            {
                throw new InvalidOperationException(
                    "Draft permission must be consumed before success is recorded.");
            }

            Volatile.Write(ref _created, 1);
        }

        internal void MarkUpdated()
        {
            if (!IsConsumed)
            {
                throw new InvalidOperationException(
                    "Draft permission must be consumed before success is recorded.");
            }

            Volatile.Write(ref _updated, 1);
        }
    }
}
