using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace Scribble.Testing
{
    // Retry only calls which Office explicitly rejected before executing them.
    // Retrying a whole preparation step could duplicate a workbook or draft.
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public sealed class TestLabComMessageFilter : ITestLabOleMessageFilter, IDisposable
    {
        private ITestLabOleMessageFilter previous;
        private readonly CancellationToken cancel;
        private readonly int thread;
        private bool disposed;

        public TestLabComMessageFilter(CancellationToken cancel)
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Office retry handling requires the runner's STA thread.");
            this.cancel = cancel; thread = Thread.CurrentThread.ManagedThreadId;
            Marshal.ThrowExceptionForHR(CoRegisterMessageFilter(this, out previous));
        }
        public int HandleInComingCall(int callType, IntPtr caller, int elapsed, IntPtr info) { return 0; }
        public int RetryRejectedCall(IntPtr callee, int elapsed, int rejection)
        { return !cancel.IsCancellationRequested && elapsed >= 0 && elapsed < 30000 && (rejection == 1 || rejection == 2) ? 250 : -1; }
        public int MessagePending(IntPtr callee, int elapsed, int pendingType) { return 2; }
        public void Dispose()
        {
            if (disposed) return;
            if (thread != Thread.CurrentThread.ManagedThreadId) throw new InvalidOperationException("Restore the COM retry handler on its owning thread.");
            ITestLabOleMessageFilter removed;
            Marshal.ThrowExceptionForHR(CoRegisterMessageFilter(previous, out removed));
            disposed = true; previous = null;
            if (removed != null && Marshal.IsComObject(removed)) Marshal.ReleaseComObject(removed);
        }
        [DllImport("ole32.dll")]
        private static extern int CoRegisterMessageFilter(ITestLabOleMessageFilter current, out ITestLabOleMessageFilter previous);
    }

    [ComImport, ComVisible(true), Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ITestLabOleMessageFilter
    {
        [PreserveSig] int HandleInComingCall(int callType, IntPtr caller, int elapsed, IntPtr info);
        [PreserveSig] int RetryRejectedCall(IntPtr callee, int elapsed, int rejection);
        [PreserveSig] int MessagePending(IntPtr callee, int elapsed, int pendingType);
    }
}
