using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Scribble.Office
{
    // Browser continuations run on pool threads. Office automation requires a
    // pumped STA, including continuations after source/vision model review.
    internal static class OfficeThread
    {
        internal static Task<T> RunAsync<T>(Func<Task<T>> action, CancellationToken token)
        {
            var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            var thread = new Thread(() =>
            {
                try
                {
                    using (var control = new Control())
                    using (var context = new ApplicationContext())
                    {
                        var handle = control.Handle;
                        SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                        control.BeginInvoke(new Action(async () =>
                        {
                            // A new STA has no COM message filter, so an Office
                            // server that is momentarily busy (PowerPoint starting
                            // a chart's data grid, Excel recalculating) rejects the
                            // call outright. Retry only calls Office rejected
                            // before executing them; the filter is restored on
                            // this same thread before it exits.
                            Scribble.Testing.TestLabComMessageFilter retry = null;
                            try { retry = new Scribble.Testing.TestLabComMessageFilter(token); }
                            catch (Exception) { retry = null; }
                            try { token.ThrowIfCancellationRequested(); completion.TrySetResult(await action()); }
                            catch (OperationCanceledException) { completion.TrySetCanceled(); }
                            catch (Exception exception) { completion.TrySetException(exception); }
                            finally
                            {
                                try { retry?.Dispose(); } catch (Exception) { }
                                context.ExitThread();
                            }
                        }));
                        Application.Run(context);
                    }
                }
                catch (Exception exception) { completion.TrySetException(exception); }
            });
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            return completion.Task;
        }
    }
}
