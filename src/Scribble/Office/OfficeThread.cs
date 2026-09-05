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
                            try { token.ThrowIfCancellationRequested(); completion.TrySetResult(await action()); }
                            catch (OperationCanceledException) { completion.TrySetCanceled(); }
                            catch (Exception exception) { completion.TrySetException(exception); }
                            finally { context.ExitThread(); }
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
