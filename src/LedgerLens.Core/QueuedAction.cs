using System;
using System.Threading;
using System.Threading.Tasks;

namespace LedgerLens.Core
{
    public static class QueuedAction
    {
        /// <summary>Cancels work while queued. Once a synchronous mutation starts, its real outcome wins.</summary>
        public static Task<T> Run<T>(Action<Action> enqueue, Func<T> action, CancellationToken cancellation)
        {
            var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            // 0 = queued, 1 = executing/completed, 2 = canceled before execution.
            var state = 0;
            var registration = cancellation.Register(() =>
            {
                if (Interlocked.CompareExchange(ref state, 2, 0) == 0)
                    completion.TrySetCanceled();
            });
            try
            {
                enqueue(() =>
                {
                    try
                    {
                        if (Interlocked.CompareExchange(ref state, 1, 0) != 0)
                            return;
                        completion.TrySetResult(action());
                    }
                    catch (Exception error) { completion.TrySetException(error); }
                    finally { registration.Dispose(); }
                });
            }
            catch (Exception error)
            {
                registration.Dispose();
                completion.TrySetException(error);
            }
            return completion.Task;
        }
    }
}
