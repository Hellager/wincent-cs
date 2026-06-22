using System;
using System.Runtime.ExceptionServices;
using System.Threading;

namespace Wincent
{
    /// <summary>
    /// Runs Shell COM work on a dedicated single-threaded apartment worker.
    /// </summary>
    /// <remarks>
    /// Windows Shell verbs and namespace operations can hang or fail when invoked from an arbitrary caller thread,
    /// especially on Windows 11 where Explorer often performs cross-process COM calls for Quick Access actions such as
    /// <c>pintohome</c>. A fresh STA worker gives COM a predictable apartment for those calls.
    /// </remarks>
    internal static class StaThreadRunner
    {
        // Normal Quick Access usage has at most one or two concurrent Shell operations. This limit is a guardrail
        // against repeated timeouts accumulating background STA workers; it is not intended as throughput tuning.
        internal const int MaxActiveStaWorkers = 4;
        private static int _activeStaWorkers;

        // Exposed internally for tests and diagnostics. The count tracks workers that are still alive, not callers
        // still waiting for them.
        internal static int ActiveWorkerCount => Volatile.Read(ref _activeStaWorkers);

        public static void Run(
            Action action,
            TimeSpan timeout,
            INativeMethods nativeMethods = null,
            bool disableOle1Dde = false)
        {
            if (action == null)
                throw new ArgumentNullException(nameof(action));

            Run(
                () =>
                {
                    action();
                    return true;
                },
                timeout,
                nativeMethods,
                disableOle1Dde);
        }

        public static T Run<T>(
            Func<T> action,
            TimeSpan timeout,
            INativeMethods nativeMethods = null,
            bool disableOle1Dde = false)
        {
            if (action == null)
                throw new ArgumentNullException(nameof(action));
            if (timeout <= TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(timeout), "Timeout must be positive.");

            nativeMethods = nativeMethods ?? new DefaultNativeMethods();
            ExceptionDispatchInfo exception = null;
            T result = default(T);

            if (!TryAcquireActiveWorker())
            {
                throw new TimeoutException(
                    "Too many Shell COM operations are still running; previous timed-out operations may not have finished yet.");
            }

            var thread = new Thread(() =>
            {
                try
                {
                    // COM is initialized and uninitialized on the worker itself. If the caller times out below, this
                    // block may still complete later and mutate Explorer state.
                    using (ComGuard.InitializeSta(nativeMethods, disableOle1Dde))
                    {
                        result = action();
                    }
                }
                catch (Exception ex)
                {
                    exception = ExceptionDispatchInfo.Capture(ex);
                }
                finally
                {
                    // Release the worker slot only when the STA thread exits naturally. A caller-side timeout does not
                    // cancel Shell COM, so releasing here prevents a burst of timeouts from spawning unlimited workers.
                    Interlocked.Decrement(ref _activeStaWorkers);
                }
            });

            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            try
            {
                thread.Start();
            }
            catch
            {
                Interlocked.Decrement(ref _activeStaWorkers);
                throw;
            }

            if (!thread.Join(timeout))
                throw new TimeoutException($"COM STA thread timed out after {timeout.TotalSeconds:0.###} seconds.");

            if (exception != null)
                exception.Throw();

            return result;
        }

        private static bool TryAcquireActiveWorker()
        {
            while (true)
            {
                int current = Volatile.Read(ref _activeStaWorkers);
                if (current >= MaxActiveStaWorkers)
                    return false;

                if (Interlocked.CompareExchange(ref _activeStaWorkers, current + 1, current) == current)
                    return true;
            }
        }
    }
}
