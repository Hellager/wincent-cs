using System;
using System.Runtime.ExceptionServices;
using System.Threading;

namespace Wincent
{
    internal static class StaThreadRunner
    {
        internal const int MaxActiveStaWorkers = 4;
        private static int _activeStaWorkers;

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
