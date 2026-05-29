using System;
using System.Runtime.ExceptionServices;
using System.Threading;

namespace Wincent
{
    internal static class StaThreadRunner
    {
        public static void Run(Action action, TimeSpan timeout, INativeMethods nativeMethods = null)
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
                nativeMethods);
        }

        public static T Run<T>(Func<T> action, TimeSpan timeout, INativeMethods nativeMethods = null)
        {
            if (action == null)
                throw new ArgumentNullException(nameof(action));
            if (timeout <= TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(timeout), "Timeout must be positive.");

            nativeMethods = nativeMethods ?? new DefaultNativeMethods();
            ExceptionDispatchInfo exception = null;
            T result = default(T);

            var thread = new Thread(() =>
            {
                try
                {
                    using (ComGuard.InitializeSta(nativeMethods))
                    {
                        result = action();
                    }
                }
                catch (Exception ex)
                {
                    exception = ExceptionDispatchInfo.Capture(ex);
                }
            });

            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            if (!thread.Join(timeout))
                throw new TimeoutException($"COM STA thread timed out after {timeout.TotalSeconds:0.###} seconds.");

            if (exception != null)
                exception.Throw();

            return result;
        }
    }
}
