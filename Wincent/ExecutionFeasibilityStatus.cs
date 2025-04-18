using System;
using System.Threading.Tasks;

namespace Wincent
{
    /// <summary>
    /// Script execution feasibility status
    /// </summary>
    public class ExecutionFeasibilityStatus
    {
        /// <summary>
        /// Script executor interface (internal use)
        /// </summary>
        public interface IScriptExecutor
        {
            /// <summary>
            /// Executes PowerShell script
            /// </summary>
            /// <param name="script">Script type</param>
            /// <param name="parameter">Script parameter</param>
            /// <param name="timeoutSeconds">Timeout duration in seconds</param>
            /// <returns>Script execution result</returns>
            Task<ScriptResult> ExecutePSScriptWithTimeout(PSScript script, string parameter, int timeoutSeconds);
        }

        /// <summary>
        /// Indicates if system information can be queried
        /// </summary>
        public bool Query { get; private set; }

        /// <summary>
        /// Indicates if system settings can be modified
        /// </summary>
        public bool Handle { get; private set; }

        /// <summary>
        /// Creates new instance
        /// </summary>
        public ExecutionFeasibilityStatus()
        {
        }

        /// <summary>
        /// Creates new instance with specified feasibility
        /// </summary>
        /// <param name="query">Query capability</param>
        /// <param name="handle">Modification capability</param>
        public ExecutionFeasibilityStatus(bool query, bool handle)
        {
            Query = query;
            Handle = handle;
        }

        /// <summary>
        /// Checks PowerShell script execution environment feasibility
        /// </summary>
        /// <param name="executor">Script executor</param>
        /// <param name="timeoutSeconds">Timeout in seconds (0 uses default 10 seconds)</param>
        /// <returns>Execution feasibility status</returns>
        /// <exception cref="ArgumentNullException">Thrown when executor is null</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown for negative timeout</exception>
        public static async Task<ExecutionFeasibilityStatus> CheckAsync(IScriptExecutor executor, int timeoutSeconds = 10)
        {
            if (executor == null)
                throw new ArgumentNullException(nameof(executor), "Script executor cannot be null");

            if (timeoutSeconds < 0)
                throw new ArgumentOutOfRangeException(nameof(timeoutSeconds), "Timeout cannot be negative");

            // Use default 10 seconds if timeout is 0
            int actualTimeout = timeoutSeconds == 0 ? 10 : timeoutSeconds;

            // Parallel execution for performance
            var queryTask = CheckFeasibilityAsync(executor, PSScript.CheckQueryFeasible, actualTimeout);
            var handleTask = CheckFeasibilityAsync(executor, PSScript.CheckPinUnpinFeasible, actualTimeout);

            await Task.WhenAll(queryTask, handleTask);

            return new ExecutionFeasibilityStatus(
                await queryTask,
                await handleTask
            );
        }

        /// <summary>
        /// Performs feasibility check using Wincent.IScriptExecutor
        /// </summary>
        /// <param name="executor">Script executor</param>
        /// <param name="timeoutSeconds">Timeout in seconds (0 uses default 10 seconds)</param>
        /// <returns>Execution feasibility status</returns>
        public static async Task<ExecutionFeasibilityStatus> CheckAsync(Wincent.IScriptExecutor executor, int timeoutSeconds = 10)
        {
            if (executor == null)
                throw new ArgumentNullException(nameof(executor), "Script executor cannot be null");

            // Convert Wincent.IScriptExecutor via adapter
            var adapter = new ScriptExecutorAdapter(executor);
            return await CheckAsync(adapter, timeoutSeconds);
        }

        /// <summary>
        /// Checks PowerShell environment feasibility using default executor
        /// </summary>
        /// <param name="timeoutSeconds">Timeout in seconds (0 uses default 10 seconds)</param>
        /// <returns>Execution feasibility status</returns>
        public static async Task<ExecutionFeasibilityStatus> CheckAsync(int timeoutSeconds = 10)
        {
            // Using default ScriptExecutor
            using (var executor = new ScriptExecutor())
            {
                return await CheckAsync((Wincent.IScriptExecutor)executor, timeoutSeconds);
            }
        }

        private static async Task<bool> CheckFeasibilityAsync(
            IScriptExecutor executor,
            PSScript script,
            int timeoutSeconds)
        {
            try
            {
                var result = await executor.ExecutePSScriptWithTimeout(script, null, timeoutSeconds);
                return result.ExitCode == 0;
            }
            catch (Exception)
            {
                // Any exception indicates failure
                return false;
            }
        }

        /// <summary>
        /// Adapter for converting Wincent.IScriptExecutor to ExecutionFeasibilityStatus.IScriptExecutor
        /// </summary>
        private class ScriptExecutorAdapter : IScriptExecutor
        {
            private readonly Wincent.IScriptExecutor _executor;

            public ScriptExecutorAdapter(Wincent.IScriptExecutor executor)
            {
                _executor = executor ?? throw new ArgumentNullException(nameof(executor));
            }

            public Task<ScriptResult> ExecutePSScriptWithTimeout(PSScript script, string parameter, int timeoutSeconds)
            {
                return _executor.ExecutePSScriptWithTimeout(script, parameter, timeoutSeconds);
            }
        }
    }
}
