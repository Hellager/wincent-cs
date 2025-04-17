//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Security;
//using System.Threading.Tasks;

//namespace Wincent
//{
//    public enum QuickAccess
//    {
//        All,
//        RecentFiles,
//        FrequentFolders
//    }

//    /// <summary>
//    /// Static class providing quick access item query functionality
//    /// </summary>
//    public static class QuickAccessQuery
//    {
//        // Internal interface for dependency injection
//        internal interface IScriptExecutorService
//        {
//            Task<ScriptResult> ExecutePSScript(PSScript scriptType, string parameter);
//        }

//        internal interface IExecutionFeasibleService
//        {
//            bool CheckScriptFeasible();
//            bool IsAdministrator();
//            void FixExecutionPolicy();
//        }

//        // Default implementation
//        private class DefaultScriptExecutorService : IScriptExecutorService
//        {
//            public Task<ScriptResult> ExecutePSScript(PSScript scriptType, string parameter)
//            {
//                return ScriptExecutor.ExecutePSScript(scriptType, parameter);
//            }
//        }

//        private class DefaultExecutionFeasibleService : IExecutionFeasibleService
//        {
//            public bool CheckScriptFeasible()
//            {
//                return ExecutionFeasible.CheckScriptFeasible();
//            }

//            public bool IsAdministrator()
//            {
//                return ExecutionFeasible.IsAdministrator();
//            }

//            public void FixExecutionPolicy()
//            {
//                ExecutionFeasible.FixExecutionPolicy();
//            }
//        }

//        // Service instances
//        private static IScriptExecutorService _scriptExecutorService = new DefaultScriptExecutorService();
//        private static IExecutionFeasibleService _executionFeasibleService = new DefaultExecutionFeasibleService();

//        // Service replacement method for testing
//        internal static void SetServices(IScriptExecutorService scriptExecutor, IExecutionFeasibleService executionFeasible)
//        {
//            _scriptExecutorService = scriptExecutor ?? new DefaultScriptExecutorService();
//            _executionFeasibleService = executionFeasible ?? new DefaultExecutionFeasibleService();
//        }

//        // Service reset method for testing
//        internal static void ResetServices()
//        {
//            _scriptExecutorService = new DefaultScriptExecutorService();
//            _executionFeasibleService = new DefaultExecutionFeasibleService();
//        }

//        /// <summary>
//        /// Retrieves the list of recent files
//        /// </summary>
//        /// <param name="maxRetries">Maximum number of retries</param>
//        /// <returns>List of recent file paths</returns>
//        public static async Task<List<string>> GetRecentFilesAsync(int maxRetries = 2)
//        {
//            return await ExecuteWithValidationAsync(QuickAccessType.RecentFiles, maxRetries);
//        }

//        /// <summary>
//        /// Retrieves the list of frequent folders
//        /// </summary>
//        /// <param name="maxRetries">Maximum number of retries</param>
//        /// <returns>List of frequent folder paths</returns>
//        public static async Task<List<string>> GetFrequentFoldersAsync(int maxRetries = 2)
//        {
//            return await ExecuteWithValidationAsync(QuickAccessType.FrequentFolders, maxRetries);
//        }

//        /// <summary>
//        /// Retrieves all quick access items
//        /// </summary>
//        /// <param name="maxRetries">Maximum number of retries</param>
//        /// <returns>List of all quick access item paths</returns>
//        public static async Task<List<string>> GetAllItemsAsync(int maxRetries = 2)
//        {
//            return await ExecuteWithValidationAsync(QuickAccessType.All, maxRetries);
//        }

//        private static async Task<List<string>> ExecuteWithValidationAsync(QuickAccessType queryType, int maxRetries)
//        {
//            ValidateExecutionPrerequisites(queryType);

//            for (int attempt = 1; attempt <= maxRetries; attempt++)
//            {
//                try
//                {
//                    return await QueryItemsAsync(queryType);
//                }
//                catch (ScriptExecutionException ex) when (attempt <= maxRetries)
//                {
//                    HandleTransientError(ex, queryType, attempt);
//                }
//            }
//            throw new InvalidOperationException($"Query failed, retried {maxRetries} times.");
//        }

//        private static void ValidateExecutionPrerequisites(QuickAccessType queryType)
//        {
//            if (!_executionFeasibleService.CheckScriptFeasible())
//                throw new SecurityException("PowerShell execution policy restrictions.");
//        }

//        private static async Task<List<string>> QueryItemsAsync(QuickAccessType queryType)
//        {
//            var scriptType = MapToScriptType(queryType);
//            var result = await _scriptExecutorService.ExecutePSScript(scriptType, null);

//            if (result.ExitCode != 0)
//                throw new ScriptExecutionException("wrong exit code", result.Error, result.Output);

//            return ParsePowerShellOutput(result.Output);
//        }

//        private static PSScript MapToScriptType(QuickAccessType queryType)
//        {
//            switch (queryType)
//            {
//                case QuickAccessType.RecentFiles:
//                    return PSScript.QueryRecentFile;
//                case QuickAccessType.FrequentFolders:
//                    return PSScript.QueryFrequentFolder;
//                case QuickAccessType.All:
//                    return PSScript.QueryQuickAccess;
//                default:
//                    throw new ArgumentOutOfRangeException(nameof(queryType));
//            }
//        }

//        private static List<string> ParsePowerShellOutput(string output)
//        {
//            return output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
//                         .Select(s => s.Trim())
//                         .Where(s => !string.IsNullOrEmpty(s))
//                         .ToList();
//        }

//        private static void HandleTransientError(Exception ex, QuickAccessType queryType, int attempt)
//        {
//            System.Diagnostics.Debug.WriteLine($"Attempt {attempt} failed: {ex.Message}.");

//            if (_executionFeasibleService.IsAdministrator())
//            {
//                try { _executionFeasibleService.FixExecutionPolicy(); }
//                catch { /* Ignore failed fixing */ }
//            }
//        }

//        private enum QuickAccessType
//        {
//            RecentFiles,
//            FrequentFolders,
//            All
//        }
//    }
//}
