using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Threading.Tasks;

namespace Wincent
{
    public enum QuickAccess
    {
        All,
        RecentFiles,
        FrequentFolders
    }

    public static class QuickAccessQuery
    {
        public static async Task<List<string>> GetRecentFilesAsync(int maxRetries = 2)
        {
            return await ExecuteWithValidationAsync(QuickAccessType.RecentFiles, maxRetries);
        }

        public static async Task<List<string>> GetFrequentFoldersAsync(int maxRetries = 2)
        {
            return await ExecuteWithValidationAsync(QuickAccessType.FrequentFolders, maxRetries);
        }

        public static async Task<List<string>> GetAllItemsAsync(int maxRetries = 2)
        {
            return await ExecuteWithValidationAsync(QuickAccessType.All, maxRetries);
        }

        private static async Task<List<string>> ExecuteWithValidationAsync(QuickAccessType queryType, int maxRetries)
        {
            ValidateExecutionPrerequisites(queryType);

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    return await QueryItemsAsync(queryType);
                }
                catch (ScriptExecutionException ex) when (attempt < maxRetries)
                {
                    HandleTransientError(ex, queryType, attempt);
                }
            }
            throw new InvalidOperationException($"Query failed, retried {maxRetries} times.");
        }

        private static void ValidateExecutionPrerequisites(QuickAccessType queryType)
        {
            if (!ExecutionFeasible.CheckScriptFeasible())
                throw new SecurityException("PowerShell execution policy restrictions.");
        }

        private static async Task<List<string>> QueryItemsAsync(QuickAccessType queryType)
        {
            var scriptType = MapToScriptType(queryType);
            var result = await ScriptExecutor.ExecutePSScript(scriptType, null);

            if (result.ExitCode != 0)
                throw new ScriptExecutionException("wrong exit code", result.Error, result.Output);

            return ParsePowerShellOutput(result.Output);
        }

        private static PSScript MapToScriptType(QuickAccessType queryType)
        {
            switch (queryType)
            {
                case QuickAccessType.RecentFiles:
                    return PSScript.QueryRecentFile;
                case QuickAccessType.FrequentFolders:
                    return PSScript.QueryFrequentFolder;
                case QuickAccessType.All:
                    return PSScript.QueryQuickAccess;
                default:
                    throw new ArgumentOutOfRangeException(nameof(queryType));
            }
        }

        private static List<string> ParsePowerShellOutput(string output)
        {
            return output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                         .Select(s => s.Trim())
                         .Where(s => !string.IsNullOrEmpty(s))
                         .ToList();
        }

        private static void HandleTransientError(Exception ex, QuickAccessType queryType, int attempt)
        {
            System.Diagnostics.Debug.WriteLine($"Attempt {attempt} failed: {ex.Message}.");

            if (ExecutionFeasible.IsAdministrator())
            {
                try { ExecutionFeasible.FixExecutionPolicy(); }
                catch { /* Ignore failed fixing */ }
            }
        }

        private enum QuickAccessType
        {
            RecentFiles,
            FrequentFolders,
            All
        }
    }
}
