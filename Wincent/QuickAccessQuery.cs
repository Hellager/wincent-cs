using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Wincent
{
    public enum QuickAccess
    {
        FrequentFolders,
        RecentFiles,
        All

    }
    public static class QuickAccessQuery
    {
        private static async Task<List<string>> QueryWithPsScriptAsync(QuickAccess qaType)
        {
            try
            {
                var scriptType = qaType switch
                {
                    QuickAccess.All => PSScript.QueryQuickAccess,
                    QuickAccess.RecentFiles => PSScript.QueryRecentFile,
                    QuickAccess.FrequentFolders => PSScript.QueryFrequentFolder,
                    _ => throw new ArgumentOutOfRangeException(nameof(qaType), qaType, null)
                };

                var result = await ScriptExecutor.ExecutePSScript(scriptType, null);

                if (result.ExitCode != 0)
                    throw new ScriptExecutionException($"Script execution failed with exit code: {result.ExitCode}", result.Error, result.Output);

                return result.Output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                  .Select(s => s.Trim())
                                  .Where(s => !string.IsNullOrEmpty(s))
                                  .ToList();
            }
            catch (ScriptExecutionException ex)
            {
                throw new InvalidOperationException($"Script execution error: {ex.Error}", ex);
            }
        }

        public static async Task<List<string>> GetRecentFilesAsync()
        {
            if (!FeasibleChecker.CheckScriptFeasible())
                throw new InvalidOperationException("PowerShell script execution is not available");

            return await QueryWithPsScriptAsync(QuickAccess.RecentFiles);
        }

        public static async Task<List<string>> GetFrequentFoldersAsync()
        {
            if (!FeasibleChecker.CheckScriptFeasible())
                throw new InvalidOperationException("PowerShell script execution is not available");

            return await QueryWithPsScriptAsync(QuickAccess.FrequentFolders);
        }

        public static async Task<List<string>> GetAllItemsAsync()
        {
            if (!FeasibleChecker.CheckScriptFeasible())
                throw new InvalidOperationException("PowerShell script execution is not available");

            return await QueryWithPsScriptAsync(QuickAccess.All);
        }

        public static async Task<bool> ContainsInRecentFilesAsync(string keyword)
        {
            var items = await GetRecentFilesAsync();
            return items.Any(x => x.Contains(keyword, StringComparison.OrdinalIgnoreCase));
        }

        public static async Task<bool> ContainsInFrequentFoldersAsync(string keyword)
        {
            var items = await GetFrequentFoldersAsync();
            return items.Any(x => x.Contains(keyword, StringComparison.OrdinalIgnoreCase));
        }

        public static async Task<bool> ContainsInQuickAccessAsync(string keyword)
        {
            var items = await GetAllItemsAsync();
            return items.Any(x => x.Contains(keyword, StringComparison.OrdinalIgnoreCase));
        }
    }
}
