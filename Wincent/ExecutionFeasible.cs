using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Security;
using System.Security.Principal;
using System.Threading.Tasks;

namespace Wincent
{
    /// <summary>
    /// Represents PowerShell execution policy status
    /// </summary>
    public class ExecutionFeasibilityStatus
    {
        /// <summary>
        /// Indicates if system information can be queried
        /// </summary>
        public bool Query { get; }

        /// <summary>
        /// Indicates if system settings can be modified
        /// </summary>
        public bool Handle { get; }

        /// <summary>
        /// Creates execution feasibility status
        /// </summary>
        /// <param name="query">Can query system information</param>
        /// <param name="handle">Can modify system settings</param>
        public ExecutionFeasibilityStatus(bool query, bool handle)
        {
            Query = query;
            Handle = handle;
        }

        /// <summary>
        /// Script executor interface for dependency injection
        /// </summary>
        public interface IScriptExecutor
        {
            Task<ScriptResult> ExecutePSScriptWithTimeout(PSScript script, string parameter, int timeoutSeconds);
        }

        /// <summary>
        /// Checks PowerShell script execution environment feasibility
        /// </summary>
        /// <param name="executor">Script executor instance</param>
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
        /// Checks PowerShell environment feasibility using default executor
        /// </summary>
        /// <param name="timeoutSeconds">Timeout in seconds (0 uses default 10 seconds)</param>
        /// <returns>Execution feasibility status</returns>
        public static async Task<ExecutionFeasibilityStatus> CheckAsync(int timeoutSeconds = 10)
        {
            using (var executor = new ScriptExecutor())
            {
                return await CheckAsync(executor, timeoutSeconds);
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
    }

    /// <summary>
    /// Static class providing PowerShell execution policy checking and remediation
    /// </summary>
    public static class ExecutionFeasible
    {
        // Internal interface for dependency injection
        internal interface IRegistryService
        {
            string GetExecutionPolicy();
            void SetExecutionPolicy(string policy);
            bool IsAdministrator();
        }

        // Default implementation
        private class DefaultRegistryService : IRegistryService
        {
            private static readonly string ExecutionPolicyKeyPath = @"SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell";

            public string GetExecutionPolicy()
            {
                try
                {
                    using (var regKey = OpenExecutionPolicyKey(false))
                    {
                        return regKey?.GetValue("ExecutionPolicy", "NotSet") as string ?? "NotSet";
                    }
                }
                catch (SecurityException)
                {
                    return "AccessDenied";
                }
            }

            public void SetExecutionPolicy(string policy)
            {
                try
                {
                    using (var regKey = OpenExecutionPolicyKey(true))
                    {
                        regKey?.SetValue("ExecutionPolicy", policy, RegistryValueKind.String);
                    }
                }
                catch (SecurityException ex)
                {
                    throw new SecurityException("Administrator privileges required to modify execution policy.", ex);
                }
            }

            public bool IsAdministrator()
            {
                try
                {
                    using (var identity = WindowsIdentity.GetCurrent())
                    {
                        var principal = new WindowsPrincipal(identity);
                        return principal.IsInRole(WindowsBuiltInRole.Administrator);
                    }
                }
                catch
                {
                    return false;
                }
            }

            private RegistryKey OpenExecutionPolicyKey(bool writable)
            {
                try
                {
                    return Registry.CurrentUser.OpenSubKey(ExecutionPolicyKeyPath, writable)
                        ?? Registry.CurrentUser.CreateSubKey(ExecutionPolicyKeyPath);
                }
                catch (SecurityException)
                {
                    return null;
                }
            }
        }

        // Service instances
        private static IRegistryService _registryService = new DefaultRegistryService();
        private static readonly string[] FeasiblePolicies = { "AllSigned", "Bypass", "RemoteSigned", "Unrestricted" };

        // Test service replacement
        internal static void SetRegistryService(IRegistryService service)
        {
            _registryService = service ?? new DefaultRegistryService();
        }

        // Reset service for testing
        internal static void ResetRegistryService()
        {
            _registryService = new DefaultRegistryService();
        }

        /// <summary>
        /// Asynchronously checks PowerShell execution environment feasibility
        /// </summary>
        /// <param name="timeoutSeconds">Timeout in seconds</param>
        /// <returns>Execution feasibility status</returns>
        public static async Task<ExecutionFeasibilityStatus> CheckFeasibilityAsync(int timeoutSeconds = 10)
        {
            return await ExecutionFeasibilityStatus.CheckAsync(timeoutSeconds);
        }

        /// <summary>
        /// Checks if current execution policy allows script execution
        /// </summary>
        /// <returns>True if execution policy permits scripts</returns>
        public static bool CheckScriptFeasible()
        {
            try
            {
                var value = _registryService.GetExecutionPolicy();
                return IsValidExecutionPolicy(value);
            }
            catch (SecurityException)
            {
                return false;
            }
        }

        /// <summary>
        /// Fixes execution policy by setting to RemoteSigned
        /// </summary>
        /// <exception cref="SecurityException">Thrown with insufficient privileges</exception>
        public static void FixExecutionPolicy()
        {
            try
            {
                _registryService.SetExecutionPolicy("RemoteSigned");
            }
            catch (SecurityException ex)
            {
                throw new SecurityException("Administrator privileges required to modify policy.", ex);
            }
        }

        /// <summary>
        /// Retrieves current execution policy
        /// </summary>
        /// <returns>Current policy as string</returns>
        public static string GetExecutionPolicy()
        {
            return _registryService.GetExecutionPolicy();
        }

        /// <summary>
        /// Checks if current user has administrator privileges
        /// </summary>
        /// <returns>True for administrator account</returns>
        public static bool IsAdministrator()
        {
            return _registryService.IsAdministrator();
        }

        private static bool IsValidExecutionPolicy(string policy)
        {
            if (string.IsNullOrEmpty(policy))
                return false;

            foreach (var validPolicy in FeasiblePolicies)
            {
                if (string.Equals(policy, validPolicy, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }
    }

    /// <summary>
    /// Static class for PowerShell environment validation
    /// </summary>
    public static class PowerShellFeasible
    {
        // Internal interface for dependency injection
        internal interface IPowerShellService
        {
            string ExecuteScript(string script);
            Task<string> ExecuteScriptAsync(string script);
        }

        // Default implementation
        private class DefaultPowerShellService : IPowerShellService
        {
            public string ExecuteScript(string script)
            {
                using (var process = new Process())
                {
                    process.StartInfo = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-Command \"{script}\"",
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };

                    process.Start();
                    return process.StandardOutput.ReadToEnd().Trim();
                }
            }

            public async Task<string> ExecuteScriptAsync(string script)
            {
                using (var process = new Process())
                {
                    process.StartInfo = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-Command \"{script}\"",
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };

                    var tcs = new TaskCompletionSource<string>();

                    process.EnableRaisingEvents = true;
                    process.Exited += (sender, args) =>
                    {
                        var output = process.StandardOutput.ReadToEnd().Trim();
                        tcs.SetResult(output);
                    };

                    process.Start();
                    return await tcs.Task;
                }
            }
        }

        // Service instance
        private static IPowerShellService _powerShellService = new DefaultPowerShellService();

        // Test service replacement
        internal static void SetPowerShellService(IPowerShellService service)
        {
            _powerShellService = service ?? new DefaultPowerShellService();
        }

        // Reset service for testing
        internal static void ResetPowerShellService()
        {
            _powerShellService = new DefaultPowerShellService();
        }

        /// <summary>
        /// Validates PowerShell script execution capability
        /// </summary>
        /// <returns>True if PowerShell environment is functional</returns>
        public static bool CheckScriptExecution()
        {
            try
            {
                var script = "$PSVersionTable.PSVersion.Major";
                var output = _powerShellService.ExecuteScript(script);
                return !string.IsNullOrEmpty(output) && int.TryParse(output, out var version) && version >= 3;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Asynchronously validates PowerShell script execution capability
        /// </summary>
        /// <returns>True if PowerShell environment is functional</returns>
        public static async Task<bool> CheckScriptExecutionAsync()
        {
            try
            {
                var script = "$PSVersionTable.PSVersion.Major";
                var output = await _powerShellService.ExecuteScriptAsync(script);
                return !string.IsNullOrEmpty(output) && int.TryParse(output, out var version) && version >= 3;
            }
            catch
            {
                return false;
            }
        }
    }
}
