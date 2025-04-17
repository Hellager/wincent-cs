//using Microsoft.Win32;
//using System;
//using System.Diagnostics;
//using System.Security;
//using System.Security.Principal;
//using System.Threading.Tasks;

//namespace Wincent
//{
//    /// <summary>
//    /// Static class providing PowerShell execution policy checking and fixing functionality
//    /// </summary>
//    public static class ExecutionFeasible
//    {
//        // Internal interface for dependency injection
//        internal interface IRegistryService
//        {
//            string GetExecutionPolicy();
//            void SetExecutionPolicy(string policy);
//            bool IsAdministrator();
//        }

//        // Default implementation
//        private class DefaultRegistryService : IRegistryService
//        {
//            private static readonly string ExecutionPolicyKeyPath = @"SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell";

//            public string GetExecutionPolicy()
//            {
//                try
//                {
//                    using (var regKey = OpenExecutionPolicyKey(false))
//                    {
//                        return regKey?.GetValue("ExecutionPolicy", "NotSet") as string ?? "NotSet";
//                    }
//                }
//                catch (SecurityException)
//                {
//                    return "AccessDenied";
//                }
//            }

//            public void SetExecutionPolicy(string policy)
//            {
//                try
//                {
//                    using (var regKey = OpenExecutionPolicyKey(true))
//                    {
//                        regKey?.SetValue("ExecutionPolicy", policy, RegistryValueKind.String);
//                    }
//                }
//                catch (SecurityException ex)
//                {
//                    throw new SecurityException("Need administrator privileges to modify the execution policy.", ex);
//                }
//            }

//            public bool IsAdministrator()
//            {
//                try
//                {
//                    using (var identity = WindowsIdentity.GetCurrent())
//                    {
//                        var principal = new WindowsPrincipal(identity);
//                        return principal.IsInRole(WindowsBuiltInRole.Administrator);
//                    }
//                }
//                catch
//                {
//                    return false;
//                }
//            }

//            private RegistryKey OpenExecutionPolicyKey(bool writable)
//            {
//                try
//                {
//                    return Registry.CurrentUser.OpenSubKey(ExecutionPolicyKeyPath, writable)
//                        ?? Registry.CurrentUser.CreateSubKey(ExecutionPolicyKeyPath);
//                }
//                catch (SecurityException)
//                {
//                    return null;
//                }
//            }
//        }

//        // Service instances
//        private static IRegistryService _registryService = new DefaultRegistryService();
//        private static readonly string[] FeasiblePolicies = { "AllSigned", "Bypass", "RemoteSigned", "Unrestricted" };

//        // Service replacement method for testing
//        internal static void SetRegistryService(IRegistryService service)
//        {
//            _registryService = service ?? new DefaultRegistryService();
//        }

//        // Service reset method for testing
//        internal static void ResetRegistryService()
//        {
//            _registryService = new DefaultRegistryService();
//        }

//        /// <summary>
//        /// Checks if current PowerShell execution policy allows script execution
//        /// </summary>
//        /// <returns>True if execution policy permits script execution, otherwise false</returns>
//        public static bool CheckScriptFeasible()
//        {
//            try
//            {
//                var value = _registryService.GetExecutionPolicy();
//                return IsValidExecutionPolicy(value);
//            }
//            catch (SecurityException)
//            {
//                return false;
//            }
//        }

//        /// <summary>
//        /// Fixes PowerShell execution policy by setting to RemoteSigned
//        /// </summary>
//        /// <exception cref="SecurityException">Thrown when lacking sufficient privileges to modify policy</exception>
//        public static void FixExecutionPolicy()
//        {
//            try
//            {
//                _registryService.SetExecutionPolicy("RemoteSigned");
//            }
//            catch (SecurityException ex)
//            {
//                throw new SecurityException("Need administrator privileges to modify the execution policy.", ex);
//            }
//        }

//        /// <summary>
//        /// Retrieves current PowerShell execution policy
//        /// </summary>
//        /// <returns>String representation of current execution policy</returns>
//        public static string GetExecutionPolicy()
//        {
//            return _registryService.GetExecutionPolicy();
//        }

//        /// <summary>
//        /// Checks if current user has administrator privileges
//        /// </summary>
//        /// <returns>True if current user is administrator, otherwise false</returns>
//        public static bool IsAdministrator()
//        {
//            return _registryService.IsAdministrator();
//        }

//        private static bool IsValidExecutionPolicy(string policy)
//        {
//            if (string.IsNullOrEmpty(policy))
//                return false;

//            foreach (var validPolicy in FeasiblePolicies)
//            {
//                if (string.Equals(policy, validPolicy, StringComparison.OrdinalIgnoreCase))
//                    return true;
//            }
//            return false;
//        }

//        /// <summary>
//        /// Checks if quick access query functionality is available
//        /// </summary>
//        public static async Task<bool> CheckQueryFeasible()
//        {
//            try
//            {
//                var result = await ScriptExecutor.ExecutePSScript(PSScript.CheckQueryFeasible, null);
//                return result.ExitCode == 0;
//            }
//            catch
//            {
//                return false;
//            }
//        }

//        /// <summary>
//        /// Checks if pin/unpin functionality is available
//        /// </summary>
//        public static async Task<bool> CheckPinUnpinFeasible()
//        {
//            try
//            {
//                var result = await ScriptExecutor.ExecutePSScript(PSScript.CheckPinUnpinFeasible, null);
//                return result.ExitCode == 0;
//            }
//            catch
//            {
//                return false;
//            }
//        }
//    }

//    public static class PowerShellFeasible
//    {
//        // Internal interface for dependency injection
//        internal interface IPowerShellService
//        {
//            string ExecuteScript(string script);
//        }

//        // Default implementation
//        private class DefaultPowerShellService : IPowerShellService
//        {
//            public string ExecuteScript(string script)
//            {
//                using (var process = new Process())
//                {
//                    process.StartInfo = new ProcessStartInfo
//                    {
//                        FileName = "powershell.exe",
//                        Arguments = $"-Command \"{script}\"",
//                        RedirectStandardOutput = true,
//                        UseShellExecute = false,
//                        CreateNoWindow = true
//                    };

//                    process.Start();
//                    return process.StandardOutput.ReadToEnd().Trim();
//                }
//            }
//        }

//        // Service instances
//        private static IPowerShellService _powerShellService = new DefaultPowerShellService();

//        // Service replacement method for testing
//        internal static void SetPowerShellService(IPowerShellService service)
//        {
//            _powerShellService = service ?? new DefaultPowerShellService();
//        }

//        // Service reset method for testing
//        internal static void ResetPowerShellService()
//        {
//            _powerShellService = new DefaultPowerShellService();
//        }

//        /// <summary>
//        /// Checks if PowerShell script execution environment is available
//        /// </summary>
//        /// <returns>True if PowerShell environment is available, otherwise false</returns>
//        public static bool CheckScriptExecution()
//        {
//            try
//            {
//                var script = "$PSVersionTable.PSVersion.Major";
//                var output = _powerShellService.ExecuteScript(script);
//                return !string.IsNullOrEmpty(output) && int.TryParse(output, out var version) && version >= 3;
//            }
//            catch
//            {
//                return false;
//            }
//        }
//    }
//}
