using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Security;
using System.Security.Principal;

namespace Wincent
{
    public static class ExecutionFeasible
    {
        private static readonly string ExecutionPolicyKeyPath = @"SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell";
        private static readonly string[] FeasiblePolicies = { "AllSigned", "Bypass", "RemoteSigned", "Unrestricted" };

        public static bool CheckScriptFeasible()
        {
            try
            {
                using (var regKey = OpenExecutionPolicyKey(false))
                {
                    var value = regKey?.GetValue("ExecutionPolicy", "NotSet") as string;
                    return IsValidExecutionPolicy(value);
                }
            }
            catch (SecurityException)
            {
                return false;
            }
        }

        public static void FixExecutionPolicy()
        {
            try
            {
                using (var regKey = OpenExecutionPolicyKey(true))
                {
                    regKey?.SetValue("ExecutionPolicy", "RemoteSigned", RegistryValueKind.String);
                }
            }
            catch (SecurityException ex)
            {
                throw new SecurityException("Need administrator privileges to modify the execution policy.", ex);
            }
        }

        public static string GetExecutionPolicy()
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

        public static bool IsAdministrator()
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

        #region Helper Methods
        private static RegistryKey OpenExecutionPolicyKey(bool writable)
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
        #endregion
    }

    public static class PowerShellFeasible
    {
        public static bool CheckScriptExecution()
        {
            try
            {
                var script = "$PSVersionTable.PSVersion.Major";
                var output = ExecutePowerShellScript(script);
                return !string.IsNullOrEmpty(output) && int.TryParse(output, out var version) && version >= 3;
            }
            catch
            {
                return false;
            }
        }

        private static string ExecutePowerShellScript(string script)
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
    }
}
