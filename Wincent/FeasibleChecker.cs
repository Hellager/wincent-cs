using Microsoft.Win32;
using System.Security;

namespace Wincent
{
    public interface IRegistryKeyProxy : IDisposable
    {
        object? GetValue(string name, object? defaultValue = null);
        void SetValue(string name, object value, RegistryValueKind valueKind = RegistryValueKind.String);
        void Close();
    }

    public interface IRegistryOperations
    {
        IRegistryKeyProxy? OpenCurrentUserSubKey(string path, bool writable);
        IRegistryKeyProxy CreateCurrentUserSubKey(string path);
    }

    internal class RegistryKeyProxy : IRegistryKeyProxy
    {
        private readonly RegistryKey _key;

        public RegistryKeyProxy(RegistryKey key)
        {
            _key = key ?? throw new ArgumentNullException(nameof(key));
        }

        public object? GetValue(string name, object? defaultValue = null)
            => _key.GetValue(name, defaultValue);

        public void SetValue(string name, object value, RegistryValueKind valueKind = RegistryValueKind.String)
            => _key.SetValue(name, value, valueKind);

        public void Close()
            => _key.Close();

        public void Dispose()
            => _key.Dispose();
    }


    internal class DefaultRegistryOperations : IRegistryOperations
    {
        public IRegistryKeyProxy? OpenCurrentUserSubKey(string path, bool writable)
        {
            var key = Registry.CurrentUser.OpenSubKey(path, writable);
            return key != null ? new RegistryKeyProxy(key) : null;
        }

        public IRegistryKeyProxy CreateCurrentUserSubKey(string path)
        {
            var key = Registry.CurrentUser.CreateSubKey(path);
            return new RegistryKeyProxy(key);
        }
    }

    public static class FeasibleChecker
    {
        private static readonly object _syncRoot = new object();
        private static volatile IRegistryOperations _registry = new DefaultRegistryOperations();
        private static string _executionPolicyKeyPath = @"SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell";
        private static readonly string[] FeasiblePolicies = ["AllSigned", "Bypass", "RemoteSigned", "Unrestricted"];

        public static void InjectDependencies(
            IRegistryOperations registryOps)
        {
            ArgumentNullException.ThrowIfNull(registryOps);

            lock (_syncRoot)
            {
                _registry = registryOps;
                Thread.MemoryBarrier();
            }
        }

        public static void ResetDependencies()
        {
            lock (_syncRoot)
            {
                _registry = new DefaultRegistryOperations();
            }
        }

        public static IRegistryOperations GetCurrentRegistry()
        {
            lock (_syncRoot)
            {
                return _registry;
            }
        }

        private static IRegistryKeyProxy GetExecutionPolicyRegistryKey(bool writable = true)
        {
            IRegistryOperations registry = GetCurrentRegistry();

            try
            {
                IRegistryKeyProxy key = registry.OpenCurrentUserSubKey(_executionPolicyKeyPath, writable)
                               ?? registry.CreateCurrentUserSubKey(_executionPolicyKeyPath);
                return key ?? throw new InvalidOperationException("Unable to create or open registry key");
            }
            catch (Exception ex) when (ex is SecurityException || ex is UnauthorizedAccessException)
            {
                throw new RegistryAccessException("Registry access denied, please check permissions", ex);
            }
            catch (IOException ex)
            {
                throw new RegistryOperationException("Registry I/O operation failed", ex);
            }
            catch (ArgumentException ex)
            {
                throw new InvalidRegistryPathException("Invalid registry path format", ex);
            }
        }

        public static bool CheckScriptFeasible()
        {
            try
            {
                using var regKey = GetExecutionPolicyRegistryKey(false);
                var value = regKey.GetValue("ExecutionPolicy", "NotSet") as string;

                if (string.IsNullOrEmpty(value))
                    return false;

                var cleanValue = new string(value.Where(c => c != '\0').ToArray());
                return FeasiblePolicies.Contains(cleanValue ?? "", StringComparer.OrdinalIgnoreCase);
            }
            catch (RegistryAccessException)
            {
                return false;
            }
        }

        public static void FixExecutionPolicy()
        {
            try
            {
                using var regKey = GetExecutionPolicyRegistryKey();
                regKey.SetValue("ExecutionPolicy", "RemoteSigned", RegistryValueKind.String);
            }
            catch (ArgumentException ex) when (ex.ParamName == "value")
            {
                throw new InvalidRegistryValueException("Invalid registry value type", ex);
            }
        }

        public static bool CheckQueryFeasibleWithScript()
        {
            try
            {
                var result = ScriptExecutor.ExecutePSScript(PSScript.CheckQueryFeasible, null).Result;
                return result?.ExitCode == 0 && string.IsNullOrEmpty(result.Error);
            }
            catch (AggregateException ae) when (ae.InnerException is ScriptExecutionException)
            {
                throw new ScriptValidationException("Script execution failed", ae.InnerException);
            }
            catch (NullReferenceException)
            {
                throw new InvalidOperationException("Script execution result is abnormal");
            }
        }

        public static bool CheckPinUnpinFeasibleWithScript()
        {
            try
            {
                var result = ScriptExecutor.ExecutePSScript(PSScript.CheckPinUnpinFeasible, null).Result;
                return result?.ExitCode == 0 && string.IsNullOrEmpty(result.Error);
            }
            catch (AggregateException ae) when (ae.InnerException is ScriptExecutionException)
            {
                throw new ScriptValidationException("Script execution failed", ae.InnerException);
            }
            catch (NullReferenceException)
            {
                throw new InvalidOperationException("Script execution result is abnormal");
            }
        }

        public static bool RegistryPathExists(string registryPath, string rootKey = "HKEY_CURRENT_USER")
        {
            if (string.IsNullOrWhiteSpace(registryPath))
                throw new ArgumentException("Registry path cannot be empty", nameof(registryPath));

            var parts = registryPath.Split('\\');
            if (parts.Length < 1)
                throw new InvalidRegistryPathException("Path cannot be empty");

            try
            {
                var root = rootKey switch
                {
                    "HKEY_CURRENT_USER" => Registry.CurrentUser,
                    "HKEY_LOCAL_MACHINE" => Registry.LocalMachine,
                    _ => throw new InvalidRegistryPathException($"Unsupported root key: {rootKey}")
                };

                using var key = root.OpenSubKey(registryPath, false);
                return key != null;
            }
            catch (SecurityException ex)
            {
                throw new RegistryAccessException($"No permission to access registry path: {registryPath}", ex);
            }
            catch (IOException ex)
            {
                throw new RegistryOperationException($"Registry I/O error: {registryPath}", ex);
            }
            catch (ArgumentException ex)
            {
                throw new InvalidRegistryPathException($"Invalid path: {registryPath}", ex);
            }
        }

        public static string GetExecutionPolicy()
        {
            try
            {
                using var regKey = GetExecutionPolicyRegistryKey(false);
                var currentRegistry = GetCurrentRegistry();

                var value = regKey.GetValue("ExecutionPolicy", defaultValue: "NotSet") as string;

                return new string(value?.Where(c => c != '\0').ToArray() ?? Array.Empty<char>());
            }
            catch (RegistryAccessException ex)
            {
                return $"Error: {ex.Message}";
            }
            catch (ObjectDisposedException ex)
            {
                throw new InvalidOperationException("Registry key has been disposed", ex);
            }
        }
    }

    public class RegistryAccessException(string message, Exception inner) : Exception(message, inner)
    {
    }

    public class RegistryOperationException(string message, Exception inner) : Exception(message, inner)
    {
    }

    public class InvalidRegistryPathException : Exception
    {
        public InvalidRegistryPathException(string message)
            : base(message)
        { }

        public InvalidRegistryPathException(string message, Exception inner) : base(message, inner) { }
    }

    public class ScriptValidationException(string message, Exception inner) : Exception(message, inner)
    {
    }

    public class InvalidRegistryValueException(string message, Exception inner) : Exception(message, inner)
    {
    }
}
