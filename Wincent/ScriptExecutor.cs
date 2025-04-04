using System;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.IO;

namespace Wincent
{
    public enum ExecutionPolicy
    {
        Restricted,
        AllSigned,
        RemoteSigned,
        Unrestricted,
        Bypass,
        Undefined
    }

    [Serializable]
    public class ScriptResult
    {
        public int ExitCode { get; }
        public string Output { get; }
        public string Error { get; }

        public ScriptResult(int exitCode, string output, string error)
        {
            ExitCode = exitCode;
            Output = output;
            Error = error;
        }
    }

    public class ScriptExecutionException : Exception
    {
        public string Output { get; }
        public string Error { get; }

        public ScriptExecutionException(
            string message,
            Exception inner,
            string output,
            string error)
            : base(message, inner)
        {
            Output = output;
            Error = error;
        }

        public ScriptExecutionException(
            string message,
            string output,
            string error)
            : base(message)
        {
            Output = output;
            Error = error;
        }
    }

    public class ScriptTimeoutException : Exception
    {
        public string Output { get; }
        public string Error { get; }

        public ScriptTimeoutException(
            string message,
            string output,
            string error)
            : base(message)
        {
            Output = output;
            Error = error;
        }
    }

    public class ScriptProcessException : ScriptExecutionException
    {
        public int ProcessId { get; }

        public ScriptProcessException(int processId, string message)
            : base(message, null, "", "")
            => ProcessId = processId;
    }

    public class InvalidPathException : Exception
    {
        public string Path { get; }

        public InvalidPathException(string path, string message)
            : base(message)
        {
            Path = path;
        }

        public InvalidPathException(string path, string message, Exception inner)
            : base(message, inner)
        {
            Path = path;
        }
    }

    public sealed class ScriptExecutor : IDisposable
    {
        private static readonly ScriptResultCache _cache = new ScriptResultCache();

        /// <summary>
        /// File system service interface for dependency injection
        /// </summary>
        internal interface IFileSystemService
        {
            bool FileExists(string path);
            bool DirectoryExists(string path);
        }

        /// <summary>
        /// Default file system service implementation
        /// </summary>
        private class DefaultFileSystemService : IFileSystemService
        {
            public bool FileExists(string path)
            {
                return File.Exists(path);
            }

            public bool DirectoryExists(string path)
            {
                return Directory.Exists(path);
            }
        }

        private static IFileSystemService _fileSystemService = new DefaultFileSystemService();

        /// <summary>
        /// Set file system service (for testing)
        /// </summary>
        internal static void SetFileSystemService(IFileSystemService service)
        {
            _fileSystemService = service ?? throw new ArgumentNullException(nameof(service));
        }

        /// <summary>
        /// Reset to default file system service
        /// </summary>
        internal static void ResetFileSystemService()
        {
            _fileSystemService = new DefaultFileSystemService();
        }

        public static byte[] AddUtf8Bom(byte[] content)
        {
            byte[] bom = Encoding.UTF8.GetPreamble();

            if (content.Length >= bom.Length)
            {
                bool hasBom = true;
                for (int i = 0; i < bom.Length; i++)
                {
                    if (content[i] != bom[i])
                    {
                        hasBom = false;
                        break;
                    }
                }

                if (hasBom)
                    return content;
            }

            byte[] result = new byte[bom.Length + content.Length];
            Buffer.BlockCopy(bom, 0, result, 0, bom.Length);
            Buffer.BlockCopy(content, 0, result, bom.Length, content.Length);

            return result;
        }

        private async Task<ScriptResult> ExecuteCoreAsync(string scriptPath, TimeSpan? timeout = null)
        {
            if (!File.Exists(scriptPath))
                throw new FileNotFoundException("Script file not found", scriptPath);

            var maxRunspaces = 5;
            using (var runspacePool = RunspaceFactory.CreateRunspacePool())
            {
                runspacePool.SetMaxRunspaces(maxRunspaces);
                runspacePool.Open();

                using (var ps = PowerShell.Create())
                {
                    ps.RunspacePool = runspacePool;
                    ps.AddCommand(scriptPath);

                    // Thread-safe data collectors
                    var outputStream = new PSDataCollection<PSObject>();
                    var errorStream = new PSDataCollection<ErrorRecord>();
                    var output = new StringBuilder();
                    var errors = new StringBuilder();
                    var hadErrors = false;

                    // Correct BeginInvoke parameter signature
                    IAsyncResult asyncResult = ps.BeginInvoke<PSObject, PSObject>(
                        input: null,
                        output: outputStream,
                        settings: null,  // Explicit parameter naming
                        callback: null,
                        state: null);

                    // Configure async processing
                    var invokeTask = Task.Factory.FromAsync(asyncResult, _ => ps.EndInvoke(_));
                    var timeoutTask = timeout.HasValue
                        ? Task.Delay(timeout.Value)
                        : Task.Delay(-1);

                    var completedTask = await Task.WhenAny(invokeTask, timeoutTask);

                    // Handle timeout
                    if (completedTask == timeoutTask)
                    {
                        ps.Stop();
                        throw new TimeoutException($"Execution timed out after {timeout.Value.TotalSeconds}s");
                    }

                    // Collect final results
                    foreach (var item in outputStream)
                        output.AppendLine(item.ToString());

                    foreach (var error in errorStream.ReadAll())
                    {
                        hadErrors = true;
                        errors.AppendLine(error.ToString());
                    }

                    return new ScriptResult(
                        hadErrors ? 1 : 0,
                        output.ToString(),
                        errors.ToString());
                }
            }
        }

        internal static bool FileOrDirectoryExists(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return false;

            try
            {
                return _fileSystemService.FileExists(name) || _fileSystemService.DirectoryExists(name);
            }
            catch
            {
                return false;
            }
        }

        public static async Task<ScriptResult> ExecutePSScript(PSScript method, string para)
        {
            try
            {
                // Try get from cache first
                var cachedResult = _cache.GetCachedResult(method);
                if (cachedResult != null)
                {
                    return cachedResult;
                }

                if (para != null)
                {
                    var isValid = FileOrDirectoryExists(para);
                    if (!isValid)
                    {
                        throw new InvalidPathException(para, "File or directory does not exist");
                    }
                }

                // Get script path
                string scriptPath;
                if (para != null && ScriptStorage.IsParameterizedScript(method))
                {
                    scriptPath = ScriptStorage.GetDynamicScriptPath(method, para);
                }
                else
                {
                    scriptPath = ScriptStorage.GetScriptPath(method);
                }

                // Execute script
                var executor = new ScriptExecutor();
                try
                {
                    var result = await executor.ExecuteCoreAsync(scriptPath, TimeSpan.FromSeconds(5));
                    
                    // Update cache if successful
                    if (result.ExitCode == 0)
                    {
                        _cache.UpdateCache(method, result);
                    }
                    
                    return result;
                }
                catch (ScriptExecutionException ex)
                {
                    return new ScriptResult(-1, "", ex.Message);
                }
            }
            catch (Exception ex)
            {
                return new ScriptResult(-1, "", ex.Message);
            }
            finally
            {
                // Cleanup expired dynamic scripts
                if (ScriptStorage.IsParameterizedScript(method) && para != null)
                {
                    ScriptStorage.CleanupDynamicScripts();
                }
            }
        }

        public void Dispose()
        {
        }
    }
}
