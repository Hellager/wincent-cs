using System;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;

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

    /// <summary>
    /// Cache key for storage
    /// </summary>
    internal class CacheKey : IEquatable<CacheKey>
    {
        public PSScript ScriptType { get; }
        public string Parameter { get; }

        public CacheKey(PSScript scriptType, string parameter)
        {
            ScriptType = scriptType;
            Parameter = parameter;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as CacheKey);
        }

        public bool Equals(CacheKey other)
        {
            if (other == null)
                return false;

            return ScriptType == other.ScriptType &&
                   Parameter == other.Parameter;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = hash * 23 + ScriptType.GetHashCode();
                hash = hash * 23 + (Parameter?.GetHashCode() ?? 0);
                return hash;
            }
        }
    }

    /// <summary>
    /// Cache entry
    /// </summary>
    internal class CacheEntry
    {
        public List<string> Result { get; }
        public DateTime Timestamp { get; }

        public CacheEntry(List<string> result, DateTime timestamp)
        {
            Result = result;
            Timestamp = timestamp;
        }
    }

    /// <summary>
    /// File system service interface for dependency injection
    /// </summary>
    public interface IFileSystemService
    {
        bool FileExists(string path);
        bool DirectoryExists(string path);
    }

    /// <summary>
    /// Default file system service implementation
    /// </summary>
    public class DefaultFileSystemService : IFileSystemService
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

    /// <summary>
    /// Script storage service interface for dependency injection
    /// </summary>
    public interface IScriptStorageService
    {
        string GetScriptPath(PSScript script);
        string GetDynamicScriptPath(PSScript script, string parameter);
        bool IsParameterizedScript(PSScript script);
        void CleanupDynamicScripts();
    }

    /// <summary>
    /// Default script storage service implementation
    /// </summary>
    public class DefaultScriptStorageService : IScriptStorageService
    {
        public string GetScriptPath(PSScript script)
        {
            return ScriptStorage.GetScriptPath(script);
        }

        public string GetDynamicScriptPath(PSScript script, string parameter)
        {
            return ScriptStorage.GetDynamicScriptPath(script, parameter);
        }

        public bool IsParameterizedScript(PSScript script)
        {
            return ScriptStorage.IsParameterizedScript(script);
        }

        public void CleanupDynamicScripts()
        {
            ScriptStorage.CleanupDynamicScripts();
        }
    }

    public sealed class ScriptExecutor : IDisposable
    {
        private readonly IFileSystemService _fileSystemService;
        private readonly IScriptStorageService _scriptStorageService;
        private readonly ConcurrentDictionary<CacheKey, CacheEntry> _cache = new ConcurrentDictionary<CacheKey, CacheEntry>();
        private readonly QuickAccessDataFiles _dataFiles;

        /// <summary>
        /// Creates ScriptExecutor instance with default dependencies
        /// </summary>
        public ScriptExecutor()
            : this(new DefaultFileSystemService(), new DefaultScriptStorageService(), new QuickAccessDataFiles())
        {
        }

        /// <summary>
        /// Creates ScriptExecutor instance with specified dependencies
        /// </summary>
        /// <param name="fileSystemService">File system service</param>
        /// <param name="scriptStorageService">Script storage service</param>
        /// <param name="dataFiles">Quick access data files</param>
        public ScriptExecutor(
            IFileSystemService fileSystemService,
            IScriptStorageService scriptStorageService,
            QuickAccessDataFiles dataFiles)
        {
            _fileSystemService = fileSystemService ?? throw new ArgumentNullException(nameof(fileSystemService));
            _scriptStorageService = scriptStorageService ?? throw new ArgumentNullException(nameof(scriptStorageService));
            _dataFiles = dataFiles ?? throw new ArgumentNullException(nameof(dataFiles));
        }

        /// <summary>
        /// Determines if script type should be cached
        /// </summary>
        private bool ShouldCache(PSScript scriptType)
        {
            // Only cache query-type scripts
            return scriptType == PSScript.QueryQuickAccess ||
                   scriptType == PSScript.QueryRecentFile ||
                   scriptType == PSScript.QueryFrequentFolder;
        }

        public async Task<ScriptResult> ExecuteCoreAsync(string scriptPath, TimeSpan? timeout = null, string parameter = null)
        {
            if (!_fileSystemService.FileExists(scriptPath))
                throw new FileNotFoundException("Script file not found", scriptPath);

            string arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -File \"{scriptPath}\"";

            // Add parameter to command line arguments if provided
            if (!string.IsNullOrEmpty(parameter))
            {
                arguments += $" \"{parameter}\"";
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            using (var process = new Process { StartInfo = startInfo })
            {
                var output = new StringBuilder();
                var error = new StringBuilder();
                var outputTask = new TaskCompletionSource<bool>();
                var errorTask = new TaskCompletionSource<bool>();

                process.OutputDataReceived += (sender, e) =>
                {
                    if (e.Data == null)
                        outputTask.TrySetResult(true);
                    else
                        output.AppendLine(e.Data);
                };

                process.ErrorDataReceived += (sender, e) =>
                {
                    if (e.Data == null)
                        errorTask.TrySetResult(true);
                    else
                        error.AppendLine(e.Data);
                };

                try
                {
                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();

                    var processTask = Task.Run(() => process.WaitForExit());
                    var timeoutTask = timeout.HasValue
                        ? Task.Delay(timeout.Value)
                        : Task.Delay(-1);

                    // Wait for process completion or timeout
                    var completedTask = await Task.WhenAny(processTask, timeoutTask);

                    if (completedTask == timeoutTask && timeout.HasValue)
                    {
                        try
                        {
                            // Attempt graceful termination
                            if (!process.HasExited)
                            {
                                process.Kill();
                                await Task.Delay(100); // Allow time for final output collection
                            }
                        }
                        catch
                        {
                            // Ignore errors during process termination
                        }

                        throw new ScriptTimeoutException(
                            $"Script execution timeout ({timeout.Value.TotalSeconds} seconds)",
                            output.ToString(),
                            error.ToString());
                    }

                    // Wait for output streams to complete
                    await Task.WhenAll(outputTask.Task, errorTask.Task);

                    // Retrieve exit code
                    int exitCode = process.ExitCode;

                    return new ScriptResult(
                        exitCode,
                        output.ToString(),
                        error.ToString());
                }
                catch (Exception ex) when (!(ex is ScriptTimeoutException))
                {
                    // Handle non-timeout exceptions
                    try
                    {
                        if (!process.HasExited)
                        {
                            process.Kill();
                            await Task.Delay(100); // Allow time for final output collection
                        }
                    }
                    catch
                    {
                        // Ignore termination errors
                    }

                    string errorMessage = error.ToString();
                    if (string.IsNullOrEmpty(errorMessage))
                    {
                        errorMessage = $"Execution error: {ex.GetType().Name}: {ex.Message}";
                        if (ex.InnerException != null)
                        {
                            errorMessage += $"\nInner exception: {ex.InnerException.Message}";
                        }
                    }

                    return new ScriptResult(-1, output.ToString(), errorMessage);
                }
            }
        }

        /// <summary>
        /// Checks existence of file or directory
        /// </summary>
        public bool FileOrDirectoryExists(string name)
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

        /// <summary>
        /// Parses script output into string list
        /// </summary>
        private List<string> ParseScriptOutputToList(string output)
        {
            if (string.IsNullOrEmpty(output))
                return new List<string>();

            return new List<string>(
                output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                      .Select(s => s.Trim())
                      .Where(s => !string.IsNullOrEmpty(s))
            );
        }

        /// <summary>
        /// Executes PowerShell script with caching and optional timeout
        /// </summary>
        /// <param name="method">Script type</param>
        /// <param name="para">Script parameter (required for parameterized scripts)</param>
        /// <param name="timeoutSeconds">Timeout in seconds (0 or negative for no timeout)</param>
        /// <returns>Script execution result as string list</returns>
        public async Task<List<string>> ExecutePSScriptWithCache(PSScript method, string para, int timeoutSeconds = 0)
        {
            // Bypass cache for non-query scripts
            if (!ShouldCache(method))
            {
                ScriptResult result;
                if (timeoutSeconds > 0)
                {
                    result = await ExecutePSScriptWithTimeout(method, para, timeoutSeconds);
                }
                else
                {
                    result = await ExecutePSScript(method, para);
                }
                return ParseScriptOutputToList(result.Output);
            }

            // Create cache key for query
            var key = new CacheKey(method, para);

            // Get data file modified time for cache validation
            DateTime currentModifiedTime = _dataFiles.GetModifiedTimeForScript(method);

            // Check cache
            if (_cache.TryGetValue(key, out CacheEntry entry))
            {
                // Validate cache using modification timestamp
                if (entry.Timestamp >= currentModifiedTime)
                {
                    return entry.Result;
                }
            }

            // Cache miss: execute and store
            ScriptResult scriptResult;
            if (timeoutSeconds > 0)
            {
                scriptResult = await ExecutePSScriptWithTimeout(method, para, timeoutSeconds);
            }
            else
            {
                scriptResult = await ExecutePSScript(method, para);
            }

            if (scriptResult.ExitCode != 0)
            {
                throw new ScriptExecutionException(
                    "Script execution failed with non-zero exit code",
                    scriptResult.Output,
                    scriptResult.Error);
            }

            var parsedResult = ParseScriptOutputToList(scriptResult.Output);

            // Update cache
            _cache[key] = new CacheEntry(parsedResult, currentModifiedTime);

            return parsedResult;
        }

        /// <summary>
        /// Executes PowerShell script without timeout
        /// </summary>
        /// <param name="method">Script type</param>
        /// <param name="para">Script parameter (required for parameterized scripts)</param>
        /// <returns>Script execution result</returns>
        public async Task<ScriptResult> ExecutePSScript(PSScript method, string para)
        {
            try
            {
                // Validate parameter
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
                if (para != null && _scriptStorageService.IsParameterizedScript(method))
                {
                    scriptPath = _scriptStorageService.GetDynamicScriptPath(method, para);
                }
                else
                {
                    scriptPath = _scriptStorageService.GetScriptPath(method);
                }

                // Execute without timeout
                var result = await ExecuteCoreAsync(scriptPath, null, para);
                return result;
            }
            catch (InvalidPathException)
            {
                // Re-throw for caller handling
                throw;
            }
            catch (ScriptExecutionException)
            {
                // Re-throw for caller handling
                throw;
            }
            catch (ScriptTimeoutException)
            {
                // Re-throw for caller handling
                throw;
            }
            catch (Exception ex)
            {
                // Handle other exceptions
                return new ScriptResult(-1, "", ex.Message);
            }
            finally
            {
                // Cleanup dynamic scripts
                if (_scriptStorageService.IsParameterizedScript(method) && para != null)
                {
                    _scriptStorageService.CleanupDynamicScripts();
                }
            }
        }

        /// <summary>
        /// Executes PowerShell script with custom timeout
        /// </summary>
        /// <param name="method">Script type</param>
        /// <param name="para">Script parameter (required for parameterized scripts)</param>
        /// <param name="timeoutSeconds">Timeout in seconds (0 for no timeout)</param>
        /// <returns>Script execution result</returns>
        public async Task<ScriptResult> ExecutePSScriptWithTimeout(PSScript method, string para, int timeoutSeconds)
        {
            try
            {
                // Validate parameter
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
                if (para != null && _scriptStorageService.IsParameterizedScript(method))
                {
                    scriptPath = _scriptStorageService.GetDynamicScriptPath(method, para);
                }
                else
                {
                    scriptPath = _scriptStorageService.GetScriptPath(method);
                }

                // Set timeout duration
                TimeSpan? timeout = null;
                if (timeoutSeconds > 0)
                {
                    timeout = TimeSpan.FromSeconds(timeoutSeconds);
                }

                // Execute with timeout
                var result = await ExecuteCoreAsync(scriptPath, timeout, para);
                return result;
            }
            catch (InvalidPathException)
            {
                // Re-throw for caller handling
                throw;
            }
            catch (ScriptTimeoutException ex)
            {
                // Return timeout result without throwing
                return new ScriptResult(-1, ex.Output, $"Timeout: {ex.Message}");
            }
            catch (ScriptExecutionException)
            {
                // Re-throw for caller handling
                throw;
            }
            catch (Exception ex)
            {
                // Handle other exceptions
                return new ScriptResult(-1, "", ex.Message);
            }
            finally
            {
                // Cleanup dynamic scripts
                if (_scriptStorageService.IsParameterizedScript(method) && para != null)
                {
                    _scriptStorageService.CleanupDynamicScripts();
                }
            }
        }

        /// <summary>
        /// Clears script execution cache
        /// </summary>
        public void ClearCache()
        {
            _cache.Clear();
        }

        public void Dispose()
        {
            // Clean resources
            ClearCache();
        }
    }
}
