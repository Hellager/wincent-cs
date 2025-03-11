using System.Diagnostics;
using System.Text;
using System.Collections.Concurrent;

namespace Wincent
{
    public record ScriptResult(
        int ExitCode,
        string Output,
        string Error);

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
            : base(message, null!, "", "")
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

    public sealed class ScriptExecutor(IPSScriptStrategyFactory? strategyFactory = null)
    {
        private readonly IPSScriptStrategyFactory _strategyFactory = strategyFactory ?? new DefaultPSScriptStrategyFactory();
        private static readonly ConcurrentDictionary<PSScript, string> ScriptCache = new();

        public string GetScriptContent(PSScript method, string? parameter)
        {
            if (!Enum.IsDefined(typeof(PSScript), method))
                throw new ArgumentOutOfRangeException(nameof(method), "Invalid script type");

            try
            {
                var strategy = _strategyFactory.GetStrategy(method);

                return strategy.GenerateScript(parameter);
            }
            catch (NotSupportedException ex)
            {
                throw new ScriptGenerationException($"Script generation failed: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Execute PowerShell script (text content)
        /// </summary>
        /// <param name="scriptContent">Script content</param>
        /// <param name="timeout">Timeout (unlimited by default)</param>
        public async Task<ScriptResult> ExecutePowerShellScriptAsync(
            string scriptContent,
            TimeSpan? timeout = null)
        {
            byte[] contentBytes = AddUtf8Bom(Encoding.UTF8.GetBytes(scriptContent));
            using var tempFile = TempFile.Create(contentBytes, "ps1");
            return await ExecuteCoreAsync(tempFile.FullPath, timeout);
        }

        /// <summary>
        /// Execute PowerShell script (binary content)
        /// </summary>
        /// <param name="scriptBytes">Script byte array</param>
        /// <param name="extension">File extension</param>
        /// <param name="timeout">Timeout</param>
        public async Task<ScriptResult> ExecutePowerShellScriptAsync(
            byte[] scriptBytes,
            string extension = "ps1",
            TimeSpan? timeout = null)
        {
            byte[] contentWithBom = AddUtf8Bom(scriptBytes);
            using var tempFile = TempFile.Create(contentWithBom, extension);
            return await ExecuteCoreAsync(tempFile.FullPath, timeout);
        }

        /// <summary>
        /// Ensure byte array starts with UTF8-BOM
        /// </summary>
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

        private async Task<ScriptResult> ExecuteCoreAsync(
            string scriptPath,
            TimeSpan? timeout)
        {
            var processStartInfo = new ProcessStartInfo
            {
                FileName = GetPowerShellPath(),
                Arguments = $"-ExecutionPolicy Bypass -File \"{scriptPath}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            var output = new StringBuilder();
            var error = new StringBuilder();

            using var process = new Process { StartInfo = processStartInfo };
            using var cts = timeout.HasValue
                ? new CancellationTokenSource(timeout.Value)
                : new CancellationTokenSource();

            try
            {
                process.Start();

                var readOutputTask = ReadStreamAsync(process.StandardOutput);
                var readErrorTask = ReadStreamAsync(process.StandardError);

                try
                {
                    await process.WaitForExitAsync(cts.Token);

                    output.Append(await readOutputTask);
                    error.Append(await readErrorTask);
                }
                catch (OperationCanceledException)
                {
                    if (!process.HasExited)
                    {
                        process.Kill();
                        await Task.Delay(100);
                    }

                    using var outputCts = new CancellationTokenSource(TimeSpan.FromMilliseconds(100));
                    try
                    {
                        output.Append(await readOutputTask.WaitAsync(outputCts.Token));
                        error.Append(await readErrorTask.WaitAsync(outputCts.Token));
                    }
                    catch (OperationCanceledException)
                    {
                        // If reading times out, use existing content
                    }

                    throw new ScriptTimeoutException(
                        $"Script execution timed out ({timeout})",
                        output.ToString(),
                        error.ToString().Trim());
                }

                int exitCode = process.ExitCode;
                string errorOutput = error.ToString().Trim();
                if (exitCode == 0 && !string.IsNullOrEmpty(errorOutput))
                {
                    exitCode = -1;
                }

                return new ScriptResult(
                    exitCode,
                    output.ToString(),
                    errorOutput);
            }
            catch (Exception ex) when (!(ex is OperationCanceledException || ex is ScriptTimeoutException))
            {
                try
                {
                    if (!process.HasExited)
                        process.Kill();
                }
                catch { /* Ignore errors when terminating process */ }

                throw new ScriptExecutionException(
                    "Script execution failed",
                    ex,
                    output.ToString(),
                    error.ToString().Trim());
            }
        }

        private async Task<string> ReadStreamAsync(StreamReader reader)
        {
            var output = new StringBuilder();
            char[] buffer = new char[1024];

            while (true)
            {
                int bytesRead = await reader.ReadAsync(buffer);
                if (bytesRead == 0) break;
                output.Append(buffer, 0, bytesRead);
            }

            return output.ToString();
        }

        private static string GetPowerShellPath()
        {
            return OperatingSystem.IsWindows()
                ? "powershell.exe"
                : "pwsh";
        }

        internal static bool FileOrDirectoryExists(string name)
        {
            return (Directory.Exists(name) || File.Exists(name));
        }

        public static async Task<ScriptResult> ExecutePSScript(PSScript method, string? para)
        {
            var executor = new ScriptExecutor();

            if (para != null)
            {
                var isValid = FileOrDirectoryExists(para);
                if (!isValid)
                {
                    throw new InvalidPathException(para, "File or directory does not exist");
                }
            }

            var script = executor.GetScriptContent(method, para);
            byte[] data = Encoding.UTF8.GetBytes(script);

            try
            {
                var result = await executor.ExecutePowerShellScriptAsync(data, "ps1", TimeSpan.FromSeconds(5));
                return result;
            }
            catch (ScriptExecutionException ex)
            {
                return new ScriptResult(-1, "", ex.Message);
            }
        }
    }
}
