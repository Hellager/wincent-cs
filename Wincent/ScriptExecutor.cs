using System;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Collections.Concurrent;
using System.IO;
using Microsoft.PowerShell;

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
        private readonly IPSScriptStrategyFactory _strategyFactory;
        private static readonly ConcurrentDictionary<PSScript, string> ScriptCache = new ConcurrentDictionary<PSScript, string>();
        private readonly RunspacePool _runspacePool;

        public ScriptExecutor(IPSScriptStrategyFactory strategyFactory = null)
        {
            _strategyFactory = strategyFactory ?? new DefaultPSScriptStrategyFactory();
            InitialSessionState initialSessionState = InitialSessionState.CreateDefault();
            initialSessionState.Commands.Add(new SessionStateCmdletEntry(
                "Set-ExecutionPolicy",
                typeof(SetExecutionPolicyCommand),
                "Set-ExecutionPolicy Bypass -Scope Process"));

            _runspacePool = RunspaceFactory.CreateRunspacePool(initialSessionState);
            _runspacePool.SetMaxRunspaces(5);
            _runspacePool.ThreadOptions = PSThreadOptions.UseNewThread;
            _runspacePool.Open();
        }

        public class SetExecutionPolicyCommand : PSCmdlet
        {
            [Parameter(Position = 0)]
            public ExecutionPolicy ExecutionPolicy { get; set; }

            [Parameter(Position = 1)]
            public ExecutionPolicyScope Scope { get; set; } = ExecutionPolicyScope.Process;

            protected override void ProcessRecord()
            {
            
            }
        }

        public string GetScriptContent(PSScript method, string parameter)
        {
            if (!Enum.IsDefined(typeof(PSScript), method))
                throw new ArgumentOutOfRangeException(nameof(method), "Invalid script type");

            try
            {
                return ScriptCache.GetOrAdd(method, key => {
                    var strategy = _strategyFactory.GetStrategy(key);
                    return strategy.GenerateScript(parameter);
                });
            }
            catch (NotSupportedException ex)
            {
                throw new ScriptGenerationException($"Script generation failed: {ex.Message}", ex);
            }
        }

        public async Task<ScriptResult> ExecutePowerShellScriptAsync(
            string scriptContent,
            TimeSpan? timeout = null)
        {
            if (string.IsNullOrWhiteSpace(scriptContent))
                throw new ArgumentNullException(nameof(scriptContent));

            byte[] contentBytes = AddUtf8Bom(Encoding.UTF8.GetBytes(scriptContent));
            using (var tempFile = TempFile.Create(contentBytes, "ps1"))
            {
                return await ExecuteCoreAsync(tempFile.FullPath, timeout);
            }
        }

        public async Task<ScriptResult> ExecutePowerShellScriptAsync(
            byte[] scriptBytes,
            string extension = "ps1",
            TimeSpan? timeout = null)
        {
            byte[] contentWithBom = AddUtf8Bom(scriptBytes);
            using (var tempFile = TempFile.Create(contentWithBom, extension))
            {
                return await ExecuteCoreAsync(tempFile.FullPath, timeout);
            }
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

        private async Task<ScriptResult> ExecuteCoreAsync(
            string scriptPath,
            TimeSpan? timeout)
        {
            using (var powerShell = PowerShell.Create())
            {
                powerShell.RunspacePool = _runspacePool;
                powerShell.AddCommand(scriptPath);

                var output = new PSDataCollection<PSObject>();
                var errors = new PSDataCollection<ErrorRecord>();
                var outputResult = new StringBuilder();
                var errorResult = new StringBuilder();

                var asyncResult = powerShell.BeginInvoke<PSObject, PSObject>(null, output);

                output.DataAdded += (s, e) => CaptureOutput(output, outputResult);
                errors.DataAdded += (s, e) => CaptureErrors(errors, errorResult);

                try
                {
                    await WaitForCompletion(asyncResult, timeout);
                }
                catch (TimeoutException)
                {
                    powerShell.Stop();
                    throw new ScriptTimeoutException(
                        $"Execution timed out after {timeout}",
                        outputResult.ToString(),
                        errorResult.ToString());
                }
                catch (RuntimeException ex)
                {
                    throw new ScriptExecutionException(
                        "PowerShell runtime error",
                        ex,
                        outputResult.ToString(),
                        errorResult.ToString());
                }

                return new ScriptResult(
                    powerShell.HadErrors ? -1 : 0,
                    outputResult.ToString(),
                    errorResult.ToString());
            }
        }

        internal static bool FileOrDirectoryExists(string name)
        {
            try
            {
                return File.GetAttributes(name) != (FileAttributes)(-1);
            }
            catch
            {
                return false;
            }
        }

        public static async Task<ScriptResult> ExecutePSScript(PSScript method, string para)
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

        private void CaptureOutput(PSDataCollection<PSObject> source, StringBuilder target)
        {
            foreach (var item in source.ReadAll())
            {
                target.AppendLine(item.ToString());
            }
        }

        private void CaptureErrors(PSDataCollection<ErrorRecord> source, StringBuilder target)
        {
            foreach (var error in source.ReadAll())
            {
                target.AppendLine(error.ToString());
            }
        }

        private async Task WaitForCompletion(IAsyncResult asyncResult, TimeSpan? timeout)
        {
            var completionTask = Task.Factory.FromAsync(
                asyncResult,
                result => { /* EndInvoke will auto handle */ });

            if (timeout.HasValue)
            {
                await Task.WhenAny(completionTask, Task.Delay(timeout.Value));
                if (!completionTask.IsCompleted)
                {
                    throw new TimeoutException();
                }
            }

            await completionTask;
        }

        public void Dispose()
        {
            if (_runspacePool != null)
            {
                _runspacePool.Close();
                _runspacePool.Dispose();
            }
        }
    }
}
