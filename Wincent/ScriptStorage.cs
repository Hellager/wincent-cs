using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace Wincent
{
    internal static class ScriptStorage
    {
        /// <summary>
        /// Current script version from assembly version (major.minor.build)
        /// </summary>
        private static readonly string CurrentVersion = GetCurrentVersion();

        /// <summary>
        /// Script root directory located under Windows Temp directory in 'Wincent' folder
        /// </summary>
        public static readonly string ScriptRoot = Path.Combine(
            Path.GetTempPath(),
            "Wincent");

        /// <summary>
        /// Static script directory (parameterless scripts)
        /// </summary>
        public static readonly string StaticScriptDir = Path.Combine(ScriptRoot, "static");

        /// <summary>
        /// Dynamic script directory (parameterized scripts)
        /// </summary>
        public static readonly string DynamicScriptDir = Path.Combine(ScriptRoot, "dynamic");

        private static readonly IPSScriptStrategyFactory _strategyFactory = new DefaultPSScriptStrategyFactory();
        private static readonly object InitializationLock = new object();
        private static readonly ConcurrentDictionary<string, object> ScriptCreationLocks =
            new ConcurrentDictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        private static int _initialized;

        /// <summary>
        /// Adds UTF-8 BOM to the beginning of a byte array if not already present
        /// </summary>
        /// <param name="content">Content bytes</param>
        /// <returns>Content with UTF-8 BOM</returns>
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

        /// <summary>
        /// Gets the current version from assembly version (major.minor.build)
        /// </summary>
        private static string GetCurrentVersion()
        {
            try
            {
                var version = Assembly.GetExecutingAssembly().GetName().Version;
                return $"v{version.Major}.{version.Minor}.{version.Build}";
            }
            catch
            {
                return "v0.1.0";
            }
        }

        /// <summary>
        /// Gets the full script path, creates the script if not exist
        /// </summary>
        /// <param name="script">Script type</param>
        /// <returns>Full path to script file</returns>
        public static string GetScriptPath(PSScript script)
        {
            EnsureInitialized();

            var fileName = $"{script}_{CurrentVersion}.ps1";
            bool isParameterized = IsParameterizedScript(script);

            string directory = isParameterized ? DynamicScriptDir : StaticScriptDir;
            string scriptPath = Path.Combine(directory, fileName);

            if (!isParameterized)
                EnsureScriptFile(scriptPath, script, null);

            return scriptPath;
        }

        /// <summary>
        /// Checks if script requires parameters
        /// </summary>
        /// <param name="script">Script type</param>
        /// <returns>True if script requires parameters</returns>
        public static bool IsParameterizedScript(PSScript script)
        {
            var parameterized = new[]
            {
                PSScript.RemoveRecentFile,
                PSScript.PinToFrequentFolder,
                PSScript.UnpinFromFrequentFolder,
                PSScript.AddRecentFile,
            };
            return parameterized.Contains(script);
        }

        /// <summary>
        /// Gets unique path for dynamic script (parameter-based), creates if not exist
        /// </summary>
        /// <param name="script">Script type</param>
        /// <param name="parameter">Script parameter</param>
        /// <returns>Unique path for dynamic script</returns>
        public static string GetDynamicScriptPath(PSScript script, string parameter)
        {
            EnsureInitialized();

            if (!IsParameterizedScript(script))
                throw new ArgumentException($"Script {script} is not a parameterized script");

            if (string.IsNullOrEmpty(parameter))
                throw new ArgumentException("Parameter cannot be null or empty for parameterized scripts");

            string paramHash = GetParameterHash(parameter);
            string fileName = $"{script}_{CurrentVersion}_{paramHash}.ps1";
            string scriptPath = Path.Combine(DynamicScriptDir, fileName);

            EnsureScriptFile(scriptPath, script, parameter);

            return scriptPath;
        }

        /// <summary>
        /// Cleans up expired dynamic scripts and scripts with different versions
        /// </summary>
        /// <param name="maxAgeHours">Maximum retention time in hours</param>
        public static void CleanupDynamicScripts(int maxAgeHours = 24)
        {
            EnsureInitialized(skipDynamicCleanup: true);
            CleanupDynamicScriptsCore(maxAgeHours);
        }

        internal static void CleanupStaticScripts()
        {
            EnsureInitialized(skipStaticCleanup: true);
            CleanupStaticScriptsByVersion();
        }

        private static void EnsureInitialized(bool skipStaticCleanup = false, bool skipDynamicCleanup = false)
        {
            if (Volatile.Read(ref _initialized) != 0)
                return;

            lock (InitializationLock)
            {
                if (_initialized != 0)
                    return;

                Directory.CreateDirectory(ScriptRoot);
                Directory.CreateDirectory(StaticScriptDir);
                Directory.CreateDirectory(DynamicScriptDir);

                if (!skipStaticCleanup)
                    CleanupStaticScriptsByVersion();
                if (!skipDynamicCleanup)
                    CleanupDynamicScriptsCore();

                Volatile.Write(ref _initialized, 1);
            }
        }

        private static void EnsureScriptFile(string scriptPath, PSScript script, string parameter)
        {
            if (File.Exists(scriptPath))
                return;

            object creationLock = ScriptCreationLocks.GetOrAdd(scriptPath, _ => new object());

            try
            {
                lock (creationLock)
                {
                    if (!File.Exists(scriptPath))
                        CreateScriptFile(scriptPath, script, parameter);
                }
            }
            finally
            {
                ScriptCreationLocks.TryRemove(scriptPath, out _);
            }
        }

        /// <summary>
        /// Creates script file
        /// </summary>
        /// <param name="scriptPath">Script file path</param>
        /// <param name="script">Script type</param>
        /// <param name="parameter">Script parameter (optional)</param>
        private static void CreateScriptFile(string scriptPath, PSScript script, string parameter)
        {
            try
            {
                var strategy = _strategyFactory.GetStrategy(script);
                string scriptContent = strategy.GenerateScript(parameter);

                byte[] scriptBytes = Encoding.UTF8.GetBytes(scriptContent);
                byte[] contentWithBom = AddUtf8Bom(scriptBytes);

                File.WriteAllBytes(scriptPath, contentWithBom);
            }
            catch (Exception ex)
            {
                throw new IOException($"Failed to create script file: {scriptPath}", ex);
            }
        }

        private static void CleanupDynamicScriptsCore(int maxAgeHours = 24)
        {
            try
            {
                CleanupByAge(maxAgeHours);
                CleanupDynamicScriptsByVersion();
            }
            catch
            {
                // Ignore cleanup errors
            }
        }

        /// <summary>
        /// Cleans up scripts based on last modified time
        /// </summary>
        /// <param name="maxAgeHours">Maximum retention time in hours</param>
        private static void CleanupByAge(int maxAgeHours)
        {
            try
            {
                var directory = new DirectoryInfo(DynamicScriptDir);
                var cutoffTime = DateTime.Now.AddHours(-maxAgeHours);

                foreach (var file in directory.GetFiles("*.ps1"))
                {
                    if (file.LastWriteTime < cutoffTime)
                    {
                        try
                        {
                            file.Delete();
                        }
                        catch
                        {
                            // Ignore deletion failures
                        }
                    }
                }
            }
            catch
            {
                // Ignore cleanup errors
            }
        }

        /// <summary>
        /// Cleans up dynamic scripts with versions that don't match current version
        /// </summary>
        private static void CleanupDynamicScriptsByVersion()
        {
            try
            {
                var directory = new DirectoryInfo(DynamicScriptDir);
                var versionPattern = new Regex(@"^([A-Za-z]+)_v(\d+\.\d+\.\d+)_([0-9A-F]{8})\.ps1$");
                foreach (var file in directory.GetFiles("*.ps1"))
                {
                    var match = versionPattern.Match(file.Name);
                    if (!match.Success)
                        continue;

                    string fileVersion = match.Groups[2].Value;
                    if (fileVersion == CurrentVersion.Substring(1))
                        continue;

                    try
                    {
                        file.Delete();
                    }
                    catch
                    {
                        // Ignore deletion failures
                    }
                }
            }
            catch
            {
                // Ignore cleanup errors
            }
        }

        /// <summary>
        /// Cleans up static scripts with versions that don't match current version
        /// </summary>
        private static void CleanupStaticScriptsByVersion()
        {
            try
            {
                var directory = new DirectoryInfo(StaticScriptDir);
                var versionPattern = new Regex(@"^([A-Za-z]+)_v(\d+\.\d+\.\d+)\.ps1$");
                foreach (var file in directory.GetFiles("*.ps1"))
                {
                    var match = versionPattern.Match(file.Name);
                    if (!match.Success)
                        continue;

                    string scriptName = match.Groups[1].Value;
                    string fileVersion = match.Groups[2].Value;
                    if (fileVersion == CurrentVersion.Substring(1))
                        continue;

                    try
                    {
                        file.Delete();
                        if (Enum.TryParse<PSScript>(scriptName, out var scriptType) &&
                            !IsParameterizedScript(scriptType))
                        {
                            var newFilePath = Path.Combine(StaticScriptDir, $"{scriptType}_{CurrentVersion}.ps1");
                            EnsureScriptFile(newFilePath, scriptType, null);
                        }
                    }
                    catch
                    {
                        // Ignore deletion failures
                    }
                }
            }
            catch
            {
                // Ignore cleanup errors
            }
        }

        /// <summary>
        /// Computes parameter hash for unique filenames
        /// </summary>
        private static string GetParameterHash(string parameter)
        {
            using (var sha256 = System.Security.Cryptography.SHA256.Create())
            {
                byte[] inputBytes = Encoding.UTF8.GetBytes(parameter);
                byte[] hashBytes = sha256.ComputeHash(inputBytes);

                return BitConverter.ToString(hashBytes)
                    .Replace("-", "")
                    .Substring(0, 8);
            }
        }
    }
}
