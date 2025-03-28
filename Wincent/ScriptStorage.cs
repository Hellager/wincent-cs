using System;
using System.IO;
using System.Linq;
using System.Text;

namespace Wincent
{
    public static class ScriptStorage
    {
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

        // Script strategy factory
        private static readonly IPSScriptStrategyFactory _strategyFactory = new DefaultPSScriptStrategyFactory();

        static ScriptStorage()
        {
            // Ensure all directories exist
            Directory.CreateDirectory(ScriptRoot);
            Directory.CreateDirectory(StaticScriptDir);
            Directory.CreateDirectory(DynamicScriptDir);
        }

        /// <summary>
        /// Gets the full script path, creates the script if not exist
        /// </summary>
        /// <param name="script">Script type</param>
        /// <returns>Full path to script file</returns>
        public static string GetScriptPath(PSScript script)
        {
            var fileName = $"{script}.generated.ps1";

            // Select directory based on script type
            string directory = IsParameterizedScript(script) ? DynamicScriptDir : StaticScriptDir;
            string scriptPath = Path.Combine(directory, fileName);

            // Create script if not exists and non-parameterized
            if (!File.Exists(scriptPath) && !IsParameterizedScript(script))
            {
                CreateScriptFile(scriptPath, script, null);
            }

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
                PSScript.UnpinFromFrequentFolder
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
            if (!IsParameterizedScript(script))
                throw new ArgumentException($"Script {script} is not a parameterized script");

            if (string.IsNullOrEmpty(parameter))
                throw new ArgumentException("Parameter cannot be null or empty for parameterized scripts");

            // Create unique filename using parameter hash
            string paramHash = GetParameterHash(parameter);
            string fileName = $"{script}_{paramHash}.generated.ps1";
            string scriptPath = Path.Combine(DynamicScriptDir, fileName);

            // Create script if not exists
            if (!File.Exists(scriptPath))
            {
                CreateScriptFile(scriptPath, script, parameter);
            }

            return scriptPath;
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
                // Retrieve script strategy
                var strategy = _strategyFactory.GetStrategy(script);

                // Generate script content
                string scriptContent = strategy.GenerateScript(parameter);

                // Convert to UTF8 with BOM
                byte[] scriptBytes = Encoding.UTF8.GetBytes(scriptContent);
                byte[] contentWithBom = ScriptExecutor.AddUtf8Bom(scriptBytes);

                // Write to file
                File.WriteAllBytes(scriptPath, contentWithBom);
            }
            catch (Exception ex)
            {
                throw new IOException($"Failed to create script file: {scriptPath}", ex);
            }
        }

        /// <summary>
        /// Cleans up expired dynamic scripts
        /// </summary>
        /// <param name="maxAgeHours">Maximum retention time in hours</param>
        public static void CleanupDynamicScripts(int maxAgeHours = 24)
        {
            try
            {
                var directory = new DirectoryInfo(DynamicScriptDir);
                var cutoffTime = DateTime.Now.AddHours(-maxAgeHours);

                foreach (var file in directory.GetFiles("*.generated.ps1"))
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
        /// Computes parameter hash for unique filenames
        /// </summary>
        private static string GetParameterHash(string parameter)
        {
            // Use simple hash to avoid long filenames
            using (var md5 = System.Security.Cryptography.MD5.Create())
            {
                byte[] inputBytes = Encoding.UTF8.GetBytes(parameter);
                byte[] hashBytes = md5.ComputeHash(inputBytes);

                // Convert to short hex string
                return BitConverter.ToString(hashBytes)
                    .Replace("-", "")
                    .Substring(0, 8);
            }
        }
    }
}
