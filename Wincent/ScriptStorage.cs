using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Wincent
{
    public static class ScriptStorage
    {
        public static readonly string ScriptRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "Wincent\\Scripts");

        static ScriptStorage()
        {
            Directory.CreateDirectory(ScriptRoot);
        }

        public static string GetScriptPath(PSScript script)
        {
            var fileName = $"{script}.generated.ps1";
            return Path.Combine(ScriptRoot, fileName);
        }

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
    }

}
