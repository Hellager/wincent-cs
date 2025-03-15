using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Threading.Tasks;

namespace Wincent
{
    public enum QuickAccessItemType { File, Directory }

    public static class QuickAccessManager
    {
        private const uint SHARD_PATHW = 0x00000003;
        private const int COINIT_APARTMENTTHREADED = 0x2;

        private static Guid FOLDERID_Recent = new Guid("{AE50C081-EBD2-438A-8655-8A092E34987A}");

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern void SHAddToRecentDocs(uint uFlags, IntPtr pv);

        [DllImport("ole32.dll")]
        private static extern int CoInitializeEx(IntPtr reserved, uint dwCoInit);

        [DllImport("ole32.dll")]
        private static extern void CoUninitialize();

        [DllImport("shell32.dll")]
        private static extern int SHGetKnownFolderPath(
            [MarshalAs(UnmanagedType.LPStruct)] Guid rfid,
            uint dwFlags,
            IntPtr hToken,
            out IntPtr ppszPath);

        [DllImport("ole32.dll")]
        private static extern void CoTaskMemFree(IntPtr ptr);

        private static readonly string[] ProtectedPaths = { "System32", "Program Files" };

        public static async Task AddItemAsync(string path, QuickAccessItemType itemType)
        {
            ValidatePathSecurity(path, itemType);
            await ExecuteAddOperation(path, itemType);
        }

        public static async Task RemoveItemAsync(string path, QuickAccessItemType itemType)
        {
            ValidatePathSecurity(path, itemType);
            await ExecuteRemoveOperation(path, itemType);
        }

        public static void AddFileToRecentDocs(string filePath)
        {
            ValidatePath(filePath, QuickAccessItemType.File);

            IntPtr pathPtr = IntPtr.Zero;
            try
            {
                int hr = CoInitializeEx(IntPtr.Zero, COINIT_APARTMENTTHREADED);
                if (hr < 0) throw new Win32Exception(hr, "COM 初始化失败");

                pathPtr = Marshal.StringToHGlobalUni(filePath);
                SHAddToRecentDocs(SHARD_PATHW, pathPtr);
            }
            finally
            {
                if (pathPtr != IntPtr.Zero) Marshal.FreeHGlobal(pathPtr);
                CoUninitialize();
            }
        }

        public static void EmptyRecentFiles()
        {
            try
            {
                CoInitializeEx(IntPtr.Zero, COINIT_APARTMENTTHREADED);
                SHAddToRecentDocs(SHARD_PATHW, IntPtr.Zero);
            }
            finally
            {
                CoUninitialize();
            }
        }

        public static void EmptyFrequentFolders()
        {
            var jumplistPath = GetKnownFolderPath(FOLDERID_Recent);
            var jumplistFile = Path.Combine(jumplistPath,
                "AutomaticDestinations",
                "f01b4d95cf55d32a.automaticDestinations-ms");

            if (File.Exists(jumplistFile))
            {
                try { File.Delete(jumplistFile); }
                catch { /* Ignore failed deleting */ }
            }

            var folders = QuickAccessQuery.GetFrequentFoldersAsync().Result;
            foreach (var folder in folders)
            {
                try
                {
                    ScriptExecutor.ExecutePSScript(PSScript.UnpinFromFrequentFolder, folder).Wait();
                }
                catch { /* Ignore single failed deleting */ }
            }
        }

        public static void EmptyQuickAccess()
        {
            EmptyRecentFiles();
            EmptyFrequentFolders();
        }

        private static async Task ExecuteAddOperation(string path, QuickAccessItemType itemType)
        {
            if (itemType == QuickAccessItemType.File)
            {
                AddFileToRecentDocs(path);
            }
            else
            {
                var result = await ScriptExecutor.ExecutePSScript(PSScript.PinToFrequentFolder, path);
                if (result.ExitCode != 0)
                    throw new QuickAccessOperationException($"Addition failed: {result.Error}.");
            }
        }

        private static async Task ExecuteRemoveOperation(string path, QuickAccessItemType itemType)
        {
            var script = (itemType == QuickAccessItemType.File) ?
                PSScript.RemoveRecentFile :
                PSScript.UnpinFromFrequentFolder;

            var result = await ScriptExecutor.ExecutePSScript(script, path);
            if (result.ExitCode != 0)
                throw new QuickAccessOperationException($"Removal failed: {result.Error}.");
        }

        public static void ValidatePathSecurity(string path, QuickAccessItemType expectedType)
        {
            if (ProtectedPaths.Any(p => path.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0))
                throw new SecurityException("Protected system path.");

            ValidatePath(path, expectedType);
        }

        private static void ValidatePath(string path, QuickAccessItemType expectedType)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("The path cannot be empty.");

            bool pathExists;
            if (expectedType == QuickAccessItemType.File)
                pathExists = File.Exists(path);
            else
                pathExists = Directory.Exists(path);

            if (!pathExists)
                throw new FileNotFoundException($"The path does not exist: {path}");
        }

        private static string GetKnownFolderPath(Guid knownFolderId)
        {
            IntPtr pPath = IntPtr.Zero;
            try
            {
                int hr = SHGetKnownFolderPath(knownFolderId, 0, IntPtr.Zero, out pPath);
                if (hr != 0) throw new Win32Exception(hr);

                return Marshal.PtrToStringUni(pPath);
            }
            finally
            {
                if (pPath != IntPtr.Zero)
                    CoTaskMemFree(pPath);
            }
        }
    }

    public class QuickAccessOperationException : Exception
    {
        public QuickAccessOperationException(string message) : base(message) { }
        public QuickAccessOperationException(string message, Exception inner) : base(message, inner) { }
    }
}
