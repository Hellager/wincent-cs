using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Wincent
{
    public enum QuickAccessItemType { File, Directory }

    public static class QuickAccessManager
    {
        public const uint SHARD_PIDL = 0x00000001;
        public const uint SHARD_PATHW = 0x00000003;
        internal const int COINIT_APARTMENTTHREADED = 0x2;

        // Known Folder ID definition
        private static Guid FOLDERID_Recent = new Guid("{AE50C081-EBD2-438A-8655-8A092E34987A}");

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        internal static extern void SHAddToRecentDocs(uint uFlags, IntPtr pv);

        [DllImport("ole32.dll")]
        internal static extern int CoInitializeEx(IntPtr reserved, uint dwCoInit);

        [DllImport("ole32.dll")]
        internal static extern void CoUninitialize();

        [DllImport("shell32.dll")]
        private static extern int SHGetKnownFolderPath(
            [MarshalAs(UnmanagedType.LPStruct)] Guid rfid,
            uint dwFlags,
            IntPtr hToken,
            out IntPtr ppszPath);

        [DllImport("ole32.dll")]
        private static extern void CoTaskMemFree(IntPtr ptr);

        private static readonly string[] ProtectedPaths = { "System32", "Program Files" };

        /// <summary>
        /// Safely add an item to Quick Access (supports files and directories)
        /// </summary>
        public static async Task AddItemAsync(string path, QuickAccessItemType itemType)
        {
            ValidatePathSecurity(path, itemType);
            await ExecuteAddOperation(path, itemType);
        }

        /// <summary>
        /// Safely remove an item from Quick Access
        /// </summary>
        public static async Task RemoveItemAsync(string path, QuickAccessItemType itemType)
        {
            ValidatePathSecurity(path, itemType);
            await ExecuteRemoveOperation(path, itemType);
        }

        /// <summary>
        /// Add file to Recent Documents using Windows API
        /// </summary>
        public static void AddFileToRecentDocs(string filePath)
        {
            ValidatePath(filePath, QuickAccessItemType.File);

            // Convert path to UTF-16 format (required by Windows API)
            IntPtr pathPtr = IntPtr.Zero;
            try
            {
                // Initialize COM library
                int hr = CoInitializeEx(IntPtr.Zero, COINIT_APARTMENTTHREADED);
                if (hr < 0)
                {
                    throw new Win32Exception(hr, "COM initialization failed");
                }

                // Allocate unmanaged memory and copy path
                pathPtr = Marshal.StringToHGlobalUni(filePath);

                // Call Windows API to add file to recent access
                SHAddToRecentDocs(SHARD_PATHW, pathPtr);
            }
            finally
            {
                // Clean up resources
                if (pathPtr != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(pathPtr);
                }
                CoUninitialize();
            }
        }

        /// <summary>
        /// Clear Recent Files list
        /// </summary>
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

        /// <summary>
        /// Clear Frequent Folders list
        /// </summary>
        public static void EmptyFrequentFolders()
        {
            // Delete jump list file
            var jumplistPath = GetKnownFolderPath(FOLDERID_Recent);
            var jumplistFile = Path.Combine(jumplistPath,
                "AutomaticDestinations",
                "f01b4d95cf55d32a.automaticDestinations-ms");

            if (File.Exists(jumplistFile))
            {
                File.Delete(jumplistFile);
            }

            // 移除所有已固定的文件夹
            var folders = QuickAccessQuery.GetFrequentFoldersAsync().Result;
            foreach (var folder in folders)
            {
                ScriptExecutor.ExecutePSScript(PSScript.UnpinFromFrequentFolder, folder).Wait();
            }
        }

        /// <summary>
        /// Clear entire Quick Access
        /// </summary>
        public static void EmptyQuickAccess()
        {
            EmptyRecentFiles();
            EmptyFrequentFolders();
        }

        private static async Task ExecuteAddOperation(string path, QuickAccessItemType itemType)
        {
            switch (itemType)
            {
                case QuickAccessItemType.File:
                    AddFileToRecentDocs(path);
                    break;
                case QuickAccessItemType.Directory:
                    var result = await ScriptExecutor.ExecutePSScript(PSScript.PinToFrequentFolder, path);
                    if (result.ExitCode != 0)
                        throw new QuickAccessOperationException($"Add operation failed: {result.Error}");
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(itemType));
            }
        }

        private static async Task ExecuteRemoveOperation(string path, QuickAccessItemType itemType)
        {
            var script = itemType switch
            {
                QuickAccessItemType.File => PSScript.RemoveRecentFile,
                QuickAccessItemType.Directory => PSScript.UnpinFromFrequentFolder,
                _ => throw new ArgumentOutOfRangeException(nameof(itemType))
            };

            var result = await ScriptExecutor.ExecutePSScript(script, path);
            if (result.ExitCode != 0)
                throw new QuickAccessOperationException($"Remove operation failed: {result.Error}");
        }

        public static void ValidatePathSecurity(string path, QuickAccessItemType expectedType)
        {
            if (ProtectedPaths.Any(p => path.Contains(p, StringComparison.OrdinalIgnoreCase)))
                throw new SecurityException("Protected system path");

            ValidatePath(path, expectedType);
        }

        private static void ValidatePath(string path, QuickAccessItemType expectedType)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Path cannot be empty");

            var pathExists = expectedType switch
            {
                QuickAccessItemType.File => File.Exists(path),
                QuickAccessItemType.Directory => Directory.Exists(path),
                _ => false
            };

            if (!pathExists)
                throw new FileNotFoundException($"Path does not exist: {path}");
        }

        private static string GetKnownFolderPath(Guid knownFolderId)
        {
            IntPtr pPath = IntPtr.Zero;
            try
            {
                int hr = SHGetKnownFolderPath(knownFolderId, 0, IntPtr.Zero, out pPath);
                if (hr != 0)
                    throw new Win32Exception(hr);

                string? path = Marshal.PtrToStringUni(pPath);
                if (path == null)
                    throw new InvalidOperationException("Failed to get folder path");

                return path;
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
