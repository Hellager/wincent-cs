using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Threading.Tasks;

namespace Wincent
{
    public enum QuickAccessItemType { File, Directory }

    /// <summary>
    /// Static class providing quick access item management functionality
    /// </summary>
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

        // Internal interface for dependency injection
        internal interface IFileSystemService
        {
            bool FileExists(string path);
            bool DirectoryExists(string path);
            string GetFileExtension(string path);
            Task<bool> ExecuteScriptAsync(PSScript scriptType, string parameter);
            void AddFileToRecentDocs(string path);
            void EmptyRecentFiles();
            void EmptyFrequentFolders();
            void ValidatePathSecurity(string path, QuickAccessItemType itemType);
        }

        // Default implementation
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

            public string GetFileExtension(string path)
            {
                return Path.GetExtension(path);
            }

            public async Task<bool> ExecuteScriptAsync(PSScript scriptType, string parameter)
            {
                var result = await ScriptExecutor.ExecutePSScript(scriptType, parameter);
                return result.ExitCode == 0;
            }

            public void AddFileToRecentDocs(string path)
            {
                QuickAccessManager.AddFileToRecentDocs(path);
            }

            public void EmptyRecentFiles()
            {
                QuickAccessManager.EmptyRecentFiles();
            }

            public void EmptyFrequentFolders()
            {
                QuickAccessManager.EmptyFrequentFolders();
            }

            public void ValidatePathSecurity(string path, QuickAccessItemType itemType)
            {
                QuickAccessManager.ValidatePathSecurity(path, itemType);
            }
        }

        // Service instances
        private static IFileSystemService _fileSystemService = new DefaultFileSystemService();

        // Service replacement method for testing
        internal static void SetFileSystemService(IFileSystemService service)
        {
            _fileSystemService = service ?? new DefaultFileSystemService();
        }

        // Service reset method for testing
        internal static void ResetFileSystemService()
        {
            _fileSystemService = new DefaultFileSystemService();
        }

        /// <summary>
        /// Pins a folder to frequent folders
        /// </summary>
        /// <param name="path">Folder path to add</param>
        /// <returns>Whether the operation succeeded</returns>
        internal static async Task<bool> PinToFrequentFolderAsync(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return false;

            if (_fileSystemService.DirectoryExists(path))
            {
                return await _fileSystemService.ExecuteScriptAsync(PSScript.PinToFrequentFolder, path);
            }

            return false;
        }

        /// <summary>
        /// Unpins a folder from frequent folders
        /// </summary>
        /// <param name="path">Folder path to remove</param>
        /// <returns>Whether the operation succeeded</returns>
        internal static async Task<bool> UnpinFromFrequentFolderAsync(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return false;

            return await _fileSystemService.ExecuteScriptAsync(PSScript.UnpinFromFrequentFolder, path);
        }

        /// <summary>
        /// Adds a file to recent files list
        /// </summary>
        /// <param name="path">File path to add</param>
        /// <returns>Whether the operation succeeded</returns>
        internal static async Task<bool> AddToRecentFilesAsync(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !_fileSystemService.FileExists(path))
                return false;

            string ext = _fileSystemService.GetFileExtension(path);
            if (string.IsNullOrEmpty(ext))
                return false;

            try
            {
                await Task.Run(() => _fileSystemService.AddFileToRecentDocs(path));
                return true;
            }
            catch
            {
                return false;
            }
        }

        internal static void AddFileToRecentDocs(string filePath)
        {
            ValidatePath(filePath, QuickAccessItemType.File);

            IntPtr pathPtr = IntPtr.Zero;
            try
            {
                int hr = CoInitializeEx(IntPtr.Zero, COINIT_APARTMENTTHREADED);
                if (hr < 0) throw new Win32Exception(hr, "COM initialization failed");

                pathPtr = Marshal.StringToHGlobalUni(filePath);
                SHAddToRecentDocs(SHARD_PATHW, pathPtr);
            }
            finally
            {
                if (pathPtr != IntPtr.Zero) Marshal.FreeHGlobal(pathPtr);
                CoUninitialize();
            }
        }

        /// <summary>
        /// Removes a file from recent files list
        /// </summary>
        /// <param name="path">File path to remove</param>
        /// <returns>Whether the operation succeeded</returns>
        internal static async Task<bool> RemoveFromRecentFilesAsync(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !_fileSystemService.FileExists(path))
                return false;

            string ext = _fileSystemService.GetFileExtension(path);
            if (string.IsNullOrEmpty(ext))
                return false;

            return await _fileSystemService.ExecuteScriptAsync(PSScript.RemoveRecentFile, path);
        }

        /// <summary>
        /// Clears recent files list
        /// </summary>
        /// <returns>Whether the operation succeeded</returns>
        internal static async Task<bool> ClearRecentFilesAsync()
        {
            try
            {
                await Task.Run(() => _fileSystemService.EmptyRecentFiles());
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Clears frequent folders list
        /// </summary>
        /// <returns>Whether the operation succeeded</returns>
        internal static async Task<bool> ClearFrequentFoldersAsync()
        {
            try
            {
                await Task.Run(() => _fileSystemService.EmptyFrequentFolders());
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Retrieves all items in quick access
        /// </summary>
        /// <returns>List of quick access items</returns>
        public static async Task<List<string>> GetQuickAccessItemsAsync()
        {
            return await QuickAccessQuery.GetAllItemsAsync();
        }

        /// <summary>
        /// Retrieves recent files list
        /// </summary>
        /// <returns>Recent files list</returns>
        public static async Task<List<string>> GetRecentFilesAsync()
        {
            return await QuickAccessQuery.GetRecentFilesAsync();
        }

        /// <summary>
        /// Retrieves frequent folders list
        /// </summary>
        /// <returns>Frequent folders list</returns>
        public static async Task<List<string>> GetFrequentFoldersAsync()
        {
            return await QuickAccessQuery.GetFrequentFoldersAsync();
        }

        internal static void EmptyRecentFiles()
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

        internal static void EmptyFrequentFolders()
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

        internal static void EmptyQuickAccess()
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

        /// <summary>
        /// Adds an item to quick access
        /// </summary>
        /// <param name="path">Item path to add</param>
        /// <param name="itemType">Item type (File/Directory)</param>
        /// <returns>Whether the operation succeeded</returns>
        public static async Task<bool> AddItemAsync(string path, QuickAccessItemType itemType)
        {
            try
            {
                _fileSystemService.ValidatePathSecurity(path, itemType);

                if (itemType == QuickAccessItemType.File)
                {
                    return await AddToRecentFilesAsync(path);
                }
                else
                {
                    return await PinToFrequentFolderAsync(path);
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Removes an item from quick access
        /// </summary>
        /// <param name="path">Item path to remove</param>
        /// <param name="itemType">Item type (File/Directory)</param>
        /// <returns>Whether the operation succeeded</returns>
        public static async Task<bool> RemoveItemAsync(string path, QuickAccessItemType itemType)
        {
            try
            {
                _fileSystemService.ValidatePathSecurity(path, itemType);

                if (itemType == QuickAccessItemType.File)
                {
                    return await RemoveFromRecentFilesAsync(path);
                }
                else
                {
                    return await UnpinFromFrequentFolderAsync(path);
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Clears quick access items of specified type
        /// </summary>
        /// <param name="itemType">Item type to clear (File/Directory)</param>
        /// <returns>Whether the operation succeeded</returns>
        public static async Task<bool> EmptyItemsAsync(QuickAccessItemType itemType)
        {
            try
            {
                if (itemType == QuickAccessItemType.File)
                {
                    return await ClearRecentFilesAsync();
                }
                else
                {
                    return await ClearFrequentFoldersAsync();
                }
            }
            catch (Exception)
            {
                return false;
            }
        }
    }

    public class QuickAccessOperationException : Exception
    {
        public QuickAccessOperationException(string message) : base(message) { }
        public QuickAccessOperationException(string message, Exception inner) : base(message, inner) { }
    }
}
