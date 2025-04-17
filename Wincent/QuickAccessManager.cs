//using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.IO;
//using System.Linq;
//using System.Runtime.InteropServices;
//using System.Security;
//using System.Threading.Tasks;

//namespace Wincent
//{
//    public enum QuickAccessItemType { File, Directory }

//    /// <summary>
//    /// Static class providing quick access item management functionality
//    /// </summary>
//    public class QuickAccessManager
//    {
//        private readonly IFileSystemService _fileSystemService;
//        private bool _isInitialized;
//        private readonly object _initLock = new object();
//        private bool _queryFeasible;
//        private bool _pinUnpinFeasible;

//        public QuickAccessManager() : this(new DefaultFileSystemService())
//        {
//        }

//        internal QuickAccessManager(IFileSystemService fileSystemService)
//        {
//            _fileSystemService = fileSystemService;
//        }

//        private async Task EnsureInitializedAsync()
//        {
//            if (_isInitialized) return;

//            lock (_initLock)
//            {
//                if (_isInitialized) return;

//                // Check script execution policy
//                if (!_fileSystemService.CheckScriptFeasible())
//                {
//                    try
//                    {
//                        _fileSystemService.FixExecutionPolicy();
//                    }
//                    catch
//                    {
//                        // Ignore fix failures
//                    }
//                }

//                _isInitialized = true;
//            }

//            // Check specific functionalities
//            _queryFeasible = await _fileSystemService.CheckQueryFeasibleAsync();
//            _pinUnpinFeasible = await _fileSystemService.CheckPinUnpinFeasibleAsync();
//        }

//        /// <summary>
//        /// Checks if an item exists in quick access
//        /// </summary>
//        public async Task<bool> CheckItemAsync(string path, QuickAccessItemType itemType)
//        {
//            try
//            {
//                await EnsureInitializedAsync();
//                if (!_queryFeasible)
//                    throw new QuickAccessFeasibilityException("Quick access query functionality is not available");

//                _fileSystemService.ValidatePathSecurity(path, itemType);

//                var items = await GetQuickAccessItemsAsync();
//                return items.Contains(path, StringComparer.OrdinalIgnoreCase);
//            }
//            catch (QuickAccessFeasibilityException)
//            {
//                throw;
//            }
//            catch (Exception ex)
//            {
//                // 记录异常但不抛出
//                System.Diagnostics.Debug.WriteLine($"Error checking quick access item: {ex.Message}");
//                return false;
//            }
//        }

//        /// <summary>
//        /// Pins a folder to frequent folders
//        /// </summary>
//        /// <param name="path">Folder path to add</param>
//        /// <returns>Whether the operation succeeded</returns>
//        internal async Task<bool> PinToFrequentFolderAsync(string path)
//        {
//            if (string.IsNullOrWhiteSpace(path))
//                return false;

//            if (_fileSystemService.DirectoryExists(path))
//            {
//                return await _fileSystemService.ExecuteScriptAsync(PSScript.PinToFrequentFolder, path);
//            }

//            return false;
//        }

//        /// <summary>
//        /// Unpins a folder from frequent folders
//        /// </summary>
//        /// <param name="path">Folder path to remove</param>
//        /// <returns>Whether the operation succeeded</returns>
//        internal async Task<bool> UnpinFromFrequentFolderAsync(string path)
//        {
//            if (string.IsNullOrWhiteSpace(path))
//                return false;

//            return await _fileSystemService.ExecuteScriptAsync(PSScript.UnpinFromFrequentFolder, path);
//        }

//        /// <summary>
//        /// Adds a file to recent files list
//        /// </summary>
//        /// <param name="path">File path to add</param>
//        /// <returns>Whether the operation succeeded</returns>
//        internal async Task<bool> AddToRecentFilesAsync(string path)
//        {
//            if (string.IsNullOrWhiteSpace(path) || !_fileSystemService.FileExists(path))
//                return false;

//            string ext = _fileSystemService.GetFileExtension(path);
//            if (string.IsNullOrEmpty(ext))
//                return false;

//            try
//            {
//                await Task.Run(() => _fileSystemService.AddFileToRecentDocs(path));
//                return true;
//            }
//            catch
//            {
//                return false;
//            }
//        }

//        internal void AddFileToRecentDocs(string filePath)
//        {
//            ValidatePath(filePath, QuickAccessItemType.File);

//            IntPtr pathPtr = IntPtr.Zero;
//            try
//            {
//                int hr = NativeMethods.CoInitializeEx(IntPtr.Zero, NativeMethods.COINIT_APARTMENTTHREADED);
//                if (hr < 0) throw new Win32Exception(hr, "COM initialization failed");

//                pathPtr = Marshal.StringToHGlobalUni(filePath);
//                NativeMethods.SHAddToRecentDocs(NativeMethods.SHARD_PATHW, pathPtr);
//            }
//            finally
//            {
//                if (pathPtr != IntPtr.Zero) Marshal.FreeHGlobal(pathPtr);
//                NativeMethods.CoUninitialize();
//            }
//        }

//        /// <summary>
//        /// Removes a file from recent files list
//        /// </summary>
//        /// <param name="path">File path to remove</param>
//        /// <returns>Whether the operation succeeded</returns>
//        internal async Task<bool> RemoveFromRecentFilesAsync(string path)
//        {
//            if (string.IsNullOrWhiteSpace(path) || !_fileSystemService.FileExists(path))
//                return false;

//            string ext = _fileSystemService.GetFileExtension(path);
//            if (string.IsNullOrEmpty(ext))
//                return false;

//            return await _fileSystemService.ExecuteScriptAsync(PSScript.RemoveRecentFile, path);
//        }

//        /// <summary>
//        /// Clears recent files list
//        /// </summary>
//        /// <returns>Whether the operation succeeded</returns>
//        internal async Task<bool> ClearRecentFilesAsync()
//        {
//            try
//            {
//                await Task.Run(() => _fileSystemService.EmptyRecentFiles());
//                return true;
//            }
//            catch
//            {
//                return false;
//            }
//        }

//        /// <summary>
//        /// Clears frequent folders list
//        /// </summary>
//        /// <returns>Whether the operation succeeded</returns>
//        internal async Task<bool> ClearFrequentFoldersAsync()
//        {
//            try
//            {
//                await Task.Run(() => _fileSystemService.EmptyFrequentFolders());
//                return true;
//            }
//            catch
//            {
//                return false;
//            }
//        }

//        /// <summary>
//        /// Retrieves all items in quick access
//        /// </summary>
//        /// <returns>List of quick access items</returns>
//        public async Task<List<string>> GetQuickAccessItemsAsync()
//        {
//            try
//            {
//                await EnsureInitializedAsync();
//                if (!_queryFeasible)
//                    throw new QuickAccessFeasibilityException("Quick access query functionality is not available");

//                return await _fileSystemService.GetQuickAccessItemsAsync();
//            }
//            catch (QuickAccessFeasibilityException)
//            {
//                throw;
//            }
//            catch (Exception ex)
//            {
//                System.Diagnostics.Debug.WriteLine($"Error getting quick access items: {ex.Message}");
//                return new List<string>();
//            }
//        }

//        /// <summary>
//        /// Retrieves recent files list
//        /// </summary>
//        /// <returns>Recent files list</returns>
//        public async Task<List<string>> GetRecentFilesAsync()
//        {
//            return await QuickAccessQuery.GetRecentFilesAsync();
//        }

//        /// <summary>
//        /// Retrieves frequent folders list
//        /// </summary>
//        /// <returns>Frequent folders list</returns>
//        public async Task<List<string>> GetFrequentFoldersAsync()
//        {
//            return await QuickAccessQuery.GetFrequentFoldersAsync();
//        }

//        internal void EmptyRecentFiles()
//        {
//            try
//            {
//                NativeMethods.CoInitializeEx(IntPtr.Zero, NativeMethods.COINIT_APARTMENTTHREADED);
//                NativeMethods.SHAddToRecentDocs(NativeMethods.SHARD_PATHW, IntPtr.Zero);
//            }
//            finally
//            {
//                NativeMethods.CoUninitialize();
//            }
//        }

//        internal void EmptyFrequentFolders()
//        {
//            var jumplistPath = GetKnownFolderPath(NativeMethods.FOLDERID_Recent);
//            var jumplistFile = Path.Combine(jumplistPath,
//                "AutomaticDestinations",
//                "f01b4d95cf55d32a.automaticDestinations-ms");

//            if (File.Exists(jumplistFile))
//            {
//                try { File.Delete(jumplistFile); }
//                catch { /* Ignore failed deleting */ }
//            }

//            var folders = QuickAccessQuery.GetFrequentFoldersAsync().Result;
//            foreach (var folder in folders)
//            {
//                try
//                {
//                    ScriptExecutor.ExecutePSScript(PSScript.UnpinFromFrequentFolder, folder).Wait();
//                }
//                catch { /* Ignore single failed deleting */ }
//            }
//        }

//        public void EmptyQuickAccess()
//        {
//            EmptyRecentFiles();
//            EmptyFrequentFolders();
//        }

//        public void ValidatePathSecurity(string path, QuickAccessItemType expectedType)
//        {
//            if (NativeMethods.ProtectedPaths.Any(p => path.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0))
//                throw new SecurityException("Protected system path.");

//            ValidatePath(path, expectedType);
//        }

//        private void ValidatePath(string path, QuickAccessItemType expectedType)
//        {
//            if (string.IsNullOrWhiteSpace(path))
//                throw new ArgumentException("The path cannot be empty.");

//            bool pathExists;
//            if (expectedType == QuickAccessItemType.File)
//                pathExists = File.Exists(path);
//            else
//                pathExists = Directory.Exists(path);

//            if (!pathExists)
//                throw new FileNotFoundException($"The path does not exist: {path}");
//        }

//        private string GetKnownFolderPath(Guid knownFolderId)
//        {
//            IntPtr pPath = IntPtr.Zero;
//            try
//            {
//                int hr = NativeMethods.SHGetKnownFolderPath(knownFolderId, 0, IntPtr.Zero, out pPath);
//                if (hr != 0) throw new Win32Exception(hr);

//                return Marshal.PtrToStringUni(pPath);
//            }
//            finally
//            {
//                if (pPath != IntPtr.Zero)
//                    NativeMethods.CoTaskMemFree(pPath);
//            }
//        }

//        /// <summary>
//        /// Adds an item to quick access
//        /// </summary>
//        /// <param name="path">Item path to add</param>
//        /// <param name="itemType">Item type (File/Directory)</param>
//        /// <returns>Whether the operation succeeded</returns>
//        public async Task<bool> AddItemAsync(string path, QuickAccessItemType itemType)
//        {
//            try
//            {
//                await EnsureInitializedAsync();
//                if (!_pinUnpinFeasible)
//                    throw new QuickAccessFeasibilityException("Quick access pin/unpin functionality is not available");

//                _fileSystemService.ValidatePathSecurity(path, itemType);

//                if (itemType == QuickAccessItemType.File)
//                {
//                    return await AddToRecentFilesAsync(path);
//                }
//                else
//                {
//                    return await PinToFrequentFolderAsync(path);
//                }
//            }
//            catch (QuickAccessFeasibilityException)
//            {
//                throw;
//            }
//        }

//        /// <summary>
//        /// Removes an item from quick access
//        /// </summary>
//        /// <param name="path">Item path to remove</param>
//        /// <param name="itemType">Item type (File/Directory)</param>
//        /// <returns>Whether the operation succeeded</returns>
//        public async Task<bool> RemoveItemAsync(string path, QuickAccessItemType itemType)
//        {
//            try
//            {
//                await EnsureInitializedAsync();
//                if (!_pinUnpinFeasible)
//                    throw new QuickAccessFeasibilityException("Quick access pin/unpin functionality is not available");

//                _fileSystemService.ValidatePathSecurity(path, itemType);

//                if (itemType == QuickAccessItemType.File)
//                {
//                    return await RemoveFromRecentFilesAsync(path);
//                }
//                else
//                {
//                    return await UnpinFromFrequentFolderAsync(path);
//                }
//            }
//            catch (QuickAccessFeasibilityException)
//            {
//                throw;
//            }
//        }

//        /// <summary>
//        /// Clears quick access items of specified type
//        /// </summary>
//        /// <param name="itemType">Item type to clear (File/Directory)</param>
//        /// <returns>Whether the operation succeeded</returns>
//        public async Task<bool> EmptyItemsAsync(QuickAccessItemType itemType)
//        {
//            try
//            {
//                if (itemType == QuickAccessItemType.File)
//                {
//                    return await ClearRecentFilesAsync();
//                }
//                else
//                {
//                    return await ClearFrequentFoldersAsync();
//                }
//            }
//            catch (Exception)
//            {
//                return false;
//            }
//        }

//        private static class NativeMethods
//        {
//            internal const uint SHARD_PATHW = 0x00000003;
//            internal const int COINIT_APARTMENTTHREADED = 0x2;
//            internal static readonly Guid FOLDERID_Recent = new Guid("{AE50C081-EBD2-438A-8655-8A092E34987A}");
//            internal static readonly string[] ProtectedPaths = { "System32", "Program Files" };

//            [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
//            internal static extern void SHAddToRecentDocs(uint uFlags, IntPtr pv);

//            [DllImport("ole32.dll")]
//            internal static extern int CoInitializeEx(IntPtr reserved, uint dwCoInit);

//            [DllImport("ole32.dll")]
//            internal static extern void CoUninitialize();

//            [DllImport("shell32.dll")]
//            internal static extern int SHGetKnownFolderPath(
//                [MarshalAs(UnmanagedType.LPStruct)] Guid rfid,
//                uint dwFlags,
//                IntPtr hToken,
//                out IntPtr ppszPath);

//            [DllImport("ole32.dll")]
//            internal static extern void CoTaskMemFree(IntPtr ptr);
//        }
//    }

//    public class QuickAccessOperationException : Exception
//    {
//        public QuickAccessOperationException(string message) : base(message) { }
//        public QuickAccessOperationException(string message, Exception inner) : base(message, inner) { }
//    }

//    public class QuickAccessFeasibilityException : Exception
//    {
//        public QuickAccessFeasibilityException(string message) : base(message) { }
//        public QuickAccessFeasibilityException(string message, Exception inner) : base(message, inner) { }
//    }

//    public interface IFileSystemService
//    {
//        bool FileExists(string path);
//        bool DirectoryExists(string path);
//        string GetFileExtension(string path);
//        Task<bool> ExecuteScriptAsync(PSScript scriptType, string parameter);
        
//        void AddFileToRecentDocs(string path);
//        void EmptyRecentFiles();
//        void EmptyFrequentFolders();
//        void ValidatePathSecurity(string path, QuickAccessItemType itemType);
        
//        bool CheckScriptFeasible();
//        void FixExecutionPolicy();
//        Task<bool> CheckQueryFeasibleAsync();
//        Task<bool> CheckPinUnpinFeasibleAsync();
//        Task<List<string>> GetQuickAccessItemsAsync();
//        Task<List<string>> GetRecentFilesAsync();
//        Task<List<string>> GetFrequentFoldersAsync();
//    }

//    public class DefaultFileSystemService : IFileSystemService
//    {
//        private readonly QuickAccessManager _manager;

//        public DefaultFileSystemService()
//        {
//            _manager = new QuickAccessManager();
//        }

//        public void AddFileToRecentDocs(string path)
//        {
//            _manager.AddFileToRecentDocs(path);
//        }

//        public void EmptyRecentFiles()
//        {
//            _manager.EmptyRecentFiles();
//        }

//        public void EmptyFrequentFolders()
//        {
//            _manager.EmptyFrequentFolders();
//        }

//        public void ValidatePathSecurity(string path, QuickAccessItemType itemType)
//        {
//            _manager.ValidatePathSecurity(path, itemType);
//        }

//        public bool CheckScriptFeasible()
//        {
//            return ExecutionFeasible.CheckScriptFeasible();
//        }

//        public void FixExecutionPolicy()
//        {
//            ExecutionFeasible.FixExecutionPolicy();
//        }

//        public async Task<bool> CheckQueryFeasibleAsync()
//        {
//            return await ExecutionFeasible.CheckQueryFeasible();
//        }

//        public async Task<bool> CheckPinUnpinFeasibleAsync()
//        {
//            return await ExecutionFeasible.CheckPinUnpinFeasible();
//        }

//        public bool FileExists(string path)
//        {
//            throw new NotImplementedException();
//        }

//        public bool DirectoryExists(string path)
//        {
//            throw new NotImplementedException();
//        }

//        public string GetFileExtension(string path)
//        {
//            throw new NotImplementedException();
//        }

//        public Task<bool> ExecuteScriptAsync(PSScript scriptType, string parameter)
//        {
//            throw new NotImplementedException();
//        }

//        public Task<List<string>> GetQuickAccessItemsAsync()
//        {
//            throw new NotImplementedException();
//        }

//        public Task<List<string>> GetRecentFilesAsync()
//        {
//            throw new NotImplementedException();
//        }

//        public Task<List<string>> GetFrequentFoldersAsync()
//        {
//            throw new NotImplementedException();
//        }
//    }
//}
