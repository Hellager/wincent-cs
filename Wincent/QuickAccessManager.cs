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
    /// <summary>
    /// Quick Access item types
    /// </summary>
    public enum QuickAccess
    {
        /// <summary>
        /// All items
        /// </summary>
        All,

        /// <summary>
        /// Recently used files
        /// </summary>
        RecentFiles,

        /// <summary>
        /// Frequently used folders
        /// </summary>
        FrequentFolders
    }

    /// <summary>
    /// Quick Access item types
    /// </summary>
    public enum QuickAccessItemType
    {
        /// <summary>
        /// File
        /// </summary>
        File,

        /// <summary>
        /// Directory
        /// </summary>
        Directory
    }

    /// <summary>
    /// Path types
    /// </summary>
    public enum PathType
    {
        /// <summary>
        /// File
        /// </summary>
        File,

        /// <summary>
        /// Directory
        /// </summary>
        Directory,

        /// <summary>
        /// File or Directory
        /// </summary>
        Any
    }

    /// <summary>
    /// Quick Access management interface
    /// </summary>
    public interface IQuickAccessManager : IDisposable
    {
        /// <summary>
        /// Checks system Quick Access feature availability
        /// </summary>
        /// <returns>Tuple (Query possible, Operation possible) indicating system capability</returns>
        Task<(bool QueryFeasible, bool HandleFeasible)> CheckFeasibleAsync();

        /// <summary>
        /// Retrieves Quick Access items
        /// </summary>
        /// <param name="qaType">Quick Access type</param>
        /// <returns>List of paths for specified Quick Access type</returns>
        Task<List<string>> GetItemsAsync(QuickAccess qaType);

        /// <summary>
        /// Checks if item exists in Quick Access
        /// </summary>
        /// <param name="path">Path to check</param>
        /// <param name="qaType">Quick Access type</param>
        /// <returns>True if item exists in specified category</returns>
        Task<bool> CheckItemAsync(string path, QuickAccess qaType);

        /// <summary>
        /// Adds item to Quick Access
        /// </summary>
        /// <param name="path">Path to add</param>
        /// <param name="qaType">Quick Access type</param>
        /// <param name="forceUpdate">Force UI update</param>
        /// <returns>True if operation succeeded</returns>
        Task<bool> AddItemAsync(string path, QuickAccess qaType, bool forceUpdate = false);

        /// <summary>
        /// Removes item from Quick Access
        /// </summary>
        /// <param name="path">Path to remove</param>
        /// <param name="qaType">Quick Access type</param>
        /// <returns>True if operation succeeded</returns>
        Task<bool> RemoveItemAsync(string path, QuickAccess qaType);

        /// <summary>
        /// Clears Quick Access items
        /// </summary>
        /// <param name="qaType">Quick Access type</param>
        /// <param name="forceRefresh">Force Explorer refresh</param>
        /// <param name="alsoSystemDefault">Also clear system defaults</param>
        /// <returns>True if operation succeeded</returns>
        Task<bool> EmptyItemsAsync(QuickAccess qaType, bool forceRefresh = false, bool alsoSystemDefault = false);

        /// <summary>
        /// Clears cache
        /// </summary>
        void ClearCache();
    }

    /// <summary>
    /// File system operations interface for unit test mocking
    /// </summary>
    public interface IFileSystemOperations
    {
        bool FileExists(string path);
        bool DirectoryExists(string path);
        void DeleteFile(string path);
    }

    /// <summary>
    /// Native methods interface for unit test mocking
    /// </summary>
    public interface INativeMethods
    {
        void SHAddToRecentDocs(uint uFlags, IntPtr pv);
        int CoInitializeEx(IntPtr reserved, uint dwCoInit);
        void CoUninitialize();
        int SHGetKnownFolderPath(Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr ppszPath);
        void CoTaskMemFree(IntPtr ptr);
    }

    /// <summary>
    /// Default file system operations implementation
    /// </summary>
    public class DefaultFileSystemOperations : IFileSystemOperations
    {
        public bool FileExists(string path) => File.Exists(path);
        public bool DirectoryExists(string path) => Directory.Exists(path);
        public void DeleteFile(string path) => File.Delete(path);
    }

    /// <summary>
    /// Default native methods implementation
    /// </summary>
    public class DefaultNativeMethods : INativeMethods
    {
        public void SHAddToRecentDocs(uint uFlags, IntPtr pv) => NativeMethods.SHAddToRecentDocs(uFlags, pv);
        public int CoInitializeEx(IntPtr reserved, uint dwCoInit) => NativeMethods.CoInitializeEx(reserved, dwCoInit);
        public void CoUninitialize() => NativeMethods.CoUninitialize();
        public int SHGetKnownFolderPath(Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr ppszPath)
            => NativeMethods.SHGetKnownFolderPath(rfid, dwFlags, hToken, out ppszPath);
        public void CoTaskMemFree(IntPtr ptr) => NativeMethods.CoTaskMemFree(ptr);
    }

    /// <summary>
    /// Windows native method constants and functions
    /// </summary>
    public static class NativeMethods
    {
        public const uint SHARD_PATHW = 0x00000003;
        public const int COINIT_APARTMENTTHREADED = 0x2;
        public const int COINIT_DISABLE_OLE1DDE = 0x4;
        public static readonly Guid FOLDERID_Recent = new Guid("{AE50C081-EBD2-438A-8655-8A092E34987A}");
        public static readonly string[] ProtectedPaths = { "System32", "Program Files" };

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        public static extern void SHAddToRecentDocs(uint uFlags, IntPtr pv);

        [DllImport("ole32.dll")]
        public static extern int CoInitializeEx(IntPtr reserved, uint dwCoInit);

        [DllImport("ole32.dll")]
        public static extern void CoUninitialize();

        [DllImport("shell32.dll")]
        public static extern int SHGetKnownFolderPath(
            [MarshalAs(UnmanagedType.LPStruct)] Guid rfid,
            uint dwFlags,
            IntPtr hToken,
            out IntPtr ppszPath);

        [DllImport("ole32.dll")]
        public static extern void CoTaskMemFree(IntPtr ptr);
    }

    /// <summary>
    /// Provides Windows Quick Access management capabilities
    /// </summary>
    public class QuickAccessManager : IQuickAccessManager
    {
        private readonly IScriptExecutor _executor;
        private readonly Lazy<Task<ExecutionFeasibilityStatus>> _feasibilityStatus;
        private readonly TimeSpan _lockTimeout;
        private readonly IFileSystemOperations _fileSystem;
        private readonly INativeMethods _nativeMethods;
        private readonly IQuickAccessDataFiles _dataFiles;

        /// <summary>
        /// Initializes new instance of <see cref="QuickAccessManager"/>
        /// </summary>
        public QuickAccessManager()
            : this(new ScriptExecutor(), TimeSpan.FromSeconds(10),
                  new DefaultFileSystemOperations(), new DefaultNativeMethods(), new QuickAccessDataFiles())
        {
        }

        /// <summary>
        /// Initializes new instance with specified executor and timeout
        /// </summary>
        /// <param name="executor">Script executor</param>
        /// <param name="lockTimeout">Lock timeout duration</param>
        public QuickAccessManager(IScriptExecutor executor, TimeSpan lockTimeout)
            : this(executor, lockTimeout, new DefaultFileSystemOperations(), new DefaultNativeMethods(), new QuickAccessDataFiles())
        {
        }

        /// <summary>
        /// Initializes new instance with specified dependencies
        /// </summary>
        public QuickAccessManager(
            IScriptExecutor executor,
            TimeSpan lockTimeout,
            IFileSystemOperations fileSystem,
            INativeMethods nativeMethods,
            IQuickAccessDataFiles dataFiles)
        {
            _executor = executor ?? throw new ArgumentNullException(nameof(executor));
            _lockTimeout = lockTimeout;
            _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));
            _nativeMethods = nativeMethods ?? throw new ArgumentNullException(nameof(nativeMethods));
            _dataFiles = dataFiles ?? throw new ArgumentNullException(nameof(dataFiles));
            _feasibilityStatus = new Lazy<Task<ExecutionFeasibilityStatus>>(() =>
                CheckFeasibilityInternalAsync());
        }

        private async Task<ExecutionFeasibilityStatus> CheckFeasibilityInternalAsync()
        {
            return await ExecutionFeasibilityStatus.CheckAsync(_executor, (int)_lockTimeout.TotalSeconds);
        }

        /// <summary>
        /// Checks system Quick Access feature availability
        /// </summary>
        /// <returns>Tuple indicating system capability</returns>
        public async Task<(bool QueryFeasible, bool HandleFeasible)> CheckFeasibleAsync()
        {
            var status = await _feasibilityStatus.Value;
            return (status.Query, status.Handle);
        }

        /// <summary>
        /// Validates path
        /// </summary>
        /// <param name="path">Path to validate</param>
        /// <param name="pathType">Expected path type</param>
        /// <exception cref="ArgumentException">Invalid path format</exception>
        /// <exception cref="FileNotFoundException">Path not found</exception>
        /// <exception cref="SecurityException">Protected system path</exception>
        public static void ValidatePath(string path, PathType pathType, IFileSystemOperations fileSystem = null)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Path cannot be empty", nameof(path));

            if (path.IndexOf("System32", StringComparison.OrdinalIgnoreCase) >= 0 ||
                path.IndexOf("Program Files", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                throw new SecurityException($"Protected system path: {path}");
            }

            fileSystem = fileSystem ?? new DefaultFileSystemOperations();

            bool exists = false;
            switch (pathType)
            {
                case PathType.File:
                    exists = fileSystem.FileExists(path);
                    break;
                case PathType.Directory:
                    exists = fileSystem.DirectoryExists(path);
                    break;
                case PathType.Any:
                    exists = fileSystem.FileExists(path) || fileSystem.DirectoryExists(path);
                    break;
            }

            if (!exists)
                throw new FileNotFoundException($"Path not found: {path}", path);
        }

        private PSScript MapToScriptType(QuickAccess qeType)
        {
            switch (qeType)
            {
                case QuickAccess.All:
                    return PSScript.QueryQuickAccess;
                case QuickAccess.RecentFiles:
                    return PSScript.QueryRecentFile;
                case QuickAccess.FrequentFolders:
                    return PSScript.QueryFrequentFolder;
                default:
                    throw new ArgumentOutOfRangeException(nameof(qeType), qeType, "Unsupported Quick Access type");
            }
        }

        /// <summary>
        /// Retrieves Quick Access items
        /// </summary>
        public async Task<List<string>> GetItemsAsync(QuickAccess qaType)
        {
            var scriptType = MapToScriptType(qaType);
            return await _executor.ExecutePSScriptWithCache(scriptType, null, 10);
        }

        /// <summary>
        /// Checks if item exists in Quick Access
        /// </summary>
        public async Task<bool> CheckItemAsync(string path, QuickAccess qaType)
        {
            var items = await GetItemsAsync(qaType);
            return items.Any(item => string.Equals(item, path, StringComparison.OrdinalIgnoreCase));
        }

        private async Task<bool> HandleOperationAsync(
            bool isAdd,
            string path,
            QuickAccess qaType,
            PathType pathType,
            bool forceUpdate)
        {
            ValidatePath(path, pathType, _fileSystem);

            PSScript script;
            if (isAdd)
            {
                switch (qaType)
                {
                    case QuickAccess.RecentFiles:
                        AddFileToRecentDocs(path);
                        if (forceUpdate)
                        {
                            _dataFiles.RemoveRecentFile();
                        }
                        return true;
                    case QuickAccess.FrequentFolders:
                        script = PSScript.PinToFrequentFolder;
                        break;
                    default:
                        throw new NotSupportedException($"Unsupported add operation: {qaType}");
                }
            }
            else
            {
                switch (qaType)
                {
                    case QuickAccess.RecentFiles:
                        script = PSScript.RemoveRecentFile;
                        break;
                    case QuickAccess.FrequentFolders:
                        script = PSScript.UnpinFromFrequentFolder;
                        break;
                    default:
                        throw new NotSupportedException($"Unsupported remove operation: {qaType}");
                }
            }

            var result = await _executor.ExecutePSScriptWithTimeout(script, path, 10);
            if (result.ExitCode != 0)
            {
                return false;
            }

            _executor.ClearCache();
            return true;
        }

        /// <summary>
        /// Adds item to Quick Access
        /// </summary>
        public async Task<bool> AddItemAsync(string path, QuickAccess qaType, bool forceUpdate = false)
        {
            if (await CheckItemAsync(path, qaType))
            {
                throw new InvalidOperationException($"Item already exists: {path}");
            }

            PathType pathType = qaType == QuickAccess.RecentFiles
                ? PathType.File
                : PathType.Directory;

            return await HandleOperationAsync(true, path, qaType, pathType, forceUpdate);
        }

        /// <summary>
        /// Removes item from Quick Access
        /// </summary>
        public async Task<bool> RemoveItemAsync(string path, QuickAccess qaType)
        {
            if (!await CheckItemAsync(path, qaType))
            {
                throw new InvalidOperationException($"Item not found: {path}");
            }

            PathType pathType = qaType == QuickAccess.RecentFiles
                ? PathType.File
                : PathType.Directory;

            return await HandleOperationAsync(false, path, qaType, pathType, false);
        }

        /// <summary>
        /// Clears Quick Access items
        /// </summary>
        public async Task<bool> EmptyItemsAsync(QuickAccess qaType, bool forceRefresh = false, bool alsoSystemDefault = false)
        {
            try
            {
                switch (qaType)
                {
                    case QuickAccess.RecentFiles:
                        EmptyRecentFiles();
                        break;

                    case QuickAccess.FrequentFolders:
                        await EmptyFrequentFolders(alsoSystemDefault);
                        break;

                    case QuickAccess.All:
                        await EmptyItemsAsync(QuickAccess.RecentFiles, false, alsoSystemDefault);
                        await EmptyItemsAsync(QuickAccess.FrequentFolders, false, alsoSystemDefault);
                        break;
                }

                _executor.ClearCache();

                if (forceRefresh)
                {
                    await _executor.ExecutePSScript(PSScript.RefreshExplorer, null);
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to clear Quick Access items: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Clears cache
        /// </summary>
        public void ClearCache()
        {
            _executor.ClearCache();
        }

        private void AddFileToRecentDocs(string filePath)
        {
            ValidatePath(filePath, PathType.File, _fileSystem);
            IntPtr pathPtr = IntPtr.Zero;
            bool comInitialized = false;
            try
            {
                // First attempt to uninitialize any existing COM context
                try
                {
                    _nativeMethods.CoUninitialize();
                }
                catch { /* Ignore uninitialization failure */ }
                // Re-initialize COM with combined flags
                int hr = _nativeMethods.CoInitializeEx(IntPtr.Zero, 
                    NativeMethods.COINIT_APARTMENTTHREADED | NativeMethods.COINIT_DISABLE_OLE1DDE);
                
                if (hr != 0 && hr != 1)
                {
                    throw new Win32Exception(hr, $"COM initialization failed with error code: 0x{hr:X8}");
                }
                comInitialized = true;
                // Allocate unmanaged memory for the path string
                pathPtr = Marshal.StringToHGlobalUni(filePath);
                
                // Add to recent documents using shell API
                _nativeMethods.SHAddToRecentDocs(NativeMethods.SHARD_PATHW, pathPtr);
            }
            finally
            {
                // Clean up unmanaged resources
                if (pathPtr != IntPtr.Zero) 
                    Marshal.FreeHGlobal(pathPtr);
                    
                // Properly uninitialize COM if we initialized it
                if (comInitialized)
                {
                    _nativeMethods.CoUninitialize();
                }
            }
        }

        private void EmptyRecentFiles()
        {
            try
            {
                _nativeMethods.CoInitializeEx(IntPtr.Zero, NativeMethods.COINIT_APARTMENTTHREADED);
                _nativeMethods.SHAddToRecentDocs(NativeMethods.SHARD_PATHW, IntPtr.Zero);
            }
            finally
            {
                _nativeMethods.CoUninitialize();
            }
        }

        private string GetKnownFolderPath(Guid knownFolderId)
        {
            IntPtr pPath = IntPtr.Zero;
            try
            {
                int hr = _nativeMethods.SHGetKnownFolderPath(knownFolderId, 0, IntPtr.Zero, out pPath);
                if (hr != 0) throw new Win32Exception(hr);

                return Marshal.PtrToStringUni(pPath);
            }
            finally
            {
                if (pPath != IntPtr.Zero)
                    _nativeMethods.CoTaskMemFree(pPath);
            }
        }

        private async Task EmptyFrequentFolders(bool alsoSystemDefault)
        {
            var jumplistPath = GetKnownFolderPath(NativeMethods.FOLDERID_Recent);
            var jumplistFile = Path.Combine(jumplistPath,
                "AutomaticDestinations",
                "f01b4d95cf55d32a.automaticDestinations-ms");

            if (_fileSystem.FileExists(jumplistFile))
            {
                try { _fileSystem.DeleteFile(jumplistFile); }
                catch { /* Ignore deletion failure */ }
            }

            if (alsoSystemDefault)
            {
                await _executor.ExecutePSScriptWithCache(PSScript.EmptyPinnedFolders, null);
            }
        }

        /// <summary>
        /// Releases resources
        /// </summary>
        public void Dispose()
        {
            _executor.Dispose();
        }
    }

    /// <summary>
    /// Exception for Quick Access operations
    /// </summary>
    public class QuickAccessOperationException : Exception
    {
        /// <summary>
        /// Initializes a new instance with specified error message
        /// </summary>
        /// <param name="message">Error message describing the exception</param>
        public QuickAccessOperationException(string message) : base(message) { }
        /// <summary>
        /// Initializes a new instance with specified error message and inner exception
        /// </summary>
        /// <param name="message">Error message describing the exception</param>
        /// <param name="inner">Inner exception that caused this exception</param>
        public QuickAccessOperationException(string message, Exception inner)
            : base(message, inner) { }
    }
}