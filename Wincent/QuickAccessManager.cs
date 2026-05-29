using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace Wincent
{
    /// <summary>
    /// Identifies a Windows Quick Access section.
    /// </summary>
    public enum QuickAccess
    {
        /// <summary>
        /// Both recent files and frequent folders.
        /// </summary>
        All = 0,

        /// <summary>
        /// Recently used files.
        /// </summary>
        RecentFiles = 1,

        /// <summary>
        /// Frequently used folders.
        /// </summary>
        FrequentFolders = 2
    }

    internal enum PathType
    {
        File,
        Directory,
        Any
    }

    internal interface IFileSystemOperations
    {
        bool FileExists(string path);
        bool DirectoryExists(string path);
        void DeleteFile(string path);
    }

    internal interface INativeMethods
    {
        void SHAddToRecentDocs(uint uFlags, IntPtr pv);
        int CoInitializeEx(IntPtr reserved, uint dwCoInit);
        void CoUninitialize();
        int SHGetKnownFolderPath(Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr ppszPath);
        void CoTaskMemFree(IntPtr ptr);
    }

    internal sealed class DefaultFileSystemOperations : IFileSystemOperations
    {
        public bool FileExists(string path) => File.Exists(path);

        public bool DirectoryExists(string path) => Directory.Exists(path);

        public void DeleteFile(string path) => File.Delete(path);
    }

    internal sealed class DefaultNativeMethods : INativeMethods
    {
        public void SHAddToRecentDocs(uint uFlags, IntPtr pv) => NativeMethods.SHAddToRecentDocs(uFlags, pv);

        public int CoInitializeEx(IntPtr reserved, uint dwCoInit) => NativeMethods.CoInitializeEx(reserved, dwCoInit);

        public void CoUninitialize() => NativeMethods.CoUninitialize();

        public int SHGetKnownFolderPath(Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr ppszPath)
            => NativeMethods.SHGetKnownFolderPath(rfid, dwFlags, hToken, out ppszPath);

        public void CoTaskMemFree(IntPtr ptr) => NativeMethods.CoTaskMemFree(ptr);
    }

    internal static class NativeMethods
    {
        public const uint SHARD_PATHW = 0x00000003;
        public const int COINIT_APARTMENTTHREADED = 0x2;
        public const int COINIT_DISABLE_OLE1DDE = 0x4;
        public static readonly Guid FOLDERID_Recent = new Guid("{AE50C081-EBD2-438A-8655-8A092E34987A}");

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
    /// Provides synchronous operations for Windows Quick Access.
    /// </summary>
    /// <remarks>
    /// Instances are safe to use from multiple threads because configuration is copied at construction time.
    /// Windows Explorer state may still change between calls, so thread safety does not imply a stable snapshot.
    /// Operations that add, remove, or clear items affect the current Windows user's Explorer state.
    /// </remarks>
    public sealed class QuickAccessManager : IDisposable
    {
        private readonly IScriptExecutor _executor;
        private readonly TimeSpan _timeout;
        private readonly RetryPolicy _retryPolicy;
        private readonly IFileSystemOperations _fileSystem;
        private readonly INativeMethods _nativeMethods;
        private readonly IQuickAccessDataFiles _dataFiles;

        /// <summary>
        /// Initializes a manager with default options.
        /// </summary>
        public QuickAccessManager()
            : this(new QuickAccessManagerOptions())
        {
        }

        /// <summary>
        /// Initializes a manager with the supplied options.
        /// </summary>
        /// <param name="options">The manager options to copy.</param>
        /// <exception cref="ArgumentNullException"><paramref name="options"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><see cref="QuickAccessManagerOptions.Timeout"/> is not positive.</exception>
        public QuickAccessManager(QuickAccessManagerOptions options)
            : this(
                  new ScriptExecutor(),
                  GetValidatedTimeout(options),
                  new DefaultFileSystemOperations(),
                  new DefaultNativeMethods(),
                  new QuickAccessDataFiles(),
                  options.RetryPolicy ?? RetryPolicy.Standard)
        {
        }

        internal QuickAccessManager(
            IScriptExecutor executor,
            TimeSpan timeout,
            IFileSystemOperations fileSystem,
            INativeMethods nativeMethods,
            IQuickAccessDataFiles dataFiles)
            : this(executor, timeout, fileSystem, nativeMethods, dataFiles, RetryPolicy.Standard)
        {
        }

        internal QuickAccessManager(
            IScriptExecutor executor,
            TimeSpan timeout,
            IFileSystemOperations fileSystem,
            INativeMethods nativeMethods,
            IQuickAccessDataFiles dataFiles,
            RetryPolicy retryPolicy)
        {
            if (timeout <= TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(timeout), "Timeout must be positive.");

            _executor = executor ?? throw new ArgumentNullException(nameof(executor));
            _timeout = timeout;
            _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));
            _nativeMethods = nativeMethods ?? throw new ArgumentNullException(nameof(nativeMethods));
            _dataFiles = dataFiles ?? throw new ArgumentNullException(nameof(dataFiles));
            _retryPolicy = retryPolicy ?? throw new ArgumentNullException(nameof(retryPolicy));
        }

        /// <summary>
        /// Gets the operation timeout copied from the constructor options.
        /// </summary>
        public TimeSpan Timeout => _timeout;

        /// <summary>
        /// Gets the retry policy copied from the constructor options.
        /// </summary>
        public RetryPolicy RetryPolicy => _retryPolicy;

        /// <summary>
        /// Retrieves paths from a Quick Access section.
        /// </summary>
        /// <param name="target">The section to query.</param>
        /// <returns>The current item paths.</returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="target"/> is not supported.</exception>
        /// <exception cref="PowerShellExecutionException">The PowerShell fallback query fails.</exception>
        public IReadOnlyList<string> GetItems(QuickAccess target)
        {
            var script = MapQueryScript(target);
            return ExecuteListScript(script, null, ToTimeoutSeconds());
        }

        /// <summary>
        /// Determines whether any item contains the supplied keyword.
        /// </summary>
        /// <param name="keyword">The case-sensitive substring to search for.</param>
        /// <param name="target">The section to query.</param>
        /// <returns><see langword="true"/> when an item contains <paramref name="keyword"/>.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="keyword"/> is <see langword="null"/>.</exception>
        /// <remarks>
        /// This method follows the Rust upstream substring semantics. It is not Windows path equality and may match
        /// paths that are not the exact file or folder intended by the caller.
        /// </remarks>
        public bool ContainsItem(string keyword, QuickAccess target)
        {
            if (keyword == null)
                throw new ArgumentNullException(nameof(keyword));

            return GetItems(target).Any(item => item != null && item.Contains(keyword));
        }

        /// <summary>
        /// Determines whether a section contains an exact path using Windows path semantics.
        /// </summary>
        /// <param name="path">The path to compare.</param>
        /// <param name="target">The section to query.</param>
        /// <returns><see langword="true"/> when an item matches <paramref name="path"/>.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> is <see langword="null"/>.</exception>
        public bool ContainsItemExact(string path, QuickAccess target)
        {
            if (path == null)
                throw new ArgumentNullException(nameof(path));

            return GetItems(target).Any(item => WindowsPathComparer.Equals(item, path));
        }

        /// <summary>
        /// Adds a path to Windows Quick Access.
        /// </summary>
        /// <param name="path">The file or folder path to add.</param>
        /// <param name="target">The target section.</param>
        public void AddItem(string path, QuickAccess target)
        {
            AddItem(path, target, new AddOptions());
        }

        /// <summary>
        /// Adds a path to Windows Quick Access.
        /// </summary>
        /// <param name="path">The file or folder path to add.</param>
        /// <param name="target">The target section.</param>
        /// <param name="options">The add options.</param>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> or <paramref name="options"/> is <see langword="null"/>.</exception>
        /// <exception cref="QuickAccessItemAlreadyExistsException">The item is already present.</exception>
        /// <exception cref="UnsupportedQuickAccessOperationException"><paramref name="target"/> is <see cref="QuickAccess.All"/>.</exception>
        /// <remarks>
        /// This method modifies the current Windows user's Quick Access state. When
        /// <see cref="AddOptions.RefreshRecentFiles"/> is enabled for recent files, the current Recent Files backing
        /// data is removed to force Explorer to rebuild it.
        /// </remarks>
        public void AddItem(string path, QuickAccess target, AddOptions options)
        {
            if (path == null)
                throw new ArgumentNullException(nameof(path));
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            EnsureSingleItemTarget(target, "AddItem");

            if (target == QuickAccess.RecentFiles)
            {
                ValidatePath(path, PathType.File, _fileSystem);
                if (ContainsItemExact(path, target))
                    throw new QuickAccessItemAlreadyExistsException(path, target);

                AddFileToRecentDocs(path);
                if (options.RefreshRecentFiles)
                    _dataFiles.RemoveRecentFile();
            }
            else
            {
                ValidatePath(path, PathType.Directory, _fileSystem);
                if (ContainsItemExact(path, target))
                    throw new QuickAccessItemAlreadyExistsException(path, target);

                ExecuteMutationScript(PSScript.PinToFrequentFolder, path, PowerShellOperation.PinFrequentFolder);
            }

            _executor.ClearCache();
        }

        /// <summary>
        /// Removes a path from Windows Quick Access.
        /// </summary>
        /// <param name="path">The file or folder path to remove.</param>
        /// <param name="target">The target section.</param>
        public void RemoveItem(string path, QuickAccess target)
        {
            RemoveItem(path, target, new RemoveOptions());
        }

        /// <summary>
        /// Removes a path from Windows Quick Access.
        /// </summary>
        /// <param name="path">The file or folder path to remove.</param>
        /// <param name="target">The target section.</param>
        /// <param name="options">The remove options.</param>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> or <paramref name="options"/> is <see langword="null"/>.</exception>
        /// <exception cref="QuickAccessItemNotFoundException">The item is not present.</exception>
        /// <exception cref="UnsupportedQuickAccessOperationException"><paramref name="target"/> is <see cref="QuickAccess.All"/>.</exception>
        /// <remarks>
        /// This method modifies the current Windows user's Quick Access state. Deep cleanup of Recent shortcut files is
        /// reserved for a later migration phase.
        /// </remarks>
        public void RemoveItem(string path, QuickAccess target, RemoveOptions options)
        {
            if (path == null)
                throw new ArgumentNullException(nameof(path));
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            EnsureSingleItemTarget(target, "RemoveItem");

            if (target == QuickAccess.RecentFiles)
            {
                ValidatePath(path, PathType.File, _fileSystem);
                if (!ContainsItemExact(path, target))
                    throw new QuickAccessItemNotFoundException(path, target);

                ExecuteMutationScript(PSScript.RemoveRecentFile, path, PowerShellOperation.RemoveRecentFile);
            }
            else
            {
                ValidatePath(path, PathType.Directory, _fileSystem);
                if (!ContainsItemExact(path, target))
                    throw new QuickAccessItemNotFoundException(path, target);

                ExecuteMutationScript(PSScript.UnpinFromFrequentFolder, path, PowerShellOperation.UnpinFrequentFolder);
            }

            _executor.ClearCache();
        }

        /// <summary>
        /// Adds multiple Quick Access items.
        /// </summary>
        /// <param name="items">The items to add.</param>
        /// <returns>The batch result.</returns>
        public BatchResult AddItems(IEnumerable<QuickAccessItem> items)
        {
            return AddItems(items, new BatchOptions());
        }

        /// <summary>
        /// Adds multiple Quick Access items.
        /// </summary>
        /// <param name="items">The items to add.</param>
        /// <param name="options">The batch options.</param>
        /// <returns>The batch result.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="items"/> or <paramref name="options"/> is <see langword="null"/>.</exception>
        public BatchResult AddItems(IEnumerable<QuickAccessItem> items, BatchOptions options)
        {
            if (items == null)
                throw new ArgumentNullException(nameof(items));
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            var succeeded = new List<QuickAccessItem>();
            var failed = new List<BatchFailure>();
            bool refreshRecentFiles = false;

            foreach (var item in items)
            {
                if (item == null)
                    throw new ArgumentException("Batch items cannot contain null.", nameof(items));

                try
                {
                    AddItem(item.Path, item.Target, new AddOptions());
                    succeeded.Add(item);
                    refreshRecentFiles |= options.RefreshRecentFiles && item.Target == QuickAccess.RecentFiles;
                }
                catch (Exception ex)
                {
                    failed.Add(new BatchFailure(item, ex));
                }
            }

            if (refreshRecentFiles)
                _dataFiles.RemoveRecentFile();

            return new BatchResult(succeeded, failed);
        }

        /// <summary>
        /// Removes multiple Quick Access items.
        /// </summary>
        /// <param name="items">The items to remove.</param>
        /// <returns>The batch result.</returns>
        public BatchResult RemoveItems(IEnumerable<QuickAccessItem> items)
        {
            return RemoveItems(items, new RemoveOptions());
        }

        /// <summary>
        /// Removes multiple Quick Access items.
        /// </summary>
        /// <param name="items">The items to remove.</param>
        /// <param name="options">The remove options.</param>
        /// <returns>The batch result.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="items"/> or <paramref name="options"/> is <see langword="null"/>.</exception>
        public BatchResult RemoveItems(IEnumerable<QuickAccessItem> items, RemoveOptions options)
        {
            if (items == null)
                throw new ArgumentNullException(nameof(items));
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            var succeeded = new List<QuickAccessItem>();
            var failed = new List<BatchFailure>();

            foreach (var item in items)
            {
                if (item == null)
                    throw new ArgumentException("Batch items cannot contain null.", nameof(items));

                try
                {
                    RemoveItem(item.Path, item.Target, options);
                    succeeded.Add(item);
                }
                catch (Exception ex)
                {
                    failed.Add(new BatchFailure(item, ex));
                }
            }

            return new BatchResult(succeeded, failed);
        }

        /// <summary>
        /// Clears a Quick Access section.
        /// </summary>
        /// <param name="target">The section to clear.</param>
        public void ClearItems(QuickAccess target)
        {
            ClearItems(target, new ClearOptions());
        }

        /// <summary>
        /// Clears a Quick Access section.
        /// </summary>
        /// <param name="target">The section to clear.</param>
        /// <param name="options">The clear options.</param>
        /// <exception cref="ArgumentNullException"><paramref name="options"/> is <see langword="null"/>.</exception>
        /// <exception cref="PartialClearException">Only part of <see cref="QuickAccess.All"/> was cleared.</exception>
        /// <remarks>
        /// This method modifies the current Windows user's Quick Access state. Explorer refresh uses the current
        /// PowerShell fallback; native refresh is planned for a later migration phase.
        /// </remarks>
        public void ClearItems(QuickAccess target, ClearOptions options)
        {
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            EnsureClearTarget(target);

            bool recentCleared = false;
            bool frequentCleared = false;
            Exception firstError = null;

            if (target == QuickAccess.RecentFiles || target == QuickAccess.All)
            {
                try
                {
                    ClearRecentFiles();
                    recentCleared = true;
                }
                catch (Exception ex)
                {
                    firstError = firstError ?? ex;
                    if (target != QuickAccess.All)
                        throw;
                }
            }

            if (target == QuickAccess.FrequentFolders || target == QuickAccess.All)
            {
                try
                {
                    ClearFrequentFolders(options.RemovePinnedFolders);
                    frequentCleared = true;
                }
                catch (Exception ex)
                {
                    firstError = firstError ?? ex;
                    if (target != QuickAccess.All)
                        throw;
                }
            }

            _executor.ClearCache();

            if (target == QuickAccess.All && firstError != null)
            {
                TryRefreshExplorer(options.RefreshExplorer);
                throw new PartialClearException(recentCleared, frequentCleared, firstError);
            }

            if (options.RefreshExplorer)
                ExecuteMutationScript(PSScript.RefreshExplorer, null, PowerShellOperation.RefreshExplorer);
        }

        /// <summary>
        /// Releases resources owned by the manager.
        /// </summary>
        public void Dispose()
        {
            _executor.Dispose();
        }

        internal static void ValidatePath(string path, PathType pathType, IFileSystemOperations fileSystem = null)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Path cannot be empty.", nameof(path));

            fileSystem = fileSystem ?? new DefaultFileSystemOperations();

            bool exists;
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
                default:
                    throw new ArgumentOutOfRangeException(nameof(pathType), pathType, "Unsupported path type.");
            }

            if (!exists)
                throw new FileNotFoundException($"Path not found: {path}", path);
        }

        private static TimeSpan GetValidatedTimeout(QuickAccessManagerOptions options)
        {
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            if (options.Timeout <= TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(options), "Timeout must be positive.");

            return options.Timeout;
        }

        private static PSScript MapQueryScript(QuickAccess target)
        {
            switch (target)
            {
                case QuickAccess.All:
                    return PSScript.QueryQuickAccess;
                case QuickAccess.RecentFiles:
                    return PSScript.QueryRecentFile;
                case QuickAccess.FrequentFolders:
                    return PSScript.QueryFrequentFolder;
                default:
                    throw new ArgumentOutOfRangeException(nameof(target), target, "Unsupported Quick Access target.");
            }
        }

        private static void EnsureSingleItemTarget(QuickAccess target, string operation)
        {
            if (target == QuickAccess.All)
                throw new UnsupportedQuickAccessOperationException(target, operation);

            if (target != QuickAccess.RecentFiles && target != QuickAccess.FrequentFolders)
                throw new ArgumentOutOfRangeException(nameof(target), target, "Unsupported Quick Access target.");
        }

        private static void EnsureClearTarget(QuickAccess target)
        {
            if (target != QuickAccess.All && target != QuickAccess.RecentFiles && target != QuickAccess.FrequentFolders)
                throw new ArgumentOutOfRangeException(nameof(target), target, "Unsupported Quick Access target.");
        }

        private IReadOnlyList<string> ExecuteListScript(PSScript script, string parameter, int timeoutSeconds)
        {
            try
            {
                return _executor.ExecutePSScriptWithCache(script, parameter, timeoutSeconds)
                    .GetAwaiter()
                    .GetResult()
                    .AsReadOnly();
            }
            catch (ScriptExecutionException ex)
            {
                throw CreatePowerShellException(script, parameter, null, ex.Output, ex.Error, ex);
            }
        }

        private void ExecuteMutationScript(PSScript script, string parameter, PowerShellOperation operation)
        {
            var result = _executor.ExecutePSScriptWithTimeout(script, parameter, ToTimeoutSeconds())
                .GetAwaiter()
                .GetResult();

            if (result.ExitCode != 0)
            {
                throw new PowerShellExecutionException(
                    operation,
                    PowerShellErrorKind.ProcessFailed,
                    result.ExitCode,
                    result.Output,
                    result.Error,
                    null,
                    parameter,
                    null,
                    null);
            }
        }

        private PowerShellExecutionException CreatePowerShellException(
            PSScript script,
            string parameter,
            int? exitCode,
            string output,
            string error,
            Exception inner)
        {
            return new PowerShellExecutionException(
                MapPowerShellOperation(script),
                PowerShellErrorKind.ProcessFailed,
                exitCode,
                output,
                error,
                null,
                parameter,
                null,
                null,
                inner);
        }

        private static PowerShellOperation MapPowerShellOperation(PSScript script)
        {
            switch (script)
            {
                case PSScript.RefreshExplorer:
                    return PowerShellOperation.RefreshExplorer;
                case PSScript.QueryQuickAccess:
                    return PowerShellOperation.QueryQuickAccess;
                case PSScript.QueryRecentFile:
                    return PowerShellOperation.QueryRecentFiles;
                case PSScript.QueryFrequentFolder:
                    return PowerShellOperation.QueryFrequentFolders;
                case PSScript.AddRecentFile:
                    return PowerShellOperation.AddRecentFile;
                case PSScript.RemoveRecentFile:
                    return PowerShellOperation.RemoveRecentFile;
                case PSScript.PinToFrequentFolder:
                    return PowerShellOperation.PinFrequentFolder;
                case PSScript.UnpinFromFrequentFolder:
                    return PowerShellOperation.UnpinFrequentFolder;
                case PSScript.EmptyPinnedFolders:
                    return PowerShellOperation.ClearPinnedFolders;
                default:
                    return PowerShellOperation.QueryQuickAccess;
            }
        }

        private int ToTimeoutSeconds()
        {
            return Math.Max(1, (int)Math.Ceiling(_timeout.TotalSeconds));
        }

        private void AddFileToRecentDocs(string filePath)
        {
            IntPtr pathPtr = IntPtr.Zero;

            try
            {
                using (ComGuard.InitializeSta(_nativeMethods, disableOle1Dde: true))
                {
                    pathPtr = Marshal.StringToHGlobalUni(filePath);
                    _nativeMethods.SHAddToRecentDocs(NativeMethods.SHARD_PATHW, pathPtr);
                }
            }
            finally
            {
                if (pathPtr != IntPtr.Zero)
                    Marshal.FreeHGlobal(pathPtr);
            }
        }

        private void ClearRecentFiles()
        {
            using (ComGuard.InitializeSta(_nativeMethods, disableOle1Dde: true))
            {
                _nativeMethods.SHAddToRecentDocs(NativeMethods.SHARD_PATHW, IntPtr.Zero);
            }
        }

        private void ClearFrequentFolders(bool removePinnedFolders)
        {
            var recentFolder = new WindowsRecentFolder(_nativeMethods, _fileSystem).GetPath();
            var jumpListFile = Path.Combine(
                recentFolder,
                "AutomaticDestinations",
                "f01b4d95cf55d32a.automaticDestinations-ms");

            if (_fileSystem.FileExists(jumpListFile))
                _fileSystem.DeleteFile(jumpListFile);

            if (removePinnedFolders)
                ExecuteMutationScript(PSScript.EmptyPinnedFolders, null, PowerShellOperation.ClearPinnedFolders);
        }

        private void TryRefreshExplorer(bool shouldRefresh)
        {
            if (!shouldRefresh)
                return;

            try
            {
                ExecuteMutationScript(PSScript.RefreshExplorer, null, PowerShellOperation.RefreshExplorer);
            }
            catch
            {
                // Partial clear preserves the original failure.
            }
        }
    }
}
