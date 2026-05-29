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
        private readonly IQuickAccessNativeQuery _nativeQuery;
        private readonly IQuickAccessNativeMutation _nativeMutation;
        private readonly IExplorerRefresher _explorerRefresher;
        private readonly IRecentLinksCleaner _recentLinksCleaner;

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
                  options.RetryPolicy ?? RetryPolicy.Standard,
                  new ShellQuickAccessNativeQuery(new DefaultNativeMethods()),
                  new ShellQuickAccessNativeMutation(new DefaultNativeMethods()),
                  new ShellExplorerRefresher(new DefaultNativeMethods()),
                  new RecentLinksCleaner(
                      new WindowsRecentFolder(new DefaultNativeMethods()),
                      new ShellLinkTargetResolver(new DefaultNativeMethods()),
                      new DefaultRecentLinkFileSystem()))
        {
        }

        internal QuickAccessManager(
            IScriptExecutor executor,
            TimeSpan timeout,
            IFileSystemOperations fileSystem,
            INativeMethods nativeMethods,
            IQuickAccessDataFiles dataFiles)
            // Dependency-injected instances keep the legacy PowerShell query path unless a native seam is supplied.
            : this(
                  executor,
                  timeout,
                  fileSystem,
                  nativeMethods,
                  dataFiles,
                  RetryPolicy.Standard,
                  new PowerShellFallbackNativeQuery(),
                  new PowerShellFallbackNativeMutation(),
                  new PowerShellFallbackExplorerRefresher(),
                  new NoOpRecentLinksCleaner())
        {
        }

        internal QuickAccessManager(
            IScriptExecutor executor,
            TimeSpan timeout,
            IFileSystemOperations fileSystem,
            INativeMethods nativeMethods,
            IQuickAccessDataFiles dataFiles,
            RetryPolicy retryPolicy)
            : this(
                  executor,
                  timeout,
                  fileSystem,
                  nativeMethods,
                  dataFiles,
                  retryPolicy,
                  new PowerShellFallbackNativeQuery(),
                  new PowerShellFallbackNativeMutation(),
                  new PowerShellFallbackExplorerRefresher(),
                  new NoOpRecentLinksCleaner())
        {
        }

        internal QuickAccessManager(
            IScriptExecutor executor,
            TimeSpan timeout,
            IFileSystemOperations fileSystem,
            INativeMethods nativeMethods,
            IQuickAccessDataFiles dataFiles,
            RetryPolicy retryPolicy,
            IQuickAccessNativeQuery nativeQuery)
            : this(
                  executor,
                  timeout,
                  fileSystem,
                  nativeMethods,
                  dataFiles,
                  retryPolicy,
                  nativeQuery,
                  new PowerShellFallbackNativeMutation(),
                  new PowerShellFallbackExplorerRefresher(),
                  new NoOpRecentLinksCleaner())
        {
        }

        internal QuickAccessManager(
            IScriptExecutor executor,
            TimeSpan timeout,
            IFileSystemOperations fileSystem,
            INativeMethods nativeMethods,
            IQuickAccessDataFiles dataFiles,
            RetryPolicy retryPolicy,
            IQuickAccessNativeQuery nativeQuery,
            IQuickAccessNativeMutation nativeMutation)
            : this(
                  executor,
                  timeout,
                  fileSystem,
                  nativeMethods,
                  dataFiles,
                  retryPolicy,
                  nativeQuery,
                  nativeMutation,
                  new PowerShellFallbackExplorerRefresher(),
                  new NoOpRecentLinksCleaner())
        {
        }

        internal QuickAccessManager(
            IScriptExecutor executor,
            TimeSpan timeout,
            IFileSystemOperations fileSystem,
            INativeMethods nativeMethods,
            IQuickAccessDataFiles dataFiles,
            RetryPolicy retryPolicy,
            IQuickAccessNativeQuery nativeQuery,
            IQuickAccessNativeMutation nativeMutation,
            IExplorerRefresher explorerRefresher)
            : this(
                  executor,
                  timeout,
                  fileSystem,
                  nativeMethods,
                  dataFiles,
                  retryPolicy,
                  nativeQuery,
                  nativeMutation,
                  explorerRefresher,
                  new NoOpRecentLinksCleaner())
        {
        }

        internal QuickAccessManager(
            IScriptExecutor executor,
            TimeSpan timeout,
            IFileSystemOperations fileSystem,
            INativeMethods nativeMethods,
            IQuickAccessDataFiles dataFiles,
            RetryPolicy retryPolicy,
            IQuickAccessNativeQuery nativeQuery,
            IQuickAccessNativeMutation nativeMutation,
            IExplorerRefresher explorerRefresher,
            IRecentLinksCleaner recentLinksCleaner)
        {
            if (timeout <= TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(timeout), "Timeout must be positive.");

            _executor = executor ?? throw new ArgumentNullException(nameof(executor));
            _timeout = timeout;
            _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));
            _nativeMethods = nativeMethods ?? throw new ArgumentNullException(nameof(nativeMethods));
            _dataFiles = dataFiles ?? throw new ArgumentNullException(nameof(dataFiles));
            _retryPolicy = retryPolicy ?? throw new ArgumentNullException(nameof(retryPolicy));
            _nativeQuery = nativeQuery ?? throw new ArgumentNullException(nameof(nativeQuery));
            _nativeMutation = nativeMutation ?? throw new ArgumentNullException(nameof(nativeMutation));
            _explorerRefresher = explorerRefresher ?? throw new ArgumentNullException(nameof(explorerRefresher));
            _recentLinksCleaner = recentLinksCleaner ?? throw new ArgumentNullException(nameof(recentLinksCleaner));
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
            try
            {
                return _nativeQuery.GetItems(target);
            }
            catch (Exception)
            {
                return ExecuteListScript(script, null, ToTimeoutSeconds());
            }
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
        /// data is removed to force Explorer to rebuild it. Frequent folders are pinned with a native Shell verb first
        /// and fall back to PowerShell only when the native Shell operation fails.
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

                ExecuteNativeMutationWithPowerShellFallback(
                    () => _nativeMutation.PinFrequentFolder(path, _timeout),
                    PSScript.PinToFrequentFolder,
                    path,
                    PowerShellOperation.PinFrequentFolder);
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
        /// This method modifies the current Windows user's Quick Access state. Removal uses native Shell verbs first
        /// and falls back to PowerShell only for native Shell failures. When
        /// <see cref="RemoveOptions.DeepCleanRecentLinks"/> is enabled, matching shortcut files in the Windows Recent
        /// folder are deleted after the Shell removal succeeds.
        /// </remarks>
        public void RemoveItem(string path, QuickAccess target, RemoveOptions options)
        {
            if (path == null)
                throw new ArgumentNullException(nameof(path));
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            EnsureSingleItemTarget(target, "RemoveItem");

            bool removed = false;
            try
            {
                if (target == QuickAccess.RecentFiles)
                {
                    ValidatePath(path, PathType.File, _fileSystem);
                    if (!ContainsItemExact(path, target))
                        throw new QuickAccessItemNotFoundException(path, target);

                    ExecuteNativeMutationWithPowerShellFallback(
                        () => _nativeMutation.RemoveRecentFile(path, _timeout),
                        PSScript.RemoveRecentFile,
                        path,
                        PowerShellOperation.RemoveRecentFile);
                }
                else
                {
                    ValidatePath(path, PathType.Directory, _fileSystem);
                    if (!ContainsItemExact(path, target))
                        throw new QuickAccessItemNotFoundException(path, target);

                    ExecuteNativeMutationWithPowerShellFallback(
                        () => _nativeMutation.UnpinFrequentFolder(path, _timeout),
                        PSScript.UnpinFromFrequentFolder,
                        path,
                        PowerShellOperation.UnpinFrequentFolder);
                }

                removed = true;
                if (options.DeepCleanRecentLinks)
                    _recentLinksCleaner.DeleteForTarget(path, _timeout);
            }
            finally
            {
                if (removed)
                    _executor.ClearCache();
            }
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
            int refreshRecentItemIndex = -1;

            foreach (var item in items)
            {
                if (item == null)
                    throw new ArgumentException("Batch items cannot contain null.", nameof(items));

                try
                {
                    AddItem(item.Path, item.Target, new AddOptions());
                    succeeded.Add(item);
                    if (options.RefreshRecentFiles && item.Target == QuickAccess.RecentFiles)
                        refreshRecentItemIndex = succeeded.Count - 1;
                }
                catch (Exception ex)
                {
                    failed.Add(new BatchFailure(item, ex));
                }
            }

            if (refreshRecentItemIndex >= 0)
            {
                try
                {
                    _dataFiles.RemoveRecentFile();
                }
                catch (Exception ex)
                {
                    var item = succeeded[refreshRecentItemIndex];
                    succeeded.RemoveAt(refreshRecentItemIndex);
                    failed.Add(new BatchFailure(item, ex));
                }
            }

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
        /// <exception cref="PartialClearException">The clear operation only partially succeeds.</exception>
        /// <remarks>
        /// This method modifies the current Windows user's Quick Access state. Explorer refresh uses a native Shell
        /// refresh first and falls back to PowerShell when the native refresh fails. Refresh failures are propagated only
        /// after a fully successful clear; partial clear errors preserve the original clear failure.
        /// </remarks>
        public void ClearItems(QuickAccess target, ClearOptions options)
        {
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            EnsureClearTarget(target);

            bool recentCleared = false;
            bool frequentCleared = false;
            Exception firstError = null;
            PartialClearException partialError = null;

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
                }
            }

            if (target == QuickAccess.FrequentFolders || target == QuickAccess.All)
            {
                try
                {
                    ClearFrequentFolders(options.RemovePinnedFolders);
                    frequentCleared = true;
                }
                catch (PartialClearException ex)
                {
                    frequentCleared |= ex.FrequentFoldersCleared;
                    partialError = partialError ?? ex;
                    firstError = firstError ?? ex.InnerException ?? ex;
                }
                catch (Exception ex)
                {
                    firstError = firstError ?? ex;
                }
            }

            if (recentCleared || frequentCleared || firstError == null)
                _executor.ClearCache();

            if (firstError != null)
            {
                bool hasPartialProgress = partialError != null ||
                                          target == QuickAccess.All && (recentCleared || frequentCleared);
                if (hasPartialProgress)
                    TryRefreshExplorer(options.RefreshExplorer);

                if (partialError != null && target != QuickAccess.All)
                    throw partialError;

                if (target == QuickAccess.All)
                    throw new PartialClearException(recentCleared, frequentCleared, firstError);

                throw firstError;
            }

            if (options.RefreshExplorer)
                RefreshExplorer();
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
            return ExecuteWithRetry(() =>
                _executor.ExecutePSScriptWithCache(script, parameter, timeoutSeconds)
                    .GetAwaiter()
                    .GetResult()
                    .AsReadOnly());
        }

        private void ExecuteMutationScript(PSScript script, string parameter, PowerShellOperation operation)
        {
            ExecuteWithRetry(
                () =>
                {
                    _executor.ExecutePSScriptWithTimeout(script, parameter, ToTimeoutSeconds())
                        .GetAwaiter()
                        .GetResult();

                    return true;
                });
        }

        private void ExecuteNativeMutationWithPowerShellFallback(
            Action nativeAction,
            PSScript fallbackScript,
            string parameter,
            PowerShellOperation operation)
        {
            try
            {
                nativeAction();
            }
            catch (Exception ex) when (ShouldFallbackToPowerShell(ex))
            {
                ExecuteMutationScript(fallbackScript, parameter, operation);
            }
        }

        private static bool ShouldFallbackToPowerShell(Exception exception)
        {
            if (exception is QuickAccessItemNotFoundException ||
                exception is QuickAccessItemAlreadyExistsException ||
                exception is UnsupportedQuickAccessOperationException ||
                exception is ArgumentException ||
                exception is FileNotFoundException ||
                exception is DirectoryNotFoundException)
            {
                return false;
            }

            return true;
        }

        private T ExecuteWithRetry<T>(Func<T> action)
        {
            for (int retryAttempt = 0; ; retryAttempt++)
            {
                try
                {
                    return action();
                }
                catch (PowerShellExecutionException ex) when (ShouldRetry(ex, retryAttempt))
                {
                    System.Threading.Thread.Sleep(_retryPolicy.GetDelay(retryAttempt));
                }
            }
        }

        private bool ShouldRetry(PowerShellExecutionException exception, int retryAttempt)
        {
            if (_retryPolicy.MaxRetryCount == 0 || retryAttempt >= _retryPolicy.MaxRetryCount)
                return false;

            if (exception.Kind == PowerShellErrorKind.Timeout)
                return true;

            string stderr = exception.StandardError ?? string.Empty;
            return stderr.IndexOf("locked", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   stderr.IndexOf("temporarily unavailable", StringComparison.OrdinalIgnoreCase) >= 0;
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
            return script.ToPowerShellOperation();
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
            var jumpListFile = _dataFiles.FrequentFoldersPath;

            if (_fileSystem.FileExists(jumpListFile))
                _fileSystem.DeleteFile(jumpListFile);

            if (removePinnedFolders)
            {
                try
                {
                    ExecuteMutationScript(PSScript.EmptyPinnedFolders, null, PowerShellOperation.ClearPinnedFolders);
                }
                catch (Exception ex)
                {
                    throw new PartialClearException(false, true, ex);
                }
            }
        }

        private void TryRefreshExplorer(bool shouldRefresh)
        {
            if (!shouldRefresh)
                return;

            try
            {
                RefreshExplorer();
            }
            catch
            {
                // Partial clear preserves the original failure.
            }
        }

        private void RefreshExplorer()
        {
            try
            {
                _explorerRefresher.Refresh(_timeout);
            }
            catch (Exception)
            {
                ExecuteMutationScript(PSScript.RefreshExplorer, null, PowerShellOperation.RefreshExplorer);
            }
        }
    }
}
