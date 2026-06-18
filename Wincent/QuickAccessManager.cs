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
        private readonly IQuickAccessLockFactory _lockFactory;
        private readonly IQuickAccessVisibility _visibility;
        private readonly IDestListMetadataReader _destListReader;
        private readonly IQuickAccessRestoreEngine _restoreEngine;

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
                      new DefaultRecentLinkFileSystem()),
                  // Tech debt: consolidate default WindowsRecentFolder/QuickAccessDataFiles construction.
                  new QuickAccessLockFactory(
                      new QuickAccessDataFiles(),
                      new WindowsRecentFolder(new DefaultNativeMethods()),
                      new DefaultRecentLinkFileSystem(),
                      new NativeQuickAccessBackingFileHandleOpener()),
                  new RegistryQuickAccessVisibility(new CurrentUserExplorerVisibilityRegistry()),
                  new DefaultDestListMetadataReader(),
                  new QuickAccessRestoreEngine(
                      new QuickAccessDataFiles(),
                      new WindowsRecentFolder(new DefaultNativeMethods()),
                      new ShellLinkTargetResolver(new DefaultNativeMethods()),
                      new DefaultRecentLinkFileSystem(),
                      new DefaultFileSystemOperations(),
                      new DefaultDestListMetadataReader(),
                      new DefaultQuickAccessRestoreDelay()))
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
                  new NoOpRecentLinksCleaner(),
                  new NoOpQuickAccessLockFactory(),
                  new NoOpQuickAccessVisibility(),
                  new NoOpDestListMetadataReader(),
                  new NoOpQuickAccessRestoreEngine())
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
                  new NoOpRecentLinksCleaner(),
                  new NoOpQuickAccessLockFactory(),
                  new NoOpQuickAccessVisibility(),
                  new NoOpDestListMetadataReader(),
                  new NoOpQuickAccessRestoreEngine())
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
                  new NoOpRecentLinksCleaner(),
                  new NoOpQuickAccessLockFactory(),
                  new NoOpQuickAccessVisibility(),
                  new NoOpDestListMetadataReader(),
                  new NoOpQuickAccessRestoreEngine())
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
                  new NoOpRecentLinksCleaner(),
                  new NoOpQuickAccessLockFactory(),
                  new NoOpQuickAccessVisibility(),
                  new NoOpDestListMetadataReader(),
                  new NoOpQuickAccessRestoreEngine())
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
                  new NoOpRecentLinksCleaner(),
                  new NoOpQuickAccessLockFactory(),
                  new NoOpQuickAccessVisibility(),
                  new NoOpDestListMetadataReader(),
                  new NoOpQuickAccessRestoreEngine())
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
                  recentLinksCleaner,
                  new NoOpQuickAccessLockFactory(),
                  new NoOpQuickAccessVisibility(),
                  new NoOpDestListMetadataReader(),
                  new NoOpQuickAccessRestoreEngine())
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
            IRecentLinksCleaner recentLinksCleaner,
            IQuickAccessLockFactory lockFactory)
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
                  recentLinksCleaner,
                  lockFactory,
                  new NoOpQuickAccessVisibility(),
                  new NoOpDestListMetadataReader(),
                  new NoOpQuickAccessRestoreEngine())
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
            IRecentLinksCleaner recentLinksCleaner,
            IQuickAccessLockFactory lockFactory,
            IQuickAccessVisibility visibility)
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
                  recentLinksCleaner,
                  lockFactory,
                  visibility,
                  new NoOpDestListMetadataReader(),
                  new NoOpQuickAccessRestoreEngine())
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
            IRecentLinksCleaner recentLinksCleaner,
            IQuickAccessLockFactory lockFactory,
            IQuickAccessVisibility visibility,
            IDestListMetadataReader destListReader)
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
                  recentLinksCleaner,
                  lockFactory,
                  visibility,
                  destListReader,
                  new NoOpQuickAccessRestoreEngine())
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
            IRecentLinksCleaner recentLinksCleaner,
            IQuickAccessLockFactory lockFactory,
            IQuickAccessVisibility visibility,
            IDestListMetadataReader destListReader,
            IQuickAccessRestoreEngine restoreEngine)
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
            _lockFactory = lockFactory ?? throw new ArgumentNullException(nameof(lockFactory));
            _visibility = visibility ?? throw new ArgumentNullException(nameof(visibility));
            _destListReader = destListReader ?? throw new ArgumentNullException(nameof(destListReader));
            _restoreEngine = restoreEngine ?? throw new ArgumentNullException(nameof(restoreEngine));
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
            return ExecuteWithRetry(() =>
            {
                try
                {
                    return _nativeQuery.GetItems(target, _timeout);
                }
                catch (Exception)
                {
                    if (target == QuickAccess.All)
                        return QuickAccessQueryMerger.MergeRecentAndFrequent(
                            ExecuteListScript(PSScript.QueryRecentFile, null, ToTimeoutSeconds()),
                            ExecuteListScript(PSScript.QueryFrequentFolder, null, ToTimeoutSeconds()));

                    var script = MapQueryScript(target);
                    return ExecuteListScript(script, null, ToTimeoutSeconds());
                }
            });
        }

        /// <summary>
        /// Retrieves Quick Access paths as path objects.
        /// </summary>
        /// <param name="target">The section to query.</param>
        /// <returns>The current item paths.</returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="target"/> is not supported.</exception>
        /// <exception cref="PowerShellExecutionException">The PowerShell fallback query fails.</exception>
        /// <remarks>
        /// The returned paths wrap Explorer's current path strings and may point to files or folders that no longer
        /// exist.
        /// </remarks>
        public IReadOnlyList<QuickAccessPath> GetItemPaths(QuickAccess target)
        {
            return GetItems(target)
                .Select(path => new QuickAccessPath(path))
                .ToList()
                .AsReadOnly();
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
        /// <exception cref="QuickAccessPostMutationException">The add succeeds but a requested post-mutation step fails.</exception>
        /// <exception cref="UnsupportedQuickAccessOperationException"><paramref name="target"/> is <see cref="QuickAccess.All"/>.</exception>
        /// <remarks>
        /// This method modifies the current Windows user's Quick Access state. When
        /// <see cref="AddOptions.ForceRecentFilesRebuild"/> is enabled for recent files, the current Recent Files
        /// backing data is removed to force Explorer to rebuild it. Frequent folders are pinned with a native Shell verb
        /// first and fall back to PowerShell only when the native Shell operation fails.
        /// </remarks>
        public void AddItem(string path, QuickAccess target, AddOptions options)
        {
            if (path == null)
                throw new ArgumentNullException(nameof(path));
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            EnsureSingleItemTarget(target, "AddItem");

            bool added = false;
            try
            {
                if (target == QuickAccess.RecentFiles)
                {
                    ValidatePath(path, PathType.File, _fileSystem);
                    if (ContainsItemExact(path, target))
                        throw new QuickAccessItemAlreadyExistsException(path, target);

                    AddFileToRecentDocs(path);
                    added = true;
                    if (options.RefreshRecentFiles)
                        DeleteRecentFilesBackingDataAfterMutation(path);
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
                    added = true;
                }

                if (options.RefreshExplorer)
                    RefreshExplorerAfterMutation(path, target);
            }
            finally
            {
                if (added)
                    _executor.ClearCache();
            }

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
        /// <exception cref="QuickAccessPostMutationException">The remove succeeds but a requested post-mutation step fails.</exception>
        /// <exception cref="UnsupportedQuickAccessOperationException"><paramref name="target"/> is <see cref="QuickAccess.All"/>.</exception>
        /// <remarks>
        /// This method modifies the current Windows user's Quick Access state. Removal uses native Shell verbs first
        /// and falls back to PowerShell only for native Shell failures. When
        /// <see cref="RemoveOptions.DeepCleanRecentLinks"/> is enabled, matching shortcut files in the Windows Recent
        /// folder are deleted after the Shell removal succeeds for both Recent Files and Frequent Folders.
        /// </remarks>
        public void RemoveItem(string path, QuickAccess target, RemoveOptions options)
        {
            try
            {
                RemoveItemCore(path, target, options);
            }
            catch (ShellMutationSucceededException ex)
            {
                throw ex.InnerException;
            }
        }

        private bool RemoveItemCore(string path, QuickAccess target, RemoveOptions options)
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
                if (options.RefreshExplorer)
                    RefreshExplorerAfterMutation(path, target);
                if (options.DeepCleanRecentLinks)
                {
                    try
                    {
                        _recentLinksCleaner.DeleteForTarget(path, _timeout);
                    }
                    catch (Exception ex)
                    {
                        throw new ShellMutationSucceededException(ex);
                    }
                }

                return true;
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
            int refreshExplorerItemIndex = -1;

            foreach (var item in items)
            {
                if (item == null)
                    throw new ArgumentException("Batch items cannot contain null.", nameof(items));

                try
                {
                    AddItem(item.Path, item.Target, new AddOptions());
                    succeeded.Add(item);
                    int index = succeeded.Count - 1;
                    if (options.RefreshRecentFiles && item.Target == QuickAccess.RecentFiles)
                        refreshRecentItemIndex = index;
                    if (options.RefreshExplorer)
                        refreshExplorerItemIndex = SelectBatchRefreshItemIndex(
                            succeeded,
                            refreshExplorerItemIndex,
                            item.Target,
                            preferRecentFiles: true);
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
                    var item = succeeded[refreshRecentItemIndex];
                    DeleteRecentFilesBackingDataAfterMutation(item.Path);
                }
                catch (Exception ex)
                {
                    var item = succeeded[refreshRecentItemIndex];
                    succeeded.RemoveAt(refreshRecentItemIndex);
                    if (refreshExplorerItemIndex == refreshRecentItemIndex)
                        refreshExplorerItemIndex = SelectLastSucceededIndex(succeeded);
                    else if (refreshExplorerItemIndex > refreshRecentItemIndex)
                        refreshExplorerItemIndex--;

                    failed.Add(new BatchFailure(item, ex));
                }
            }

            RecordBatchExplorerRefreshFailure(options.RefreshExplorer, succeeded, failed, refreshExplorerItemIndex);

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
            return RemoveItems(items, new BatchOptions(), options);
        }

        /// <summary>
        /// Removes multiple Quick Access items.
        /// </summary>
        /// <param name="items">The items to remove.</param>
        /// <param name="batchOptions">The batch options.</param>
        /// <param name="removeOptions">The remove options applied to each item.</param>
        /// <returns>The batch result.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="items"/>, <paramref name="batchOptions"/>, or <paramref name="removeOptions"/> is <see langword="null"/>.</exception>
        public BatchResult RemoveItems(IEnumerable<QuickAccessItem> items, BatchOptions batchOptions, RemoveOptions removeOptions)
        {
            if (items == null)
                throw new ArgumentNullException(nameof(items));
            if (batchOptions == null)
                throw new ArgumentNullException(nameof(batchOptions));
            if (removeOptions == null)
                throw new ArgumentNullException(nameof(removeOptions));

            var succeeded = new List<QuickAccessItem>();
            var failed = new List<BatchFailure>();
            int refreshExplorerItemIndex = -1;
            bool refreshExplorerWithoutAttribution = false;
            var perItemOptions = new RemoveOptions
            {
                DeepCleanRecentLinks = removeOptions.DeepCleanRecentLinks,
                RefreshExplorer = false
            };

            foreach (var item in items)
            {
                if (item == null)
                    throw new ArgumentException("Batch items cannot contain null.", nameof(items));

                try
                {
                    bool mutationSucceeded = RemoveItemCore(item.Path, item.Target, perItemOptions);
                    if (!mutationSucceeded)
                        continue;

                    succeeded.Add(item);
                    if (batchOptions.RefreshExplorer)
                        refreshExplorerItemIndex = SelectBatchRefreshItemIndex(
                            succeeded,
                            refreshExplorerItemIndex,
                            item.Target,
                            preferRecentFiles: true);
                }
                catch (Exception ex)
                {
                    var mutationSucceeded = ex as ShellMutationSucceededException;
                    if (mutationSucceeded != null)
                        refreshExplorerWithoutAttribution = true;

                    failed.Add(new BatchFailure(item, mutationSucceeded?.InnerException ?? ex));
                }
            }

            RecordBatchExplorerRefreshFailure(
                batchOptions.RefreshExplorer,
                succeeded,
                failed,
                refreshExplorerItemIndex,
                refreshExplorerWithoutAttribution);

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
            ValidateClearOptions(options);

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
                    ClearFrequentFolders(options.RemovePinnedFolders, options.PinnedFoldersTimeout ?? _timeout);
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
        /// Restores a Quick Access section to the system-default state.
        /// </summary>
        /// <param name="target">The section to restore.</param>
        /// <returns>The restore report.</returns>
        public RestoreDefaultsReport RestoreDefaults(QuickAccess target)
        {
            return RestoreDefaults(target, new RestoreDefaultsOptions());
        }

        /// <summary>
        /// Restores a Quick Access section to the system-default state.
        /// </summary>
        /// <param name="target">The section to restore.</param>
        /// <param name="options">The restore options.</param>
        /// <returns>The restore report.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="options"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="target"/> or an option value is invalid.</exception>
        public RestoreDefaultsReport RestoreDefaults(QuickAccess target, RestoreDefaultsOptions options)
        {
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            EnsureClearTarget(target);
            options.Validate();
            var report = _restoreEngine.RestoreDefaults(target, options, ClearRecentFiles, RefreshExplorer);
            if (report.Success || report.RecentReport != null || report.FrequentReport != null)
                _executor.ClearCache();

            return report;
        }

        /// <summary>
        /// Restores Recent Files to the system-default state.
        /// </summary>
        /// <returns>The restore report.</returns>
        public RecentRestoreReport RestoreRecentFilesDefaults()
        {
            return RestoreRecentFilesDefaults(new RestoreDefaultsOptions());
        }

        /// <summary>
        /// Restores Recent Files to the system-default state.
        /// </summary>
        /// <param name="options">The restore options.</param>
        /// <returns>The restore report.</returns>
        public RecentRestoreReport RestoreRecentFilesDefaults(RestoreDefaultsOptions options)
        {
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            options.Validate();
            var report = _restoreEngine.RestoreRecentFilesDefaults(options, ClearRecentFiles);
            _executor.ClearCache();
            return report;
        }

        /// <summary>
        /// Restores Frequent Folders to the system-default state.
        /// </summary>
        /// <returns>The restore report.</returns>
        public FrequentRestoreReport RestoreFrequentFoldersDefaults()
        {
            return RestoreFrequentFoldersDefaults(new RestoreDefaultsOptions());
        }

        /// <summary>
        /// Restores Frequent Folders to the system-default state.
        /// </summary>
        /// <param name="options">The restore options.</param>
        /// <returns>The restore report.</returns>
        public FrequentRestoreReport RestoreFrequentFoldersDefaults(RestoreDefaultsOptions options)
        {
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            options.Validate();
            var report = _restoreEngine.RestoreFrequentFoldersDefaults(options, RefreshExplorer);
            _executor.ClearCache();
            return report;
        }

        /// <summary>
        /// Locks both Quick Access backing files.
        /// </summary>
        /// <returns>A lock that must be released by calling <see cref="QuickAccessLock.Unlock()"/> or <see cref="QuickAccessLock.Dispose"/>.</returns>
        /// <remarks>
        /// This method opens the current Windows user's Recent Files and Frequent Folders backing files and keeps the
        /// file handles alive until the returned lock is released.
        /// </remarks>
        public QuickAccessLock LockQuickAccess()
        {
            return _lockFactory.Lock(QuickAccessLockTarget.All);
        }

        /// <summary>
        /// Locks the Recent Files backing file.
        /// </summary>
        /// <returns>A lock that must be released by calling <see cref="QuickAccessLock.Unlock()"/> or <see cref="QuickAccessLock.Dispose"/>.</returns>
        /// <remarks>
        /// The returned lock also records a Windows Recent shortcut snapshot for unlock reports and optional cleanup.
        /// </remarks>
        public QuickAccessLock LockRecentFiles()
        {
            return _lockFactory.Lock(QuickAccessLockTarget.RecentFiles);
        }

        /// <summary>
        /// Locks the Frequent Folders backing file.
        /// </summary>
        /// <returns>A lock that must be released by calling <see cref="QuickAccessLock.Unlock()"/> or <see cref="QuickAccessLock.Dispose"/>.</returns>
        /// <remarks>
        /// The returned lock records a Windows Recent shortcut snapshot for unlock reports even though the backing file
        /// being locked is the Frequent Folders data file.
        /// </remarks>
        public QuickAccessLock LockFrequentFolders()
        {
            return _lockFactory.Lock(QuickAccessLockTarget.FrequentFolders);
        }

        /// <summary>
        /// Gets whether a Quick Access section is visible in Explorer.
        /// </summary>
        /// <param name="target">The section to inspect.</param>
        /// <returns><see langword="true"/> when the requested section is visible.</returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="target"/> is not supported.</exception>
        /// <remarks>
        /// This method reads the current Windows user's Explorer registry settings. <see cref="QuickAccess.RecentFiles"/>
        /// maps to Explorer's <c>ShowRecent</c> value, and <see cref="QuickAccess.FrequentFolders"/> maps to
        /// <c>ShowFrequent</c>. Missing registry values are treated as visible. Explorer's frequent-folder visibility
        /// setting only controls automatically shown, unpinned frequent folders; pinned Quick Access folders remain
        /// controlled by Explorer's pin state.
        /// </remarks>
        /// <seealso cref="SetVisible(QuickAccess, bool)"/>
        public bool IsVisible(QuickAccess target)
        {
            return _visibility.IsVisible(target);
        }

        /// <summary>
        /// Sets whether a Quick Access section is visible in Explorer.
        /// </summary>
        /// <param name="target">The section to update.</param>
        /// <param name="visible">Whether the section should be visible.</param>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="target"/> is not supported.</exception>
        /// <remarks>
        /// This method writes the current Windows user's Explorer registry settings and does not refresh Explorer
        /// windows. Use <see cref="SetVisible(QuickAccess, bool, VisibilityOptions)"/> with
        /// <see cref="VisibilityOptions.RefreshExplorer"/> to request a Shell refresh. For
        /// <see cref="QuickAccess.FrequentFolders"/>, this setting only affects automatically shown, unpinned frequent
        /// folders; pinned Quick Access folders are not hidden by <c>ShowFrequent</c>.
        /// </remarks>
        /// <seealso cref="IsVisible(QuickAccess)"/>
        public void SetVisible(QuickAccess target, bool visible)
        {
            SetVisible(target, visible, new VisibilityOptions());
        }

        /// <summary>
        /// Sets whether a Quick Access section is visible in Explorer.
        /// </summary>
        /// <param name="target">The section to update.</param>
        /// <param name="visible">Whether the section should be visible.</param>
        /// <param name="options">The visibility options.</param>
        /// <exception cref="ArgumentNullException"><paramref name="options"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="target"/> is not supported.</exception>
        /// <remarks>
        /// This method writes the current Windows user's Explorer registry settings. When
        /// <see cref="VisibilityOptions.RefreshExplorer"/> is enabled, Explorer windows are refreshed after the registry
        /// write. If the refresh fails after both native Shell refresh and PowerShell fallback, the refresh failure is
        /// propagated and the registry write is not rolled back.
        /// </remarks>
        /// <seealso cref="IsVisible(QuickAccess)"/>
        public void SetVisible(QuickAccess target, bool visible, VisibilityOptions options)
        {
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            _visibility.SetVisible(target, visible);
            if (options.RefreshExplorer)
                RefreshExplorer();
        }

        /// <summary>
        /// Shows a Quick Access section in Explorer.
        /// </summary>
        /// <param name="target">The section to show.</param>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="target"/> is not supported.</exception>
        /// <remarks>
        /// This method writes the current Windows user's Explorer registry settings and does not refresh Explorer
        /// windows. Use <see cref="ShowSection(QuickAccess, VisibilityOptions)"/> to request a Shell refresh.
        /// </remarks>
        /// <seealso cref="SetVisible(QuickAccess, bool)"/>
        public void ShowSection(QuickAccess target)
        {
            ShowSection(target, new VisibilityOptions());
        }

        /// <summary>
        /// Shows a Quick Access section in Explorer.
        /// </summary>
        /// <param name="target">The section to show.</param>
        /// <param name="options">The visibility options.</param>
        /// <exception cref="ArgumentNullException"><paramref name="options"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="target"/> is not supported.</exception>
        /// <remarks>
        /// This is a convenience wrapper for <see cref="SetVisible(QuickAccess, bool, VisibilityOptions)"/>.
        /// </remarks>
        /// <seealso cref="SetVisible(QuickAccess, bool, VisibilityOptions)"/>
        public void ShowSection(QuickAccess target, VisibilityOptions options)
        {
            SetVisible(target, true, options);
        }

        /// <summary>
        /// Hides a Quick Access section in Explorer.
        /// </summary>
        /// <param name="target">The section to hide.</param>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="target"/> is not supported.</exception>
        /// <remarks>
        /// This method writes the current Windows user's Explorer registry settings and does not refresh Explorer
        /// windows. Use <see cref="HideSection(QuickAccess, VisibilityOptions)"/> to request a Shell refresh. For
        /// <see cref="QuickAccess.FrequentFolders"/>, pinned Quick Access folders remain visible because Explorer
        /// treats pinned folders separately from automatically shown frequent folders.
        /// </remarks>
        /// <seealso cref="SetVisible(QuickAccess, bool)"/>
        public void HideSection(QuickAccess target)
        {
            HideSection(target, new VisibilityOptions());
        }

        /// <summary>
        /// Hides a Quick Access section in Explorer.
        /// </summary>
        /// <param name="target">The section to hide.</param>
        /// <param name="options">The visibility options.</param>
        /// <exception cref="ArgumentNullException"><paramref name="options"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="target"/> is not supported.</exception>
        /// <remarks>
        /// This is a convenience wrapper for <see cref="SetVisible(QuickAccess, bool, VisibilityOptions)"/>.
        /// </remarks>
        /// <seealso cref="SetVisible(QuickAccess, bool, VisibilityOptions)"/>
        public void HideSection(QuickAccess target, VisibilityOptions options)
        {
            SetVisible(target, false, options);
        }

        /// <summary>
        /// Parses Recent Files DestList metadata.
        /// </summary>
        /// <returns>The Recent Files DestList entries.</returns>
        /// <exception cref="IOException">The backing file cannot be read.</exception>
        /// <exception cref="DestListParseException">The backing file is malformed or truncated.</exception>
        /// <exception cref="DestListUnsupportedVersionException">The DestList version is unsupported.</exception>
        /// <remarks>
        /// This method reads Explorer's `.automaticDestinations-ms` backing file for the current Windows user.
        /// </remarks>
        public IReadOnlyList<DestListEntry> GetRecentFilesMetadata()
        {
            return _destListReader.ParseFile(_dataFiles.RecentFilesPath).DestList.Entries;
        }

        /// <summary>
        /// Parses Frequent Folders DestList metadata.
        /// </summary>
        /// <returns>The Frequent Folders DestList entries.</returns>
        /// <exception cref="IOException">The backing file cannot be read.</exception>
        /// <exception cref="DestListParseException">The backing file is malformed or truncated.</exception>
        /// <exception cref="DestListUnsupportedVersionException">The DestList version is unsupported.</exception>
        /// <remarks>
        /// This method reads Explorer's `.automaticDestinations-ms` backing file for the current Windows user.
        /// </remarks>
        public IReadOnlyList<DestListEntry> GetFrequentFoldersMetadata()
        {
            return _destListReader.ParseFile(_dataFiles.FrequentFoldersPath).DestList.Entries;
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

        private static void ValidateClearOptions(ClearOptions options)
        {
            if (options.RemovePinnedFolders &&
                options.PinnedFoldersTimeout.HasValue &&
                options.PinnedFoldersTimeout.Value <= TimeSpan.Zero)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(options),
                    "PinnedFoldersTimeout must be greater than zero when RemovePinnedFolders is enabled.");
            }
        }

        private static int SelectBatchRefreshItemIndex(
            IReadOnlyList<QuickAccessItem> succeeded,
            int currentIndex,
            QuickAccess target,
            bool preferRecentFiles)
        {
            if (succeeded == null || succeeded.Count == 0)
                return -1;

            if (!preferRecentFiles)
                return currentIndex < 0 ? succeeded.Count - 1 : currentIndex;

            if (target == QuickAccess.RecentFiles)
                return succeeded.Count - 1;

            return currentIndex < 0 ? succeeded.Count - 1 : currentIndex;
        }

        private static int SelectLastSucceededIndex(IReadOnlyList<QuickAccessItem> succeeded)
        {
            return succeeded == null || succeeded.Count == 0 ? -1 : succeeded.Count - 1;
        }

        private void RecordBatchExplorerRefreshFailure(
            bool shouldRefresh,
            List<QuickAccessItem> succeeded,
            List<BatchFailure> failed,
            int refreshExplorerItemIndex,
            bool attemptedRefreshWithoutAttribution = false)
        {
            if (!shouldRefresh)
                return;

            if (succeeded.Count == 0 || refreshExplorerItemIndex < 0)
            {
                if (attemptedRefreshWithoutAttribution)
                {
                    try
                    {
                        RefreshExplorer();
                    }
                    catch (Exception)
                    {
                        // The mutated item is already reported as failed by its post-mutation cleanup.
                    }
                }

                return;
            }

            try
            {
                RefreshExplorer();
            }
            catch (Exception ex)
            {
                if (refreshExplorerItemIndex >= succeeded.Count)
                    refreshExplorerItemIndex = succeeded.Count - 1;

                var item = succeeded[refreshExplorerItemIndex];
                succeeded.RemoveAt(refreshExplorerItemIndex);
                failed.Add(new BatchFailure(
                    item,
                    CreatePostMutationException(
                        item.Path,
                        item.Target,
                        QuickAccessPostMutationStep.RefreshExplorer,
                        ex)));
            }
        }

        private sealed class ShellMutationSucceededException : Exception
        {
            public ShellMutationSucceededException(Exception innerException)
                : base(innerException?.Message, innerException)
            {
            }
        }

        private void DeleteRecentFilesBackingDataAfterMutation(string path)
        {
            try
            {
                _dataFiles.RemoveRecentFile();
            }
            catch (Exception ex)
            {
                throw CreatePostMutationException(
                    path,
                    QuickAccess.RecentFiles,
                    QuickAccessPostMutationStep.DeleteRecentFilesBackingData,
                    ex);
            }
        }

        private void RefreshExplorerAfterMutation(string path, QuickAccess target)
        {
            try
            {
                RefreshExplorer();
            }
            catch (Exception ex)
            {
                throw CreatePostMutationException(path, target, QuickAccessPostMutationStep.RefreshExplorer, ex);
            }
        }

        private static QuickAccessPostMutationException CreatePostMutationException(
            string path,
            QuickAccess target,
            QuickAccessPostMutationStep step,
            Exception innerException)
        {
            return new QuickAccessPostMutationException(path, target, step, innerException);
        }

        private IReadOnlyList<string> ExecuteListScript(PSScript script, string parameter, int timeoutSeconds)
        {
            return _executor.ExecutePSScriptWithCache(script, parameter, timeoutSeconds)
                .GetAwaiter()
                .GetResult()
                .AsReadOnly();
        }

        private void ExecuteMutationScript(PSScript script, string parameter, PowerShellOperation operation)
        {
            ExecuteMutationScript(script, parameter, operation, _timeout);
        }

        private void ExecuteMutationScript(PSScript script, string parameter, PowerShellOperation operation, TimeSpan timeout)
        {
            _executor.ExecutePSScriptWithTimeout(script, parameter, ToTimeoutSeconds(timeout))
                .GetAwaiter()
                .GetResult();
        }

        private void ExecuteNativeMutationWithPowerShellFallback(
            Action nativeAction,
            PSScript fallbackScript,
            string parameter,
            PowerShellOperation operation)
        {
            ExecuteWithRetry(
                () =>
                {
                    try
                    {
                        nativeAction();
                    }
                    catch (Exception ex) when (ShouldFallbackToPowerShell(ex))
                    {
                        ExecuteMutationScript(fallbackScript, parameter, operation);
                    }

                    return true;
                });
        }

        private bool ShouldRetry(Exception exception, int retryAttempt)
        {
            if (_retryPolicy.MaxRetryCount == 0 || retryAttempt >= _retryPolicy.MaxRetryCount)
                return false;

            if (exception is PowerShellExecutionException powerShellException)
                return ShouldRetry(powerShellException);

            return IsRetryableShellFailure(exception);
        }

        private static bool ShouldRetry(PowerShellExecutionException exception)
        {
            if (exception.Kind == PowerShellErrorKind.Timeout)
                return true;

            string stderr = exception.StandardError ?? string.Empty;
            return stderr.IndexOf("locked", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   stderr.IndexOf("temporarily unavailable", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsRetryableShellFailure(Exception exception)
        {
            if (exception is TimeoutException)
                return false;

            if (exception is COMException)
                return true;

            var invalidOperation = exception as InvalidOperationException;
            return invalidOperation != null &&
                   IsRecoverableNativeShellFailure(invalidOperation) &&
                   !IsNativeDisabledFallbackMarker(invalidOperation);
        }

        private T ExecuteWithRetry<T>(Func<T> action)
        {
            for (int retryAttempt = 0; ; retryAttempt++)
            {
                try
                {
                    return action();
                }
                catch (Exception ex) when (ShouldRetry(ex, retryAttempt))
                {
                    System.Threading.Thread.Sleep(_retryPolicy.GetDelay(retryAttempt));
                }
            }
        }

        private static bool ShouldFallbackToPowerShell(Exception exception)
        {
            if (exception is TimeoutException)
                return false;

            if (exception is COMException ||
                exception is QuickAccessOperationException)
                return true;

            var invalidOperation = exception as InvalidOperationException;
            return invalidOperation != null && IsRecoverableNativeShellFailure(invalidOperation);
        }

        private static bool IsRecoverableNativeShellFailure(InvalidOperationException exception)
        {
            string message = exception.Message ?? string.Empty;
            return message.StartsWith("Failed to open shell namespace:", StringComparison.Ordinal) ||
                   message.StartsWith("Failed to enumerate shell folder items.", StringComparison.Ordinal) ||
                   message.StartsWith("Failed to open shell folder:", StringComparison.Ordinal) ||
                   message.StartsWith("Failed to get shell folder self item:", StringComparison.Ordinal) ||
                   message == "Shell.Application COM object is not available." ||
                   message == "Native Quick Access query is disabled for this instance." ||
                   message == "Native Quick Access mutation is disabled for this instance.";
        }

        private static bool IsNativeDisabledFallbackMarker(InvalidOperationException exception)
        {
            string message = exception.Message ?? string.Empty;
            return message == "Native Quick Access query is disabled for this instance." ||
                   message == "Native Quick Access mutation is disabled for this instance.";
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
            return ToTimeoutSeconds(_timeout);
        }

        private static int ToTimeoutSeconds(TimeSpan timeout)
        {
            return Math.Max(1, (int)Math.Ceiling(timeout.TotalSeconds));
        }

        private void AddFileToRecentDocs(string filePath)
        {
            StaThreadRunner.Run(
                () =>
                {
                    IntPtr pathPtr = IntPtr.Zero;

                    try
                    {
                        pathPtr = Marshal.StringToHGlobalUni(filePath);
                        _nativeMethods.SHAddToRecentDocs(NativeMethods.SHARD_PATHW, pathPtr);
                    }
                    finally
                    {
                        if (pathPtr != IntPtr.Zero)
                            Marshal.FreeHGlobal(pathPtr);
                    }
                },
                _timeout,
                _nativeMethods,
                disableOle1Dde: true);
        }

        private void ClearRecentFiles()
        {
            ClearRecentFiles(_timeout);
        }

        private void ClearRecentFiles(TimeSpan timeout)
        {
            StaThreadRunner.Run(
                () => _nativeMethods.SHAddToRecentDocs(NativeMethods.SHARD_PATHW, IntPtr.Zero),
                timeout,
                _nativeMethods,
                disableOle1Dde: true);
        }

        private void ClearFrequentFolders(bool removePinnedFolders, TimeSpan pinnedFoldersTimeout)
        {
            var jumpListFile = _dataFiles.FrequentFoldersPath;
            IReadOnlyList<string> pinnedSnapshot = removePinnedFolders
                ? _nativeQuery.GetItems(QuickAccess.FrequentFolders, pinnedFoldersTimeout)
                : Array.Empty<string>();

            if (_fileSystem.FileExists(jumpListFile))
                _fileSystem.DeleteFile(jumpListFile);

            if (removePinnedFolders)
            {
                try
                {
                    ClearPinnedFoldersFromSnapshot(pinnedSnapshot, pinnedFoldersTimeout);
                }
                catch (Exception ex)
                {
                    throw new PartialClearException(false, true, ex);
                }
            }
        }

        private void ClearPinnedFoldersFromSnapshot(IReadOnlyList<string> pinnedSnapshot, TimeSpan timeout)
        {
            Exception firstError = null;
            foreach (string path in pinnedSnapshot)
            {
                try
                {
                    _nativeMutation.UnpinFrequentFolder(path, timeout);
                }
                catch (QuickAccessItemNotFoundException)
                {
                }
                catch (Exception ex)
                {
                    firstError = firstError ?? ex;
                }
            }

            if (firstError != null)
                throw firstError;
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
            RefreshExplorer(_timeout);
        }

        private void RefreshExplorer(TimeSpan timeout)
        {
            try
            {
                _explorerRefresher.Refresh(timeout);
            }
            catch (Exception)
            {
                ExecuteMutationScript(PSScript.RefreshExplorer, null, PowerShellOperation.RefreshExplorer, timeout);
            }
        }
    }
}
