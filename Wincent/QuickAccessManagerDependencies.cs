using System;

namespace Wincent
{
    internal sealed class QuickAccessManagerDependencies
    {
        public IScriptExecutor Executor { get; set; }
        public TimeSpan Timeout { get; set; }
        public RetryPolicy RetryPolicy { get; set; }
        public IFileSystemOperations FileSystem { get; set; }
        public INativeMethods NativeMethods { get; set; }
        public IQuickAccessDataFiles DataFiles { get; set; }
        public IQuickAccessNativeQuery NativeQuery { get; set; }
        public IQuickAccessNativeMutation NativeMutation { get; set; }
        public IExplorerRefresher ExplorerRefresher { get; set; }
        public IRecentLinksCleaner RecentLinksCleaner { get; set; }
        public IQuickAccessLockFactory LockFactory { get; set; }
        public IQuickAccessVisibility Visibility { get; set; }
        public IDestListMetadataReader DestListReader { get; set; }
        public IQuickAccessRestoreEngine RestoreEngine { get; set; }

        public static QuickAccessManagerDependencies CreateDefault(QuickAccessManagerOptions options)
        {
            if (options == null)
                throw new ArgumentNullException(nameof(options));
            if (options.Timeout <= TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(options), "Timeout must be positive.");

            var nativeMethods = new DefaultNativeMethods();
            var dataFiles = new QuickAccessDataFiles();
            var recentFolder = new WindowsRecentFolder(nativeMethods);
            var recentLinkFileSystem = new DefaultRecentLinkFileSystem();
            var shortcutResolver = new ShellLinkTargetResolver(nativeMethods);
            var destListReader = new DefaultDestListMetadataReader();

            return new QuickAccessManagerDependencies
            {
                Executor = new ScriptExecutor(),
                Timeout = options.Timeout,
                RetryPolicy = options.RetryPolicy ?? RetryPolicy.Standard,
                FileSystem = new DefaultFileSystemOperations(),
                NativeMethods = nativeMethods,
                DataFiles = dataFiles,
                NativeQuery = new ShellQuickAccessNativeQuery(nativeMethods),
                NativeMutation = new ShellQuickAccessNativeMutation(nativeMethods),
                ExplorerRefresher = new ShellExplorerRefresher(nativeMethods),
                RecentLinksCleaner = new RecentLinksCleaner(
                    recentFolder,
                    shortcutResolver,
                    recentLinkFileSystem),
                LockFactory = new QuickAccessLockFactory(
                    dataFiles,
                    recentFolder,
                    recentLinkFileSystem,
                    new NativeQuickAccessBackingFileHandleOpener()),
                Visibility = new RegistryQuickAccessVisibility(new CurrentUserExplorerVisibilityRegistry()),
                DestListReader = destListReader,
                RestoreEngine = new QuickAccessRestoreEngine(
                    dataFiles,
                    recentFolder,
                    shortcutResolver,
                    recentLinkFileSystem,
                    new DefaultFileSystemOperations(),
                    destListReader,
                    new DefaultQuickAccessRestoreDelay())
            };
        }

        public static QuickAccessManagerDependencies CreateTestingDefaults(
            IScriptExecutor executor,
            TimeSpan timeout,
            IFileSystemOperations fileSystem,
            INativeMethods nativeMethods,
            IQuickAccessDataFiles dataFiles)
        {
            return new QuickAccessManagerDependencies
            {
                Executor = executor,
                Timeout = timeout,
                RetryPolicy = RetryPolicy.Standard,
                FileSystem = fileSystem,
                NativeMethods = nativeMethods,
                DataFiles = dataFiles,
                NativeQuery = new PowerShellFallbackNativeQuery(),
                NativeMutation = new PowerShellFallbackNativeMutation(),
                ExplorerRefresher = new PowerShellFallbackExplorerRefresher(),
                RecentLinksCleaner = new NoOpRecentLinksCleaner(),
                LockFactory = new NoOpQuickAccessLockFactory(),
                Visibility = new NoOpQuickAccessVisibility(),
                DestListReader = new NoOpDestListMetadataReader(),
                RestoreEngine = new NoOpQuickAccessRestoreEngine()
            };
        }
    }
}
