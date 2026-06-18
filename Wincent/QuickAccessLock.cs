using System;
using System.Collections.Generic;
using System.Linq;

namespace Wincent
{
    internal interface IQuickAccessLockFactory
    {
        QuickAccessLock Lock(QuickAccessLockTarget target);
    }

    internal interface IQuickAccessBackingFileHandle : IDisposable
    {
    }

    internal interface IQuickAccessBackingFileHandleOpener
    {
        IQuickAccessBackingFileHandle Open(string path);
    }

    internal sealed class NativeQuickAccessBackingFileHandleOpener : IQuickAccessBackingFileHandleOpener
    {
        public IQuickAccessBackingFileHandle Open(string path)
        {
            return new NativeQuickAccessBackingFileHandle(NativeFileHandle.OpenExistingForBackingFileLock(path));
        }

        private sealed class NativeQuickAccessBackingFileHandle : IQuickAccessBackingFileHandle
        {
            private readonly NativeFileHandle _handle;

            public NativeQuickAccessBackingFileHandle(NativeFileHandle handle)
            {
                _handle = handle ?? throw new ArgumentNullException(nameof(handle));
            }

            public void Dispose()
            {
                _handle.Dispose();
            }
        }
    }

    internal sealed class QuickAccessLockFactory : IQuickAccessLockFactory
    {
        private readonly IQuickAccessDataFiles _dataFiles;
        private readonly IWindowsRecentFolder _recentFolder;
        private readonly IRecentLinkFileSystem _fileSystem;
        private readonly IQuickAccessBackingFileHandleOpener _handleOpener;

        public QuickAccessLockFactory(
            IQuickAccessDataFiles dataFiles,
            IWindowsRecentFolder recentFolder,
            IRecentLinkFileSystem fileSystem,
            IQuickAccessBackingFileHandleOpener handleOpener)
        {
            _dataFiles = dataFiles ?? throw new ArgumentNullException(nameof(dataFiles));
            _recentFolder = recentFolder ?? throw new ArgumentNullException(nameof(recentFolder));
            _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));
            _handleOpener = handleOpener ?? throw new ArgumentNullException(nameof(handleOpener));
        }

        public QuickAccessLock Lock(QuickAccessLockTarget target)
        {
            EnsureSupportedTarget(target);

            var handles = new List<IQuickAccessBackingFileHandle>();
            try
            {
                if (target == QuickAccessLockTarget.RecentFiles || target == QuickAccessLockTarget.All)
                    handles.Add(_handleOpener.Open(_dataFiles.RecentFilesPath));

                if (target == QuickAccessLockTarget.FrequentFolders || target == QuickAccessLockTarget.All)
                    handles.Add(_handleOpener.Open(_dataFiles.FrequentFoldersPath));

                string recentFolder = _recentFolder.GetPath();
                var initialShortcutPaths = QuickAccessShortcutSnapshot.EnumerateShortcutPaths(_fileSystem, recentFolder);
                return new QuickAccessLock(target, recentFolder, initialShortcutPaths, handles, _fileSystem);
            }
            catch
            {
                foreach (var handle in handles)
                    handle.Dispose();

                throw;
            }
        }

        private static void EnsureSupportedTarget(QuickAccessLockTarget target)
        {
            if (target != QuickAccessLockTarget.RecentFiles &&
                target != QuickAccessLockTarget.FrequentFolders &&
                target != QuickAccessLockTarget.All)
            {
                throw new ArgumentOutOfRangeException(nameof(target), target, "Unsupported Quick Access lock target.");
            }
        }
    }

    internal static class QuickAccessShortcutSnapshot
    {
        public static IReadOnlyList<string> EnumerateShortcutPaths(IRecentLinkFileSystem fileSystem, string recentFolder)
        {
            return fileSystem.EnumerateFiles(recentFolder)
                .Where(RecentLinksCleaner.IsShortcutFile)
                .ToList()
                .AsReadOnly();
        }
    }

    internal sealed class NoOpQuickAccessLockFactory : IQuickAccessLockFactory
    {
        public QuickAccessLock Lock(QuickAccessLockTarget target)
        {
            return new QuickAccessLock(
                target,
                string.Empty,
                Enumerable.Empty<string>(),
                Enumerable.Empty<IQuickAccessBackingFileHandle>(),
                new DefaultRecentLinkFileSystem());
        }
    }

    /// <summary>
    /// Holds backing file locks for Windows Quick Access.
    /// </summary>
    /// <remarks>
    /// This object affects the current Windows user's Quick Access backing files while it is alive. Manage the lock
    /// with <see langword="using"/> or <see cref="Dispose"/>. The object does not declare thread safety; concurrent
    /// calls to <see cref="Unlock()"/> and <see cref="Dispose"/> are not recommended.
    /// </remarks>
    public sealed class QuickAccessLock : IDisposable
    {
        private readonly IReadOnlyList<IQuickAccessBackingFileHandle> _handles;
        private readonly IRecentLinkFileSystem _fileSystem;
        private bool _disposed;

        internal QuickAccessLock(
            QuickAccessLockTarget target,
            string recentFolder,
            IEnumerable<string> initialShortcutPaths,
            IEnumerable<IQuickAccessBackingFileHandle> handles,
            IRecentLinkFileSystem fileSystem)
        {
            Target = target;
            RecentFolder = recentFolder ?? throw new ArgumentNullException(nameof(recentFolder));
            InitialShortcutPaths = (initialShortcutPaths ?? Enumerable.Empty<string>()).ToList().AsReadOnly();
            _handles = (handles ?? Enumerable.Empty<IQuickAccessBackingFileHandle>()).ToList().AsReadOnly();
            _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));
        }

        /// <summary>
        /// Gets the locked Quick Access target.
        /// </summary>
        public QuickAccessLockTarget Target { get; }

        /// <summary>
        /// Gets the Windows Recent folder path used for shortcut snapshots.
        /// </summary>
        public string RecentFolder { get; }

        /// <summary>
        /// Gets the initial Recent shortcut snapshot captured when the lock was acquired.
        /// </summary>
        public IReadOnlyList<string> InitialShortcutPaths { get; }

        /// <summary>
        /// Gets the number of backing file handles held by this lock.
        /// </summary>
        public int LockedFileCount => _handles.Count;

        /// <summary>
        /// Releases the lock and returns a report.
        /// </summary>
        /// <returns>The unlock report.</returns>
        public QuickAccessUnlockReport Unlock()
        {
            return Unlock(new QuickAccessUnlockOptions());
        }

        /// <summary>
        /// Releases the lock and returns a report.
        /// </summary>
        /// <param name="options">The unlock options.</param>
        /// <returns>The unlock report.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="options"/> is <see langword="null"/>.</exception>
        /// <exception cref="ObjectDisposedException">The lock has already been released.</exception>
        /// <remarks>
        /// When <see cref="QuickAccessUnlockOptions.CleanupNewRecentLinks"/> is enabled, shortcut deletion failures
        /// are recorded in the report instead of being thrown. The report describes the shortcut snapshot taken before
        /// cleanup runs, so successfully deleted new shortcuts still appear in
        /// <see cref="QuickAccessUnlockReport.CurrentShortcutPaths"/> and
        /// <see cref="QuickAccessUnlockReport.NewShortcutPaths"/>.
        /// </remarks>
        public QuickAccessUnlockReport Unlock(QuickAccessUnlockOptions options)
        {
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            ThrowIfDisposed();

            try
            {
                var currentShortcutPaths = QuickAccessShortcutSnapshot.EnumerateShortcutPaths(_fileSystem, RecentFolder);
                var newShortcutPaths = currentShortcutPaths.Except(InitialShortcutPaths, StringComparer.OrdinalIgnoreCase).ToList();
                var disappearedShortcutPaths = InitialShortcutPaths.Except(currentShortcutPaths, StringComparer.OrdinalIgnoreCase).ToList();
                var deletedShortcutPaths = new List<string>();
                var failures = new List<QuickAccessUnlockFailure>();

                if (options.CleanupNewRecentLinks)
                {
                    foreach (var path in newShortcutPaths)
                    {
                        try
                        {
                            _fileSystem.DeleteFile(path);
                            deletedShortcutPaths.Add(path);
                        }
                        catch (Exception ex)
                        {
                            failures.Add(new QuickAccessUnlockFailure(path, ex));
                        }
                    }
                }

                return new QuickAccessUnlockReport(
                    RecentFolder,
                    InitialShortcutPaths,
                    currentShortcutPaths,
                    disappearedShortcutPaths,
                    deletedShortcutPaths,
                    failures);
            }
            finally
            {
                DisposeHandles();
                _disposed = true;
            }
        }

        /// <summary>
        /// Releases the lock without creating an unlock report.
        /// </summary>
        public void Dispose()
        {
            if (_disposed)
                return;

            DisposeHandles();
            _disposed = true;
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(QuickAccessLock));
        }

        private void DisposeHandles()
        {
            foreach (var handle in _handles)
                handle.Dispose();
        }
    }
}
