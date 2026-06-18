using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;

namespace Wincent
{
    /// <summary>
    /// Configures restore-defaults operations.
    /// </summary>
    public sealed class RestoreDefaultsOptions
    {
        /// <summary>
        /// Gets or sets whether Explorer should be refreshed while restoring defaults.
        /// </summary>
        public bool RefreshExplorer { get; set; } = true;

        /// <summary>
        /// Gets or sets whether unresolved or unknown-type Windows Recent shortcuts should be deleted.
        /// </summary>
        public bool DeepLnkCleanup { get; set; }

        /// <summary>
        /// Gets or sets the timeout used for resolving each Windows Recent shortcut.
        /// </summary>
        public TimeSpan LnkResolveTimeout { get; set; } = TimeSpan.FromSeconds(10);

        /// <summary>
        /// Gets or sets the timeout used for Shell clear and Explorer refresh operations.
        /// </summary>
        public TimeSpan ClearTimeout { get; set; } = TimeSpan.FromSeconds(10);

        /// <summary>
        /// Gets or sets the delay before polling for a rebuilt Frequent Folders backing file.
        /// </summary>
        public TimeSpan RebuildDelay { get; set; } = TimeSpan.FromMilliseconds(500);

        /// <summary>
        /// Gets or sets the timeout for waiting for a rebuilt Frequent Folders backing file.
        /// </summary>
        public TimeSpan RebuildPollTimeout { get; set; } = TimeSpan.FromSeconds(5);

        internal void Validate()
        {
            if (LnkResolveTimeout <= TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(LnkResolveTimeout), "LnkResolveTimeout must be positive.");
            if (ClearTimeout <= TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(ClearTimeout), "ClearTimeout must be positive.");
            if (RebuildDelay < TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(RebuildDelay), "RebuildDelay cannot be negative.");
            if (RebuildPollTimeout <= TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(RebuildPollTimeout), "RebuildPollTimeout must be positive.");
        }
    }

    /// <summary>
    /// Combined restore report for one or more Quick Access sections.
    /// </summary>
    public sealed class RestoreDefaultsReport
    {
        internal RestoreDefaultsReport(RecentRestoreReport recentReport, FrequentRestoreReport frequentReport)
        {
            RecentReport = recentReport;
            FrequentReport = frequentReport;
        }

        /// <summary>Gets the Recent Files restore report, when requested.</summary>
        public RecentRestoreReport RecentReport { get; }

        /// <summary>Gets the Frequent Folders restore report, when requested.</summary>
        public FrequentRestoreReport FrequentReport { get; }

        /// <summary>Gets whether every requested restore section succeeded.</summary>
        public bool Success
        {
            get
            {
                return (RecentReport == null || RecentReport.Success) &&
                       (FrequentReport == null || FrequentReport.Success);
            }
        }
    }

    /// <summary>
    /// Restore report for Recent Files.
    /// </summary>
    public sealed class RecentRestoreReport
    {
        internal RecentRestoreReport(IReadOnlyList<string> deletedLnkPaths, bool recentFilesCleared, Exception error)
        {
            DeletedLnkPaths = deletedLnkPaths ?? Array.Empty<string>();
            RecentFilesCleared = recentFilesCleared;
            Error = error;
        }

        /// <summary>Gets Windows Recent shortcut files deleted during restore.</summary>
        public IReadOnlyList<string> DeletedLnkPaths { get; }

        /// <summary>Gets whether the Shell Recent Files clear operation completed.</summary>
        public bool RecentFilesCleared { get; }

        /// <summary>Gets the restore failure captured in the report, if any.</summary>
        public Exception Error { get; }

        /// <summary>Gets whether Recent Files were restored successfully.</summary>
        public bool Success => RecentFilesCleared && Error == null;
    }

    /// <summary>
    /// Restore report for Frequent Folders.
    /// </summary>
    public sealed class FrequentRestoreReport
    {
        internal FrequentRestoreReport(
            IReadOnlyList<string> deletedLnkPaths,
            bool backingFileDeleted,
            bool rebuilt,
            IReadOnlyList<string> nonDefaultRawPaths,
            FrequentRawPathRemoveReport rawPathRemoveReport,
            Exception error)
        {
            DeletedLnkPaths = deletedLnkPaths ?? Array.Empty<string>();
            BackingFileDeleted = backingFileDeleted;
            Rebuilt = rebuilt;
            NonDefaultRawPaths = nonDefaultRawPaths ?? Array.Empty<string>();
            RawPathRemoveReport = rawPathRemoveReport;
            Error = error;
        }

        /// <summary>Gets Windows Recent shortcut files deleted during restore.</summary>
        public IReadOnlyList<string> DeletedLnkPaths { get; }

        /// <summary>Gets whether the Frequent Folders backing file was deleted or already absent.</summary>
        public bool BackingFileDeleted { get; }

        /// <summary>Gets whether a rebuilt Frequent Folders backing file was detected and parsed.</summary>
        public bool Rebuilt { get; }

        /// <summary>Gets non-default raw paths still present after restore cleanup.</summary>
        public IReadOnlyList<string> NonDefaultRawPaths { get; }

        /// <summary>Gets the raw-path cleanup report, when non-default raw paths required a cleanup pass.</summary>
        public FrequentRawPathRemoveReport RawPathRemoveReport { get; }

        /// <summary>Gets the restore failure captured in the report, if any.</summary>
        public Exception Error { get; }

        /// <summary>Gets whether Frequent Folders were restored successfully.</summary>
        public bool Success
        {
            get
            {
                if (!BackingFileDeleted || !Rebuilt || Error != null)
                    return false;

                return RawPathRemoveReport == null
                    ? NonDefaultRawPaths.Count == 0
                    : RawPathRemoveReport.Success;
            }
        }
    }

    /// <summary>
    /// Report for the Frequent Folders raw-path cleanup pass.
    /// </summary>
    public sealed class FrequentRawPathRemoveReport
    {
        internal FrequentRawPathRemoveReport(
            IReadOnlyList<string> requestedRawPaths,
            bool backingFileDeleted,
            bool rebuilt,
            IReadOnlyList<string> remainingNonDefaultRawPaths,
            Exception error)
        {
            RequestedRawPaths = requestedRawPaths ?? Array.Empty<string>();
            BackingFileDeleted = backingFileDeleted;
            Rebuilt = rebuilt;
            RemainingNonDefaultRawPaths = remainingNonDefaultRawPaths ?? Array.Empty<string>();
            Error = error;
        }

        /// <summary>Gets non-default raw paths requested for cleanup.</summary>
        public IReadOnlyList<string> RequestedRawPaths { get; }

        /// <summary>Gets whether the backing file was deleted or already absent.</summary>
        public bool BackingFileDeleted { get; }

        /// <summary>Gets whether a rebuilt backing file was detected and parsed.</summary>
        public bool Rebuilt { get; }

        /// <summary>Gets non-default raw paths still present after the cleanup pass.</summary>
        public IReadOnlyList<string> RemainingNonDefaultRawPaths { get; }

        /// <summary>Gets the cleanup failure captured in the report, if any.</summary>
        public Exception Error { get; }

        /// <summary>Gets whether the cleanup pass removed all non-default raw paths.</summary>
        public bool Success => BackingFileDeleted && Rebuilt && RemainingNonDefaultRawPaths.Count == 0 && Error == null;
    }

    internal interface IQuickAccessRestoreEngine
    {
        RestoreDefaultsReport RestoreDefaults(
            QuickAccess target,
            RestoreDefaultsOptions options,
            Action<TimeSpan> clearRecentFiles,
            Action<TimeSpan> refreshExplorer);

        RecentRestoreReport RestoreRecentFilesDefaults(
            RestoreDefaultsOptions options,
            Action<TimeSpan> clearRecentFiles);

        FrequentRestoreReport RestoreFrequentFoldersDefaults(
            RestoreDefaultsOptions options,
            Action<TimeSpan> refreshExplorer);
    }

    internal interface IQuickAccessRestoreDelay
    {
        void Sleep(TimeSpan delay);
    }

    internal sealed class DefaultQuickAccessRestoreDelay : IQuickAccessRestoreDelay
    {
        public void Sleep(TimeSpan delay)
        {
            Thread.Sleep(delay);
        }
    }

    internal sealed class NoOpQuickAccessRestoreEngine : IQuickAccessRestoreEngine
    {
        public RestoreDefaultsReport RestoreDefaults(
            QuickAccess target,
            RestoreDefaultsOptions options,
            Action<TimeSpan> clearRecentFiles,
            Action<TimeSpan> refreshExplorer)
        {
            throw new InvalidOperationException("Restore defaults is disabled for this instance.");
        }

        public RecentRestoreReport RestoreRecentFilesDefaults(
            RestoreDefaultsOptions options,
            Action<TimeSpan> clearRecentFiles)
        {
            throw new InvalidOperationException("Restore defaults is disabled for this instance.");
        }

        public FrequentRestoreReport RestoreFrequentFoldersDefaults(
            RestoreDefaultsOptions options,
            Action<TimeSpan> refreshExplorer)
        {
            throw new InvalidOperationException("Restore defaults is disabled for this instance.");
        }
    }

    internal sealed class QuickAccessRestoreEngine : IQuickAccessRestoreEngine
    {
        private static readonly TimeSpan PollInterval = TimeSpan.FromMilliseconds(100);

        private readonly IQuickAccessDataFiles _dataFiles;
        private readonly IWindowsRecentFolder _recentFolder;
        private readonly IShortcutTargetResolver _shortcutResolver;
        private readonly IRecentLinkFileSystem _recentLinkFileSystem;
        private readonly IFileSystemOperations _fileSystem;
        private readonly IDestListMetadataReader _destListReader;
        private readonly IQuickAccessRestoreDelay _delay;

        public QuickAccessRestoreEngine(
            IQuickAccessDataFiles dataFiles,
            IWindowsRecentFolder recentFolder,
            IShortcutTargetResolver shortcutResolver,
            IRecentLinkFileSystem recentLinkFileSystem,
            IFileSystemOperations fileSystem,
            IDestListMetadataReader destListReader,
            IQuickAccessRestoreDelay delay)
        {
            _dataFiles = dataFiles ?? throw new ArgumentNullException(nameof(dataFiles));
            _recentFolder = recentFolder ?? throw new ArgumentNullException(nameof(recentFolder));
            _shortcutResolver = shortcutResolver ?? throw new ArgumentNullException(nameof(shortcutResolver));
            _recentLinkFileSystem = recentLinkFileSystem ?? throw new ArgumentNullException(nameof(recentLinkFileSystem));
            _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));
            _destListReader = destListReader ?? throw new ArgumentNullException(nameof(destListReader));
            _delay = delay ?? throw new ArgumentNullException(nameof(delay));
        }

        public RestoreDefaultsReport RestoreDefaults(
            QuickAccess target,
            RestoreDefaultsOptions options,
            Action<TimeSpan> clearRecentFiles,
            Action<TimeSpan> refreshExplorer)
        {
            switch (target)
            {
                case QuickAccess.RecentFiles:
                    return new RestoreDefaultsReport(RestoreRecentFilesDefaults(options, clearRecentFiles), null);
                case QuickAccess.FrequentFolders:
                    return new RestoreDefaultsReport(null, RestoreFrequentFoldersDefaults(options, refreshExplorer));
                case QuickAccess.All:
                    return new RestoreDefaultsReport(
                        RestoreRecentFilesDefaults(options, clearRecentFiles),
                        RestoreFrequentFoldersDefaults(options, refreshExplorer));
                default:
                    throw new ArgumentOutOfRangeException(nameof(target), target, "Unsupported Quick Access target.");
            }
        }

        public RecentRestoreReport RestoreRecentFilesDefaults(
            RestoreDefaultsOptions options,
            Action<TimeSpan> clearRecentFiles)
        {
            if (options == null)
                throw new ArgumentNullException(nameof(options));
            if (clearRecentFiles == null)
                throw new ArgumentNullException(nameof(clearRecentFiles));

            var cleanup = DeleteMatchingLnkFiles(RestoreTarget.RecentFiles, options);
            if (cleanup.Error != null)
                return new RecentRestoreReport(cleanup.DeletedLnkPaths, false, cleanup.Error);

            try
            {
                clearRecentFiles(options.ClearTimeout);
                return new RecentRestoreReport(cleanup.DeletedLnkPaths, true, null);
            }
            catch (Exception ex)
            {
                return new RecentRestoreReport(cleanup.DeletedLnkPaths, false, ex);
            }
        }

        public FrequentRestoreReport RestoreFrequentFoldersDefaults(
            RestoreDefaultsOptions options,
            Action<TimeSpan> refreshExplorer)
        {
            if (options == null)
                throw new ArgumentNullException(nameof(options));
            if (refreshExplorer == null)
                throw new ArgumentNullException(nameof(refreshExplorer));

            var cleanup = DeleteMatchingLnkFiles(RestoreTarget.FrequentFolders, options);
            if (cleanup.Error != null)
                return new FrequentRestoreReport(cleanup.DeletedLnkPaths, false, false, Array.Empty<string>(), null, cleanup.Error);

            bool backingFileDeleted;
            try
            {
                DeleteFrequentFoldersBackingFile();
                backingFileDeleted = true;
            }
            catch (Exception ex)
            {
                return new FrequentRestoreReport(cleanup.DeletedLnkPaths, false, false, Array.Empty<string>(), null, ex);
            }

            try
            {
                RefreshExplorerIfRequested(options, refreshExplorer);
                _delay.Sleep(options.RebuildDelay);
                var rebuilt = WaitForFrequentFoldersRebuild(options.RebuildPollTimeout);
                var nonDefault = NonDefaultRawPaths(rebuilt);
                if (nonDefault.Count == 0)
                    return new FrequentRestoreReport(cleanup.DeletedLnkPaths, backingFileDeleted, true, nonDefault, null, null);

                var rawPathReport = RemoveFrequentRawPathsOnce(options, rebuilt, refreshExplorer);
                return new FrequentRestoreReport(
                    cleanup.DeletedLnkPaths,
                    backingFileDeleted,
                    true,
                    rawPathReport.RemainingNonDefaultRawPaths,
                    rawPathReport,
                    null);
            }
            catch (Exception ex)
            {
                return new FrequentRestoreReport(cleanup.DeletedLnkPaths, backingFileDeleted, false, Array.Empty<string>(), null, ex);
            }
        }

        private FrequentRawPathRemoveReport RemoveFrequentRawPathsOnce(
            RestoreDefaultsOptions options,
            IReadOnlyList<DestListEntry> entries,
            Action<TimeSpan> refreshExplorer)
        {
            var requested = NonDefaultRawPaths(entries);
            if (requested.Count == 0)
                return new FrequentRawPathRemoveReport(Array.Empty<string>(), false, false, Array.Empty<string>(), null);

            bool backingFileDeleted;
            try
            {
                DeleteFrequentFoldersBackingFile();
                backingFileDeleted = true;
            }
            catch (Exception ex)
            {
                return new FrequentRawPathRemoveReport(requested, false, false, Array.Empty<string>(), ex);
            }

            try
            {
                RefreshExplorerIfRequested(options, refreshExplorer);
                _delay.Sleep(options.RebuildDelay);
                var rebuilt = WaitForFrequentFoldersRebuild(options.RebuildPollTimeout);
                return new FrequentRawPathRemoveReport(requested, backingFileDeleted, true, NonDefaultRawPaths(rebuilt), null);
            }
            catch (Exception ex)
            {
                return new FrequentRawPathRemoveReport(requested, backingFileDeleted, false, Array.Empty<string>(), ex);
            }
        }

        private LnkCleanup DeleteMatchingLnkFiles(RestoreTarget target, RestoreDefaultsOptions options)
        {
            var deleted = new List<string>();
            IEnumerable<string> paths;
            try
            {
                string recentFolder = _recentFolder.GetPath();
                paths = _recentLinkFileSystem.EnumerateFiles(recentFolder)
                    .Where(RecentLinksCleaner.IsShortcutFile)
                    .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            catch (Exception ex)
            {
                return new LnkCleanup(deleted.AsReadOnly(), ex);
            }

            foreach (string path in paths)
            {
                ShortcutResolution resolution;
                try
                {
                    resolution = ResolveShortcut(path, options.LnkResolveTimeout);
                }
                catch (Exception ex)
                {
                    return new LnkCleanup(deleted.AsReadOnly(), ex);
                }

                if (!ShouldDeleteLnkForRestore(target, resolution, options.DeepLnkCleanup))
                    continue;

                try
                {
                    _recentLinkFileSystem.DeleteFile(path);
                    deleted.Add(path);
                }
                catch (Exception ex)
                {
                    return new LnkCleanup(deleted.AsReadOnly(), ex);
                }
            }

            return new LnkCleanup(deleted.AsReadOnly(), null);
        }

        private ShortcutResolution ResolveShortcut(string path, TimeSpan timeout)
        {
            var typedResolver = _shortcutResolver as IShortcutResolutionResolver;
            if (typedResolver != null)
                return typedResolver.Resolve(path, timeout);

            string target = _shortcutResolver.ResolveTarget(path, timeout);
            return string.IsNullOrWhiteSpace(target) ? null : new ShortcutResolution(target, null);
        }

        private static bool ShouldDeleteLnkForRestore(RestoreTarget target, ShortcutResolution resolution, bool deep)
        {
            bool? isDirectory = resolution == null ? null : resolution.IsDirectory;
            if (target == RestoreTarget.RecentFiles && isDirectory == false)
                return true;
            if (target == RestoreTarget.FrequentFolders && isDirectory == true)
                return true;

            return !isDirectory.HasValue && deep;
        }

        private void DeleteFrequentFoldersBackingFile()
        {
            string path = _dataFiles.FrequentFoldersPath;
            if (_fileSystem.FileExists(path))
                _fileSystem.DeleteFile(path);
        }

        private static void RefreshExplorerIfRequested(RestoreDefaultsOptions options, Action<TimeSpan> refreshExplorer)
        {
            if (options.RefreshExplorer)
                refreshExplorer(options.ClearTimeout);
        }

        private IReadOnlyList<DestListEntry> WaitForFrequentFoldersRebuild(TimeSpan timeout)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            Exception lastError = null;
            string path = _dataFiles.FrequentFoldersPath;

            while (true)
            {
                if (_fileSystem.FileExists(path))
                {
                    try
                    {
                        return _destListReader.ParseFile(path).DestList.Entries;
                    }
                    catch (Exception ex)
                    {
                        lastError = ex;
                    }
                }

                if (stopwatch.Elapsed >= timeout)
                    throw new TimeoutException("Timed out waiting for Frequent Folders backing file to rebuild.", lastError);

                _delay.Sleep(PollInterval);
            }
        }

        private static IReadOnlyList<string> NonDefaultRawPaths(IEnumerable<DestListEntry> entries)
        {
            return entries
                .Where(entry => !StartsWithKnownFolder(entry.RawPath))
                .Select(entry => entry.RawPath ?? string.Empty)
                .ToList()
                .AsReadOnly();
        }

        private static bool StartsWithKnownFolder(string rawPath)
        {
            return rawPath != null &&
                   rawPath.StartsWith("knownfolder:", StringComparison.OrdinalIgnoreCase);
        }

        private enum RestoreTarget
        {
            RecentFiles,
            FrequentFolders
        }

        private sealed class LnkCleanup
        {
            public LnkCleanup(IReadOnlyList<string> deletedLnkPaths, Exception error)
            {
                DeletedLnkPaths = deletedLnkPaths;
                Error = error;
            }

            public IReadOnlyList<string> DeletedLnkPaths { get; }

            public Exception Error { get; }
        }
    }
}
