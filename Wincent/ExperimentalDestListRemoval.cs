using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;

namespace Wincent
{
    /// <summary>
    /// Provides experimental DestList removal helpers.
    /// </summary>
    public static class ExperimentalDestListRemoval
    {
        /// <summary>
        /// Experimentally removes DestList entries by deleting Explorer backing data and waiting for rebuild.
        /// </summary>
        /// <param name="kind">The automatic destination family to modify.</param>
        /// <param name="targetPaths">The target paths to remove.</param>
        /// <param name="options">The experimental removal options.</param>
        /// <returns>The experimental removal report.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="targetPaths"/> or <paramref name="options"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="targetPaths"/> is empty.</exception>
        /// <remarks>
        /// This method deletes Shell-maintained files for the current Windows user. It is experimental, best-effort,
        /// non-transactional, and may temporarily affect Explorer Quick Access state.
        /// </remarks>
        public static ExperimentalRemoveReport RemoveEntryPathsByRebuild(
            AutomaticDestinationsKind kind,
            IEnumerable<string> targetPaths,
            ExperimentalRemoveOptions options)
        {
            return CreateDefaultEngine().RemoveEntryPathsByRebuild(kind, targetPaths, options);
        }

        /// <summary>
        /// Experimentally removes DestList entries by deleting Explorer backing data and waiting for rebuild.
        /// </summary>
        /// <param name="kind">The automatic destination family to modify.</param>
        /// <param name="entries">The entries to remove.</param>
        /// <param name="options">The experimental removal options.</param>
        /// <returns>The experimental removal report.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="entries"/> or <paramref name="options"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="entries"/> is empty or contains an entry with an empty path.</exception>
        /// <remarks>
        /// This method has the same risks as <see cref="RemoveEntryPathsByRebuild(AutomaticDestinationsKind, IEnumerable{string}, ExperimentalRemoveOptions)"/>.
        /// </remarks>
        public static ExperimentalRemoveReport RemoveEntriesByRebuild(
            AutomaticDestinationsKind kind,
            IEnumerable<DestListEntry> entries,
            ExperimentalRemoveOptions options)
        {
            if (entries == null)
                throw new ArgumentNullException(nameof(entries));

            return RemoveEntryPathsByRebuild(kind, entries.Select(entry => entry?.Path), options);
        }

        internal static ExperimentalDestListRemovalEngine CreateDefaultEngine()
        {
            var nativeMethods = new DefaultNativeMethods();
            return new ExperimentalDestListRemovalEngine(
                new QuickAccessDataFiles(),
                new WindowsRecentFolder(nativeMethods),
                new ShellLinkTargetResolver(nativeMethods),
                new DefaultRecentLinkFileSystem(),
                new ShellExplorerRefresher(nativeMethods),
                new DefaultDestListMetadataReader(),
                new DefaultExperimentalRemovalDelay(),
                new DefaultExperimentalRemovalFileSystem());
        }
    }

    internal interface IExperimentalRemovalDelay
    {
        void Sleep(TimeSpan delay);
    }

    internal interface IExperimentalRemovalFileSystem
    {
        bool FileExists(string path);

        void DeleteFile(string path);
    }

    internal sealed class DefaultExperimentalRemovalDelay : IExperimentalRemovalDelay
    {
        public void Sleep(TimeSpan delay)
        {
            Thread.Sleep(delay);
        }
    }

    internal sealed class DefaultExperimentalRemovalFileSystem : IExperimentalRemovalFileSystem
    {
        public bool FileExists(string path)
        {
            return File.Exists(path);
        }

        public void DeleteFile(string path)
        {
            File.Delete(path);
        }
    }

    internal sealed class ExperimentalDestListRemovalEngine
    {
        private static readonly TimeSpan InitialParseTimeout = TimeSpan.FromSeconds(2);
        private static readonly TimeSpan RebuildPollTimeout = TimeSpan.FromSeconds(5);
        private static readonly TimeSpan PollInterval = TimeSpan.FromMilliseconds(100);

        private readonly IQuickAccessDataFiles _dataFiles;
        private readonly IWindowsRecentFolder _recentFolder;
        private readonly IShortcutTargetResolver _shortcutResolver;
        private readonly IRecentLinkFileSystem _recentLinkFileSystem;
        private readonly IExplorerRefresher _explorerRefresher;
        private readonly IDestListMetadataReader _destListReader;
        private readonly IExperimentalRemovalDelay _delay;
        private readonly IExperimentalRemovalFileSystem _fileSystem;

        public ExperimentalDestListRemovalEngine(
            IQuickAccessDataFiles dataFiles,
            IWindowsRecentFolder recentFolder,
            IShortcutTargetResolver shortcutResolver,
            IRecentLinkFileSystem recentLinkFileSystem,
            IExplorerRefresher explorerRefresher,
            IDestListMetadataReader destListReader,
            IExperimentalRemovalDelay delay,
            IExperimentalRemovalFileSystem fileSystem)
        {
            _dataFiles = dataFiles ?? throw new ArgumentNullException(nameof(dataFiles));
            _recentFolder = recentFolder ?? throw new ArgumentNullException(nameof(recentFolder));
            _shortcutResolver = shortcutResolver ?? throw new ArgumentNullException(nameof(shortcutResolver));
            _recentLinkFileSystem = recentLinkFileSystem ?? throw new ArgumentNullException(nameof(recentLinkFileSystem));
            _explorerRefresher = explorerRefresher ?? throw new ArgumentNullException(nameof(explorerRefresher));
            _destListReader = destListReader ?? throw new ArgumentNullException(nameof(destListReader));
            _delay = delay ?? throw new ArgumentNullException(nameof(delay));
            _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));
        }

        public ExperimentalRemoveReport RemoveEntryPathsByRebuild(
            AutomaticDestinationsKind kind,
            IEnumerable<string> targetPaths,
            ExperimentalRemoveOptions options)
        {
            if (targetPaths == null)
                throw new ArgumentNullException(nameof(targetPaths));
            if (options == null)
                throw new ArgumentNullException(nameof(options));
            if (options.RebuildDelay < TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(options), "RebuildDelay cannot be negative.");

            var requestedPaths = targetPaths.ToList();
            if (requestedPaths.Count == 0)
                throw new ArgumentException("targetPaths must not be empty.", nameof(targetPaths));
            if (requestedPaths.Any(string.IsNullOrWhiteSpace))
                throw new ArgumentException("targetPaths cannot contain empty paths.", nameof(targetPaths));

            string recentFolder = _recentFolder.GetPath();
            string destinationPath = DestinationPathForKind(kind);

            if (!_fileSystem.FileExists(destinationPath))
                throw new FileNotFoundException($"Automatic destination file not found: {destinationPath}", destinationPath);

            var before = ParseFileWithRetries(destinationPath, InitialParseTimeout);
            var matchingPathsBefore = MatchingDestPaths(before.DestList.Entries, requestedPaths);

            _fileSystem.DeleteFile(destinationPath);
            bool destinationDeleted = true;

            IReadOnlyList<string> deletedShortcutPaths = new List<string>().AsReadOnly();
            IReadOnlyList<string> missingShortcutTargetPaths = new List<string>().AsReadOnly();
            RebuildWaitResult rebuild;
            try
            {
                var deletedLinks = DeleteMatchingRecentLinks(recentFolder, requestedPaths);
                deletedShortcutPaths = deletedLinks.Select(link => link.ShortcutPath).ToList().AsReadOnly();
                missingShortcutTargetPaths = requestedPaths
                    .Where(target => !deletedLinks.Any(link => WindowsPathComparer.Equals(link.TargetPath, target)))
                    .ToList().AsReadOnly();

                _explorerRefresher.Refresh(TimeSpan.FromSeconds(10));
                _delay.Sleep(options.RebuildDelay);

                rebuild = WaitForRebuiltDestination(destinationPath, requestedPaths, RebuildPollTimeout);
            }
            catch (Exception ex)
            {
                rebuild = new RebuildWaitResult(false, null, ex.Message, new List<string>());
            }

            bool success = rebuild.Rebuilt && rebuild.RemainingPathsAfterRebuild.Count == 0;

            return new ExperimentalRemoveReport(
                kind,
                recentFolder,
                destinationPath,
                requestedPaths,
                matchingPathsBefore,
                deletedShortcutPaths,
                missingShortcutTargetPaths,
                destinationDeleted,
                rebuild.Rebuilt,
                rebuild.RebuildParseElapsed,
                rebuild.RebuildParseError,
                rebuild.RemainingPathsAfterRebuild,
                success);
        }

        private string DestinationPathForKind(AutomaticDestinationsKind kind)
        {
            switch (kind)
            {
                case AutomaticDestinationsKind.RecentFiles:
                    return _dataFiles.RecentFilesPath;
                case AutomaticDestinationsKind.FrequentFolders:
                    return _dataFiles.FrequentFoldersPath;
                default:
                    throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unsupported automatic destination kind.");
            }
        }

        private AutomaticDestinations ParseFileWithRetries(string path, TimeSpan timeout)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            while (true)
            {
                try
                {
                    return _destListReader.ParseFile(path);
                }
                catch
                {
                    if (stopwatch.Elapsed >= timeout)
                        throw;

                    _delay.Sleep(PollInterval);
                }
            }
        }

        private RebuildWaitResult WaitForRebuiltDestination(
            string path,
            IReadOnlyList<string> requestedPaths,
            TimeSpan timeout)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            string lastParseError = null;

            while (true)
            {
                if (_fileSystem.FileExists(path))
                {
                    try
                    {
                        var parsed = _destListReader.ParseFile(path);
                        return new RebuildWaitResult(
                            true,
                            stopwatch.Elapsed,
                            lastParseError,
                            MatchingDestPaths(parsed.DestList.Entries, requestedPaths));
                    }
                    catch (Exception ex)
                    {
                        lastParseError = ex.Message;
                    }
                }

                if (stopwatch.Elapsed >= timeout)
                    return new RebuildWaitResult(false, null, lastParseError, new List<string>());

                _delay.Sleep(PollInterval);
            }
        }

        private static IReadOnlyList<string> MatchingDestPaths(
            IEnumerable<DestListEntry> entries,
            IReadOnlyList<string> targetPaths)
        {
            return entries
                .Where(entry => targetPaths.Any(target => WindowsPathComparer.Equals(entry.Path, target)))
                .Select(entry => entry.Path)
                .ToList()
                .AsReadOnly();
        }

        private IReadOnlyList<DeletedRecentLink> DeleteMatchingRecentLinks(string recentFolder, IReadOnlyList<string> targetPaths)
        {
            var deleted = new List<DeletedRecentLink>();
            foreach (var shortcutPath in _recentLinkFileSystem.EnumerateFiles(recentFolder))
            {
                if (!RecentLinksCleaner.IsShortcutFile(shortcutPath))
                    continue;

                string target = _shortcutResolver.ResolveTarget(shortcutPath, TimeSpan.FromSeconds(10));
                if (string.IsNullOrWhiteSpace(target))
                    continue;

                if (!targetPaths.Any(requested => WindowsPathComparer.Equals(target, requested)))
                    continue;

                _recentLinkFileSystem.DeleteFile(shortcutPath);
                deleted.Add(new DeletedRecentLink(shortcutPath, target));
            }

            return deleted.AsReadOnly();
        }

        private sealed class DeletedRecentLink
        {
            public DeletedRecentLink(string shortcutPath, string targetPath)
            {
                ShortcutPath = shortcutPath;
                TargetPath = targetPath;
            }

            public string ShortcutPath { get; }

            public string TargetPath { get; }
        }

        private sealed class RebuildWaitResult
        {
            public RebuildWaitResult(
                bool rebuilt,
                TimeSpan? rebuildParseElapsed,
                string rebuildParseError,
                IReadOnlyList<string> remainingPathsAfterRebuild)
            {
                Rebuilt = rebuilt;
                RebuildParseElapsed = rebuildParseElapsed;
                RebuildParseError = rebuildParseError;
                RemainingPathsAfterRebuild = remainingPathsAfterRebuild;
            }

            public bool Rebuilt { get; }

            public TimeSpan? RebuildParseElapsed { get; }

            public string RebuildParseError { get; }

            public IReadOnlyList<string> RemainingPathsAfterRebuild { get; }
        }
    }
}
