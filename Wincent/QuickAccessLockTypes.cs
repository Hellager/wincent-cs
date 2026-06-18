using System;
using System.Collections.Generic;
using System.Linq;

namespace Wincent
{
    /// <summary>
    /// Identifies a Quick Access lock target.
    /// </summary>
    public enum QuickAccessLockTarget
    {
        /// <summary>
        /// Recent files backing data.
        /// </summary>
        RecentFiles,

        /// <summary>
        /// Frequent folders backing data.
        /// </summary>
        FrequentFolders,

        /// <summary>
        /// All supported backing data.
        /// </summary>
        All
    }

    /// <summary>
    /// Configures an unlock operation.
    /// </summary>
    public sealed class QuickAccessUnlockOptions
    {
        /// <summary>
        /// Gets or sets whether new Recent shortcut files should be deleted during unlock.
        /// </summary>
        public bool CleanupNewRecentLinks { get; set; }
    }

    /// <summary>
    /// Describes a Quick Access unlock report.
    /// </summary>
    public sealed class QuickAccessUnlockReport
    {
        internal QuickAccessUnlockReport(
            string recentFolder,
            IEnumerable<string> initialShortcutPaths,
            IEnumerable<string> currentShortcutPaths,
            IEnumerable<string> disappearedShortcutPaths,
            IEnumerable<string> deletedShortcutPaths,
            IEnumerable<QuickAccessUnlockFailure> failedShortcutDeletions)
        {
            RecentFolder = recentFolder;
            InitialShortcutPaths = (initialShortcutPaths ?? Enumerable.Empty<string>()).ToList().AsReadOnly();
            CurrentShortcutPaths = (currentShortcutPaths ?? Enumerable.Empty<string>()).ToList().AsReadOnly();
            DisappearedShortcutPaths = (disappearedShortcutPaths ?? Enumerable.Empty<string>()).ToList().AsReadOnly();
            DeletedShortcutPaths = (deletedShortcutPaths ?? Enumerable.Empty<string>()).ToList().AsReadOnly();
            FailedShortcutDeletions = (failedShortcutDeletions ?? Enumerable.Empty<QuickAccessUnlockFailure>()).ToList().AsReadOnly();
            NewShortcutPaths = CurrentShortcutPaths.Except(InitialShortcutPaths, StringComparer.OrdinalIgnoreCase).ToList().AsReadOnly();
        }

        /// <summary>
        /// Gets the Windows Recent folder path.
        /// </summary>
        public string RecentFolder { get; }

        /// <summary>
        /// Gets the initial shortcut snapshot.
        /// </summary>
        public IReadOnlyList<string> InitialShortcutPaths { get; }

        /// <summary>
        /// Gets the current shortcut snapshot.
        /// </summary>
        public IReadOnlyList<string> CurrentShortcutPaths { get; }

        /// <summary>
        /// Gets the shortcuts that appeared after the lock was acquired.
        /// </summary>
        public IReadOnlyList<string> NewShortcutPaths { get; }

        /// <summary>
        /// Gets the shortcuts that were present in the initial snapshot but absent from the current snapshot.
        /// </summary>
        public IReadOnlyList<string> DisappearedShortcutPaths { get; }

        /// <summary>
        /// Gets new shortcuts that were successfully deleted during unlock cleanup.
        /// </summary>
        public IReadOnlyList<string> DeletedShortcutPaths { get; }

        /// <summary>
        /// Gets failed shortcut deletions.
        /// </summary>
        public IReadOnlyList<QuickAccessUnlockFailure> FailedShortcutDeletions { get; }
    }

    /// <summary>
    /// Describes a failed shortcut deletion.
    /// </summary>
    public sealed class QuickAccessUnlockFailure
    {
        internal QuickAccessUnlockFailure(string path, Exception error)
        {
            Path = path;
            Error = error;
        }

        /// <summary>
        /// Gets the shortcut path.
        /// </summary>
        public string Path { get; }

        /// <summary>
        /// Gets the deletion error.
        /// </summary>
        public Exception Error { get; }
    }
}
