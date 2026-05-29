using System;
using System.Collections.Generic;
using System.Linq;

namespace Wincent
{
    /// <summary>
    /// Configures a <see cref="QuickAccessManager"/>.
    /// </summary>
    public sealed class QuickAccessManagerOptions
    {
        /// <summary>
        /// Gets or sets the timeout for Windows and PowerShell operations.
        /// </summary>
        public TimeSpan Timeout { get; set; } = TimeSpan.FromSeconds(10);

        /// <summary>
        /// Gets or sets the retry policy for transient fallback operations.
        /// </summary>
        public RetryPolicy RetryPolicy { get; set; } = RetryPolicy.Standard;
    }

    /// <summary>
    /// Configures add operations.
    /// </summary>
    public sealed class AddOptions
    {
        /// <summary>
        /// Gets or sets whether Recent Files backing data should be refreshed after adding a recent file.
        /// </summary>
        /// <remarks>
        /// This maps to the Rust upstream <c>force_update</c> option through its <c>refresh_recent_files()</c>
        /// convenience method.
        /// </remarks>
        public bool RefreshRecentFiles { get; set; }
    }

    /// <summary>
    /// Configures remove operations.
    /// </summary>
    public sealed class RemoveOptions
    {
        /// <summary>
        /// Gets or sets whether matching Windows Recent shortcut files should be removed after shell removal.
        /// </summary>
        /// <remarks>
        /// The option is part of the phase 0 public API. The actual deep cleanup behavior is implemented in a later
        /// migration phase.
        /// </remarks>
        public bool DeepCleanRecentLinks { get; set; }
    }

    /// <summary>
    /// Configures clear operations.
    /// </summary>
    public sealed class ClearOptions
    {
        /// <summary>
        /// Gets or sets whether pinned folders should also be removed when clearing frequent folders.
        /// </summary>
        public bool RemovePinnedFolders { get; set; }

        /// <summary>
        /// Gets or sets whether Explorer windows should be refreshed after a successful clear.
        /// </summary>
        /// <remarks>
        /// A fully successful clear propagates refresh failures after the native Shell refresh and PowerShell fallback
        /// both fail. A partial clear treats refresh as best-effort and preserves the original clear failure.
        /// </remarks>
        public bool RefreshExplorer { get; set; }
    }

    /// <summary>
    /// Configures batch operations.
    /// </summary>
    public sealed class BatchOptions
    {
        /// <summary>
        /// Gets or sets whether Recent Files backing data should be refreshed once after a batch add.
        /// </summary>
        public bool RefreshRecentFiles { get; set; }
    }

    /// <summary>
    /// Represents a single Quick Access batch item.
    /// </summary>
    public sealed class QuickAccessItem
    {
        /// <summary>
        /// Initializes a new batch item.
        /// </summary>
        /// <param name="path">The path to add or remove.</param>
        /// <param name="target">The target Quick Access section.</param>
        public QuickAccessItem(string path, QuickAccess target)
        {
            if (path == null)
                throw new ArgumentNullException(nameof(path));

            Path = path;
            Target = target;
        }

        /// <summary>
        /// Gets the item path.
        /// </summary>
        public string Path { get; }

        /// <summary>
        /// Gets the target Quick Access section.
        /// </summary>
        public QuickAccess Target { get; }

        /// <summary>
        /// Creates a recent-file batch item.
        /// </summary>
        /// <param name="path">The file path.</param>
        /// <returns>A batch item targeting recent files.</returns>
        public static QuickAccessItem RecentFile(string path) => new QuickAccessItem(path, QuickAccess.RecentFiles);

        /// <summary>
        /// Creates a frequent-folder batch item.
        /// </summary>
        /// <param name="path">The folder path.</param>
        /// <returns>A batch item targeting frequent folders.</returns>
        public static QuickAccessItem FrequentFolder(string path) => new QuickAccessItem(path, QuickAccess.FrequentFolders);
    }

    /// <summary>
    /// Describes a batch operation result.
    /// </summary>
    public sealed class BatchResult
    {
        internal BatchResult(IEnumerable<QuickAccessItem> succeeded, IEnumerable<BatchFailure> failed)
        {
            Succeeded = (succeeded ?? Enumerable.Empty<QuickAccessItem>()).ToList().AsReadOnly();
            Failed = (failed ?? Enumerable.Empty<BatchFailure>()).ToList().AsReadOnly();
            Total = Succeeded.Count + Failed.Count;
        }

        /// <summary>
        /// Gets the items that succeeded.
        /// </summary>
        public IReadOnlyList<QuickAccessItem> Succeeded { get; }

        /// <summary>
        /// Gets the items that failed.
        /// </summary>
        public IReadOnlyList<BatchFailure> Failed { get; }

        /// <summary>
        /// Gets the total number of attempted items.
        /// </summary>
        public int Total { get; }

        /// <summary>
        /// Gets whether every attempted item succeeded.
        /// </summary>
        public bool IsCompleteSuccess => Failed.Count == 0;

        /// <summary>
        /// Gets whether the batch contains both successes and failures.
        /// </summary>
        public bool HasPartialSuccess => Succeeded.Count > 0 && Failed.Count > 0;

        /// <summary>
        /// Gets the success rate from 0.0 to 1.0.
        /// </summary>
        public double SuccessRate => Total == 0 ? 1.0 : (double)Succeeded.Count / Total;
    }

    /// <summary>
    /// Describes one failed batch item.
    /// </summary>
    public sealed class BatchFailure
    {
        internal BatchFailure(QuickAccessItem item, Exception error)
        {
            Item = item ?? throw new ArgumentNullException(nameof(item));
            Error = error ?? throw new ArgumentNullException(nameof(error));
        }

        /// <summary>
        /// Gets the item that failed.
        /// </summary>
        public QuickAccessItem Item { get; }

        /// <summary>
        /// Gets the error for the item.
        /// </summary>
        public Exception Error { get; }
    }

    /// <summary>
    /// Controls retry behavior for transient fallback operations.
    /// </summary>
    public sealed class RetryPolicy
    {
        private static readonly Random JitterRandom = new Random();
        private static readonly object JitterLock = new object();

        /// <summary>
        /// Gets a policy that performs no retries.
        /// </summary>
        public static RetryPolicy None { get; } = new RetryPolicy(0, TimeSpan.Zero, TimeSpan.Zero, 1.0, false);

        /// <summary>
        /// Gets a short retry policy.
        /// </summary>
        public static RetryPolicy Fast { get; } = new RetryPolicy(2, TimeSpan.FromMilliseconds(50), TimeSpan.FromSeconds(1), 1.5, true);

        /// <summary>
        /// Gets the standard retry policy.
        /// </summary>
        public static RetryPolicy Standard { get; } = new RetryPolicy(3, TimeSpan.FromMilliseconds(100), TimeSpan.FromSeconds(5), 2.0, true);

        /// <summary>
        /// Gets an aggressive retry policy.
        /// </summary>
        public static RetryPolicy Aggressive { get; } = new RetryPolicy(5, TimeSpan.FromMilliseconds(200), TimeSpan.FromSeconds(10), 2.0, true);

        /// <summary>
        /// Initializes a retry policy.
        /// </summary>
        public RetryPolicy(int maxRetryCount, TimeSpan initialDelay, TimeSpan maxDelay, double backoffFactor, bool useJitter)
        {
            if (maxRetryCount < 0)
                throw new ArgumentOutOfRangeException(nameof(maxRetryCount));
            if (initialDelay < TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(initialDelay));
            if (maxDelay < TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(maxDelay));
            if (backoffFactor < 1.0)
                throw new ArgumentOutOfRangeException(nameof(backoffFactor));
            if (maxRetryCount > 0 && initialDelay == TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(initialDelay));
            if (maxRetryCount > 0 && maxDelay < initialDelay)
                throw new ArgumentOutOfRangeException(nameof(maxDelay));

            MaxRetryCount = maxRetryCount;
            InitialDelay = initialDelay;
            MaxDelay = maxDelay;
            BackoffFactor = backoffFactor;
            UseJitter = useJitter;
        }

        /// <summary>
        /// Gets the number of retries after the first attempt.
        /// </summary>
        public int MaxRetryCount { get; }

        /// <summary>
        /// Gets the initial retry delay.
        /// </summary>
        public TimeSpan InitialDelay { get; }

        /// <summary>
        /// Gets the maximum retry delay.
        /// </summary>
        public TimeSpan MaxDelay { get; }

        /// <summary>
        /// Gets the exponential backoff factor.
        /// </summary>
        public double BackoffFactor { get; }

        /// <summary>
        /// Gets whether jitter is applied to retry delays.
        /// </summary>
        public bool UseJitter { get; }

        /// <summary>
        /// Gets the delay for a retry attempt.
        /// </summary>
        /// <param name="retryAttempt">The zero-based retry attempt.</param>
        /// <returns>The delay for the retry attempt.</returns>
        /// <exception cref="InvalidOperationException">
        /// This policy does not retry, such as <see cref="None"/>.
        /// </exception>
        /// <exception cref="ArgumentOutOfRangeException">
        /// <paramref name="retryAttempt"/> is outside the configured retry range.
        /// </exception>
        public TimeSpan GetDelay(int retryAttempt)
        {
            if (MaxRetryCount == 0)
                throw new InvalidOperationException("This retry policy does not retry.");
            if (retryAttempt < 0 || retryAttempt >= MaxRetryCount)
                throw new ArgumentOutOfRangeException(nameof(retryAttempt));

            double milliseconds = InitialDelay.TotalMilliseconds * Math.Pow(BackoffFactor, retryAttempt);
            milliseconds = Math.Min(milliseconds, MaxDelay.TotalMilliseconds);

            if (UseJitter)
            {
                lock (JitterLock)
                {
                    milliseconds *= 0.5 + JitterRandom.NextDouble() * 0.5;
                }
            }

            return TimeSpan.FromMilliseconds(milliseconds);
        }
    }

    /// <summary>
    /// Identifies a future Quick Access lock target.
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
    /// Configures a future unlock operation.
    /// </summary>
    public sealed class QuickAccessUnlockOptions
    {
        /// <summary>
        /// Gets or sets whether new Recent shortcut files should be deleted during unlock.
        /// </summary>
        public bool CleanupNewRecentLinks { get; set; }
    }

    /// <summary>
    /// Describes a future Quick Access unlock report.
    /// </summary>
    public sealed class QuickAccessUnlockReport
    {
        internal QuickAccessUnlockReport(
            string recentFolder,
            IEnumerable<string> initialShortcutPaths,
            IEnumerable<string> currentShortcutPaths,
            IEnumerable<string> deletedShortcutPaths,
            IEnumerable<QuickAccessUnlockFailure> failedShortcutDeletions)
        {
            RecentFolder = recentFolder;
            InitialShortcutPaths = (initialShortcutPaths ?? Enumerable.Empty<string>()).ToList().AsReadOnly();
            CurrentShortcutPaths = (currentShortcutPaths ?? Enumerable.Empty<string>()).ToList().AsReadOnly();
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
        /// Gets the shortcuts deleted during cleanup.
        /// </summary>
        public IReadOnlyList<string> DeletedShortcutPaths { get; }

        /// <summary>
        /// Gets failed shortcut deletions.
        /// </summary>
        public IReadOnlyList<QuickAccessUnlockFailure> FailedShortcutDeletions { get; }
    }

    /// <summary>
    /// Describes a failed future shortcut deletion.
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
