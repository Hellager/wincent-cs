using System;
using System.Collections.Generic;
using System.Linq;

namespace Wincent
{
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
}
