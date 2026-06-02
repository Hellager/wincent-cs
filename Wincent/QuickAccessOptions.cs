using System;

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
        /// The cleanup enumerates Windows Recent folder shortcuts, resolves each shortcut target, and deletes shortcuts
        /// whose target matches the removed item using Windows path comparison semantics.
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
}
