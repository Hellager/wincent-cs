using System;

namespace Wincent
{
    /// <summary>
    /// Base class for Wincent-specific exceptions.
    /// </summary>
    public class WincentException : Exception
    {
        /// <summary>
        /// Initializes a new exception.
        /// </summary>
        public WincentException()
        {
        }

        /// <summary>
        /// Initializes a new exception with a message.
        /// </summary>
        public WincentException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new exception with a message and inner exception.
        /// </summary>
        public WincentException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }

    /// <summary>
    /// Thrown when adding an item that already exists.
    /// </summary>
    public sealed class QuickAccessItemAlreadyExistsException : WincentException
    {
        /// <summary>
        /// Initializes the exception.
        /// </summary>
        public QuickAccessItemAlreadyExistsException(string path, QuickAccess target)
            : base($"The item already exists in {target}: {path}")
        {
            Path = path;
            Target = target;
        }

        /// <summary>
        /// Gets the item path.
        /// </summary>
        public string Path { get; }

        /// <summary>
        /// Gets the target section.
        /// </summary>
        public QuickAccess Target { get; }
    }

    /// <summary>
    /// Thrown when removing an item that does not exist.
    /// </summary>
    public sealed class QuickAccessItemNotFoundException : WincentException
    {
        /// <summary>
        /// Initializes the exception.
        /// </summary>
        public QuickAccessItemNotFoundException(string path, QuickAccess target)
            : base($"The item was not found in {target}: {path}")
        {
            Path = path;
            Target = target;
        }

        /// <summary>
        /// Gets the item path.
        /// </summary>
        public string Path { get; }

        /// <summary>
        /// Gets the target section.
        /// </summary>
        public QuickAccess Target { get; }
    }

    /// <summary>
    /// Thrown when an operation does not support the requested Quick Access target.
    /// </summary>
    public sealed class UnsupportedQuickAccessOperationException : WincentException
    {
        /// <summary>
        /// Initializes the exception.
        /// </summary>
        public UnsupportedQuickAccessOperationException(QuickAccess target, string operation)
            : base($"{operation} does not support {target}.")
        {
            Target = target;
            Operation = operation;
        }

        /// <summary>
        /// Gets the unsupported target.
        /// </summary>
        public QuickAccess Target { get; }

        /// <summary>
        /// Gets the operation name.
        /// </summary>
        public string Operation { get; }
    }

    /// <summary>
    /// Represents an unclassified Quick Access operation failure.
    /// </summary>
    public sealed class QuickAccessOperationException : WincentException
    {
        /// <summary>
        /// Initializes the exception.
        /// </summary>
        public QuickAccessOperationException(string operation, QuickAccess target, int? hResultCode, Exception innerException)
            : base($"{operation} failed for {target}.", innerException)
        {
            Operation = operation;
            Target = target;
            HResultCode = hResultCode;
        }

        /// <summary>
        /// Gets the operation name.
        /// </summary>
        public string Operation { get; }

        /// <summary>
        /// Gets the target section.
        /// </summary>
        public QuickAccess Target { get; }

        /// <summary>
        /// Gets the native HRESULT when available.
        /// </summary>
        public int? HResultCode { get; }
    }

    /// <summary>
    /// Represents a failure while reading or writing Quick Access visibility settings.
    /// </summary>
    public sealed class QuickAccessVisibilityException : WincentException
    {
        /// <summary>
        /// Initializes the exception.
        /// </summary>
        public QuickAccessVisibilityException(
            string operation,
            QuickAccess target,
            string valueName,
            Exception innerException)
            : base($"Quick Access visibility operation {operation} failed for {target}.", innerException)
        {
            Operation = operation;
            Target = target;
            ValueName = valueName;
        }

        /// <summary>
        /// Gets the visibility operation name.
        /// </summary>
        public string Operation { get; }

        /// <summary>
        /// Gets the target section.
        /// </summary>
        public QuickAccess Target { get; }

        /// <summary>
        /// Gets the Explorer registry value name being processed.
        /// </summary>
        public string ValueName { get; }
    }

    /// <summary>
    /// Identifies a post-mutation step that failed after the requested Quick Access mutation succeeded.
    /// </summary>
    public enum QuickAccessPostMutationStep
    {
        /// <summary>Deleting Explorer's Recent Files backing data failed.</summary>
        DeleteRecentFilesBackingData,

        /// <summary>Refreshing open Explorer windows failed.</summary>
        RefreshExplorer
    }

    /// <summary>
    /// Thrown when an add or remove mutation succeeds but a requested post-mutation step fails.
    /// </summary>
    public sealed class QuickAccessPostMutationException : WincentException
    {
        /// <summary>
        /// Initializes the exception.
        /// </summary>
        public QuickAccessPostMutationException(
            string path,
            QuickAccess target,
            QuickAccessPostMutationStep step,
            Exception innerException)
            : base($"Quick Access post-mutation step {step} failed for {target}: {path}", innerException)
        {
            Path = path;
            Target = target;
            Step = step;
        }

        /// <summary>Gets the path whose mutation succeeded.</summary>
        public string Path { get; }

        /// <summary>Gets the mutated Quick Access section.</summary>
        public QuickAccess Target { get; }

        /// <summary>Gets the failed post-mutation step.</summary>
        public QuickAccessPostMutationStep Step { get; }
    }

    /// <summary>
    /// Thrown when a clear operation only partially succeeds.
    /// </summary>
    public sealed class PartialClearException : WincentException
    {
        /// <summary>
        /// Initializes the exception.
        /// </summary>
        public PartialClearException(bool recentFilesCleared, bool frequentFoldersCleared, Exception innerException)
            : base(
                  $"Quick Access clear partially succeeded (recent_files_cleared: {FormatBoolean(recentFilesCleared)}, frequent_folders_cleared: {FormatBoolean(frequentFoldersCleared)}).",
                  innerException)
        {
            RecentFilesCleared = recentFilesCleared;
            FrequentFoldersCleared = frequentFoldersCleared;
            SourceException = innerException;
        }

        /// <summary>
        /// Gets whether recent files were cleared.
        /// </summary>
        public bool RecentFilesCleared { get; }

        /// <summary>
        /// Gets whether frequent folders were cleared.
        /// </summary>
        public bool FrequentFoldersCleared { get; }

        /// <summary>
        /// Gets the underlying error that prevented the full clear from completing.
        /// </summary>
        public Exception SourceException { get; }

        /// <summary>
        /// Gets whether any Quick Access section was cleared before the failure.
        /// </summary>
        public bool HasPartialProgress => RecentFilesCleared || FrequentFoldersCleared;

        /// <summary>
        /// Gets whether no Quick Access section was cleared before the failure.
        /// </summary>
        public bool IsCompleteFailure => !HasPartialProgress;

        private static string FormatBoolean(bool value)
        {
            return value ? "true" : "false";
        }
    }

    /// <summary>
    /// Thrown when COM apartment initialization conflicts with the current thread.
    /// </summary>
    public sealed class ComApartmentMismatchException : WincentException
    {
        /// <summary>
        /// Initializes the exception.
        /// </summary>
        public ComApartmentMismatchException(int hResultCode)
            : base($"COM apartment initialization failed with HRESULT 0x{hResultCode:X8}.")
        {
            HResultCode = hResultCode;
        }

        /// <summary>
        /// Gets the HRESULT.
        /// </summary>
        public int HResultCode { get; }
    }

    /// <summary>
    /// Identifies a PowerShell-backed operation.
    /// </summary>
    public enum PowerShellOperation
    {
        /// <summary>Refreshes Explorer windows.</summary>
        RefreshExplorer,
        /// <summary>Queries all Quick Access items.</summary>
        QueryQuickAccess,
        /// <summary>Queries recent files.</summary>
        QueryRecentFiles,
        /// <summary>Queries frequent folders.</summary>
        QueryFrequentFolders,
        /// <summary>Adds a recent file.</summary>
        AddRecentFile,
        /// <summary>Removes a recent file.</summary>
        RemoveRecentFile,
        /// <summary>Pins a frequent folder.</summary>
        PinFrequentFolder,
        /// <summary>Unpins a frequent folder.</summary>
        UnpinFrequentFolder,
        /// <summary>Clears pinned folders.</summary>
        ClearPinnedFolders
    }

    /// <summary>
    /// Classifies a PowerShell execution failure.
    /// </summary>
    public enum PowerShellErrorKind
    {
        /// <summary>The process failed.</summary>
        ProcessFailed,
        /// <summary>The process timed out.</summary>
        Timeout,
        /// <summary>Access was denied.</summary>
        AccessDenied,
        /// <summary>Execution policy blocked the script.</summary>
        ExecutionPolicy,
        /// <summary>A required cmdlet was not found.</summary>
        CmdletNotFound
    }

    /// <summary>
    /// Thrown when a PowerShell-backed operation fails.
    /// </summary>
    public sealed class PowerShellExecutionException : WincentException
    {
        /// <summary>
        /// Initializes the exception.
        /// </summary>
        public PowerShellExecutionException(
            PowerShellOperation operation,
            PowerShellErrorKind kind,
            int? exitCode,
            string standardOutput,
            string standardError,
            string scriptPath,
            string parameters,
            TimeSpan? duration,
            int? nativeErrorCode,
            Exception innerException = null)
            : base($"PowerShell operation {operation} failed with {kind}.", innerException)
        {
            Operation = operation;
            Kind = kind;
            ExitCode = exitCode;
            StandardOutput = standardOutput ?? string.Empty;
            StandardError = standardError ?? string.Empty;
            ScriptPath = scriptPath;
            Parameters = parameters;
            Duration = duration;
            NativeErrorCode = nativeErrorCode;
        }

        /// <summary>Gets the operation.</summary>
        public PowerShellOperation Operation { get; }

        /// <summary>Gets the error kind.</summary>
        public PowerShellErrorKind Kind { get; }

        /// <summary>Gets the process exit code.</summary>
        public int? ExitCode { get; }

        /// <summary>Gets captured standard output.</summary>
        public string StandardOutput { get; }

        /// <summary>Gets captured standard error.</summary>
        public string StandardError { get; }

        /// <summary>Gets the script path when available.</summary>
        public string ScriptPath { get; }

        /// <summary>Gets script parameters when available.</summary>
        public string Parameters { get; }

        /// <summary>Gets operation duration when available.</summary>
        public TimeSpan? Duration { get; }

        /// <summary>Gets the native error code when available.</summary>
        public int? NativeErrorCode { get; }
    }

    /// <summary>
    /// Base class for DestList parsing failures.
    /// </summary>
    public class DestListParseException : WincentException
    {
        /// <summary>
        /// Initializes the exception.
        /// </summary>
        public DestListParseException(string filePath, long? offset, string details, Exception innerException = null)
            : base(details, innerException)
        {
            FilePath = filePath;
            Offset = offset;
            Details = details;
        }

        /// <summary>Gets the parsed file path.</summary>
        public string FilePath { get; }

        /// <summary>Gets the byte offset when available.</summary>
        public long? Offset { get; }

        /// <summary>Gets parser details.</summary>
        public string Details { get; }
    }

    /// <summary>
    /// Thrown when a DestList version is unsupported.
    /// </summary>
    public sealed class DestListUnsupportedVersionException : DestListParseException
    {
        /// <summary>
        /// Initializes the exception.
        /// </summary>
        public DestListUnsupportedVersionException(string filePath, long? offset, uint version, string details = null)
            : base(filePath, offset, details ?? "Unsupported DestList version.")
        {
            Version = version;
        }

        /// <summary>Gets the unsupported version.</summary>
        public uint Version { get; }
    }
}
