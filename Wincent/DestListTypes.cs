using System;
using System.Collections.Generic;
using System.Linq;

namespace Wincent
{
    /// <summary>
    /// Represents parsed automatic destinations metadata.
    /// </summary>
    public sealed class AutomaticDestinations
    {
        internal AutomaticDestinations(CfbInfo cfbInfo, DestList destList)
        {
            CfbInfo = cfbInfo;
            DestList = destList;
        }

        /// <summary>
        /// Gets compound file binary metadata.
        /// </summary>
        public CfbInfo CfbInfo { get; }

        /// <summary>
        /// Gets DestList metadata.
        /// </summary>
        public DestList DestList { get; }
    }

    /// <summary>
    /// Represents compound file binary metadata.
    /// </summary>
    public sealed class CfbInfo
    {
        internal CfbInfo(int sectorSize, int miniSectorSize, uint miniCutoffSize, IReadOnlyList<CfbDirectoryEntry> directoryEntries)
        {
            SectorSize = sectorSize;
            MiniSectorSize = miniSectorSize;
            MiniCutoffSize = miniCutoffSize;
            DirectoryEntries = directoryEntries ?? Array.Empty<CfbDirectoryEntry>();
        }

        /// <summary>Gets the sector size.</summary>
        public int SectorSize { get; }

        /// <summary>Gets the mini sector size.</summary>
        public int MiniSectorSize { get; }

        /// <summary>Gets the mini stream cutoff size.</summary>
        public uint MiniCutoffSize { get; }

        /// <summary>Gets directory entries.</summary>
        public IReadOnlyList<CfbDirectoryEntry> DirectoryEntries { get; }
    }

    /// <summary>
    /// Represents a compound file binary directory entry.
    /// </summary>
    public sealed class CfbDirectoryEntry
    {
        internal CfbDirectoryEntry(string name, byte rawObjectType, CfbObjectType objectType, uint startSector, ulong streamSize)
        {
            Name = name;
            RawObjectType = rawObjectType;
            ObjectType = objectType;
            StartSector = startSector;
            StreamSize = streamSize;
        }

        /// <summary>Gets the entry name.</summary>
        public string Name { get; }

        /// <summary>Gets the raw object type byte.</summary>
        public byte RawObjectType { get; }

        /// <summary>Gets the semantic object type.</summary>
        public CfbObjectType ObjectType { get; }

        /// <summary>Gets the starting sector.</summary>
        public uint StartSector { get; }

        /// <summary>Gets the stream size.</summary>
        public ulong StreamSize { get; }
    }

    /// <summary>
    /// Identifies a compound file binary object type.
    /// </summary>
    public enum CfbObjectType
    {
        /// <summary>Unknown object type.</summary>
        Unknown = 0,
        /// <summary>Storage object.</summary>
        Storage = 1,
        /// <summary>Stream object.</summary>
        Stream = 2,
        /// <summary>Root storage object.</summary>
        Root = 5
    }

    /// <summary>
    /// Represents DestList metadata.
    /// </summary>
    public sealed class DestList
    {
        internal DestList()
        {
            Entries = Array.Empty<DestListEntry>();
        }

        /// <summary>Gets the DestList version.</summary>
        public uint Version { get; internal set; }

        /// <summary>Gets the declared entry count.</summary>
        public int DeclaredEntryCount { get; internal set; }

        /// <summary>Gets the pinned entry count.</summary>
        public uint PinnedEntryCount { get; internal set; }

        /// <summary>Gets the last entry identifier.</summary>
        public ulong LastEntryId { get; internal set; }

        /// <summary>Gets the last entry number.</summary>
        public uint LastEntryNumber { get; internal set; }

        /// <summary>Gets the reserved last entry number field.</summary>
        public uint LastEntryNumberReserved { get; internal set; }

        /// <summary>Gets the last revision number.</summary>
        public uint LastRevisionNumber { get; internal set; }

        /// <summary>Gets the reserved last revision number field.</summary>
        public uint LastRevisionNumberReserved { get; internal set; }

        /// <summary>Gets entries.</summary>
        public IReadOnlyList<DestListEntry> Entries { get; internal set; }
    }

    /// <summary>
    /// Represents one DestList entry.
    /// </summary>
    public sealed class DestListEntry
    {
        internal DestListEntry()
        {
        }

        /// <summary>Gets the entry offset.</summary>
        public int EntryOffset { get; internal set; }

        /// <summary>Gets the entry length.</summary>
        public int EntryLength { get; internal set; }

        /// <summary>Gets the entry identifier.</summary>
        public ulong EntryId { get; internal set; }

        /// <summary>Gets the entry number.</summary>
        public uint EntryNumber { get; internal set; }

        /// <summary>Gets the reserved entry number field.</summary>
        public uint EntryNumberReserved { get; internal set; }

        /// <summary>Gets the stream name.</summary>
        public string StreamName { get; internal set; }

        /// <summary>Gets the raw path.</summary>
        public string RawPath { get; internal set; }

        /// <summary>Gets the normalized path.</summary>
        public string Path { get; internal set; }

        /// <summary>Gets the pin status.</summary>
        public int PinStatus { get; internal set; }

        /// <summary>Gets the pin order.</summary>
        public int? PinOrder { get; internal set; }

        /// <summary>Gets whether the entry is pinned.</summary>
        public bool IsPinned { get; internal set; }

        /// <summary>Gets the entry rank.</summary>
        public int Rank { get; internal set; }

        /// <summary>Gets the recent rank.</summary>
        public int RecentRank { get; internal set; }

        /// <summary>Gets the access count.</summary>
        public uint AccessCount { get; internal set; }

        /// <summary>Gets the score.</summary>
        public float Score { get; internal set; }

        /// <summary>Gets the last access time.</summary>
        public DateTimeOffset? LastAccessTime { get; internal set; }

        /// <summary>Gets the last interaction time.</summary>
        public DateTimeOffset? LastInteractionTime { get; internal set; }

        /// <summary>Gets the serialized property store size.</summary>
        public uint? SerializedPropertyStoreSize { get; internal set; }
    }

    /// <summary>
    /// Identifies an Explorer automatic destination family for experimental rebuild-based removal.
    /// </summary>
    public enum AutomaticDestinationsKind
    {
        /// <summary>
        /// Recent Files automatic destinations.
        /// </summary>
        RecentFiles,

        /// <summary>
        /// Frequent Folders automatic destinations.
        /// </summary>
        FrequentFolders
    }

    /// <summary>
    /// Configures experimental rebuild-based removal.
    /// </summary>
    public sealed class ExperimentalRemoveOptions
    {
        /// <summary>
        /// Gets or sets the initial delay after deleting the automatic destination file.
        /// </summary>
        /// <remarks>
        /// The implementation still polls for a rebuilt file after this delay.
        /// </remarks>
        public TimeSpan RebuildDelay { get; set; } = TimeSpan.FromMilliseconds(500);
    }

    /// <summary>
    /// Describes an experimental rebuild-based removal result.
    /// </summary>
    public sealed class ExperimentalRemoveReport
    {
        internal ExperimentalRemoveReport(
            AutomaticDestinationsKind kind,
            string recentFolder,
            string destinationPath,
            IEnumerable<string> requestedPaths,
            IEnumerable<string> matchingPathsBefore,
            IEnumerable<string> deletedShortcutPaths,
            IEnumerable<string> missingShortcutTargetPaths,
            bool destinationDeleted,
            bool rebuilt,
            TimeSpan? rebuildParseElapsed,
            string rebuildParseError,
            IEnumerable<string> remainingPathsAfterRebuild,
            bool success)
        {
            Kind = kind;
            RecentFolder = recentFolder;
            DestinationPath = destinationPath;
            RequestedPaths = (requestedPaths ?? Enumerable.Empty<string>()).ToList().AsReadOnly();
            MatchingPathsBefore = (matchingPathsBefore ?? Enumerable.Empty<string>()).ToList().AsReadOnly();
            DeletedShortcutPaths = (deletedShortcutPaths ?? Enumerable.Empty<string>()).ToList().AsReadOnly();
            MissingShortcutTargetPaths = (missingShortcutTargetPaths ?? Enumerable.Empty<string>()).ToList().AsReadOnly();
            DestinationDeleted = destinationDeleted;
            Rebuilt = rebuilt;
            RebuildParseElapsed = rebuildParseElapsed;
            RebuildParseError = rebuildParseError;
            RemainingPathsAfterRebuild = (remainingPathsAfterRebuild ?? Enumerable.Empty<string>()).ToList().AsReadOnly();
            Success = success;
        }

        /// <summary>Gets the automatic destination kind that was processed.</summary>
        public AutomaticDestinationsKind Kind { get; }

        /// <summary>Gets the current Windows Recent folder.</summary>
        public string RecentFolder { get; }

        /// <summary>Gets the automatic destination file path that was deleted and monitored.</summary>
        public string DestinationPath { get; }

        /// <summary>Gets target paths requested by the caller.</summary>
        public IReadOnlyList<string> RequestedPaths { get; }

        /// <summary>Gets DestList paths that matched before deletion started.</summary>
        public IReadOnlyList<string> MatchingPathsBefore { get; }

        /// <summary>Gets Recent shortcut files deleted during removal.</summary>
        public IReadOnlyList<string> DeletedShortcutPaths { get; }

        /// <summary>Gets requested target paths that had no matching Recent shortcut file.</summary>
        public IReadOnlyList<string> MissingShortcutTargetPaths { get; }

        /// <summary>Gets whether the automatic destination file was deleted.</summary>
        public bool DestinationDeleted { get; }

        /// <summary>Gets whether Explorer rebuilt the automatic destination file during polling.</summary>
        public bool Rebuilt { get; }

        /// <summary>Gets the time spent waiting until the rebuilt file could be parsed.</summary>
        public TimeSpan? RebuildParseElapsed { get; }

        /// <summary>Gets the last parse error observed while waiting for the rebuilt file.</summary>
        public string RebuildParseError { get; }

        /// <summary>Gets matching paths still present after Explorer rebuilt the file.</summary>
        public IReadOnlyList<string> RemainingPathsAfterRebuild { get; }

        /// <summary>Gets whether all requested entries were absent after rebuild.</summary>
        public bool Success { get; }
    }
}
