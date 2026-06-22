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
            Diagnostics = Array.Empty<Diagnostic>();
        }

        /// <summary>Gets the DestList version.</summary>
        public uint Version { get; internal set; }

        /// <summary>Gets the declared entry count.</summary>
        public int DeclaredEntryCount { get; internal set; }

        /// <summary>Gets the pinned entry count.</summary>
        public uint PinnedEntryCount { get; internal set; }

        /// <summary>Gets the raw counter at DestList header offset 0x0c.</summary>
        public uint HeaderCounterRaw { get; internal set; }

        /// <summary>Gets the raw counter at DestList header offset 0x0c interpreted as a single-precision float.</summary>
        public float HeaderCounterF32 { get; internal set; }

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

        /// <summary>Gets the add/delete action count stored at DestList header offset 0x18.</summary>
        public ulong AddDeleteActionCount { get; internal set; }

        /// <summary>Gets entries.</summary>
        public IReadOnlyList<DestListEntry> Entries { get; internal set; }

        /// <summary>Gets non-fatal parse diagnostics.</summary>
        public IReadOnlyList<Diagnostic> Diagnostics { get; internal set; }
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

        /// <summary>Gets the physical entry position in the DestList stream.</summary>
        public int MruPosition { get; internal set; }

        /// <summary>Gets the entry identifier.</summary>
        public ulong EntryId { get; internal set; }

        /// <summary>Gets the entry number.</summary>
        public uint EntryNumber { get; internal set; }

        /// <summary>Gets the high 32 bits of the entry identifier.</summary>
        public uint EntryNumberUnknown { get; internal set; }

        /// <summary>Gets the reserved entry number field.</summary>
        public uint EntryNumberReserved { get; internal set; }

        /// <summary>Gets the hostname stored in the DestList entry.</summary>
        public string Hostname { get; internal set; }

        /// <summary>Gets the volume DROID GUID stored in the DestList entry, formatted from little-endian bytes.</summary>
        public string VolumeDroid { get; internal set; }

        /// <summary>Gets the file DROID GUID stored in the DestList entry, formatted from little-endian bytes.</summary>
        public string FileDroid { get; internal set; }

        /// <summary>Gets the volume birth DROID GUID stored in the DestList entry, formatted from little-endian bytes.</summary>
        public string VolumeBirthDroid { get; internal set; }

        /// <summary>Gets the file birth DROID GUID stored in the DestList entry, formatted from little-endian bytes.</summary>
        public string FileBirthDroid { get; internal set; }

        /// <summary>Gets the lower-case MAC address encoded in the last six bytes of the file DROID GUID.</summary>
        public string FileDroidMac { get; internal set; }

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

        /// <summary>Gets the compatibility alias for <see cref="AccessCount"/>.</summary>
        public uint Count { get; internal set; }

        /// <summary>Gets the score.</summary>
        public float Score { get; internal set; }

        /// <summary>Gets the raw Windows FILETIME value for the last access time.</summary>
        public ulong? LastAccessFileTime { get; internal set; }

        /// <summary>Gets the raw Windows FILETIME value for the last interaction time.</summary>
        public ulong? LastInteractionFileTime { get; internal set; }

        /// <summary>Gets the last access time.</summary>
        public DateTimeOffset? LastAccessTime { get; internal set; }

        /// <summary>Gets the last interaction time.</summary>
        public DateTimeOffset? LastInteractionTime { get; internal set; }

        /// <summary>Gets the serialized property store size.</summary>
        public uint? SerializedPropertyStoreSize { get; internal set; }

        /// <summary>Gets the reserved field at entry offset 0x78 for v3/v4/v6 entries.</summary>
        public uint? Reserved78 { get; internal set; }

        /// <summary>Gets the reserved field at entry offset 0x7c for v3/v4/v6 entries.</summary>
        public uint? Reserved7C { get; internal set; }

        /// <summary>Gets path candidates observed while resolving this entry.</summary>
        public IReadOnlyList<PathSource> PathSources { get; internal set; } = Array.Empty<PathSource>();

        /// <summary>Gets non-fatal entry-specific parse diagnostics.</summary>
        public IReadOnlyList<Diagnostic> Warnings { get; internal set; } = Array.Empty<Diagnostic>();
    }

    /// <summary>
    /// Severity for a non-fatal DestList parse diagnostic.
    /// </summary>
    public enum DiagnosticSeverity
    {
        /// <summary>Informational parser note.</summary>
        Info,
        /// <summary>Non-fatal parse issue.</summary>
        Warning
    }

    /// <summary>
    /// Describes a non-fatal DestList parse diagnostic.
    /// </summary>
    public sealed class Diagnostic
    {
        private Diagnostic(DiagnosticSeverity severity, string context, string message)
        {
            Severity = severity;
            Context = context ?? string.Empty;
            Message = message ?? string.Empty;
        }

        /// <summary>Gets the diagnostic severity.</summary>
        public DiagnosticSeverity Severity { get; }

        /// <summary>Gets the parser context that emitted the diagnostic.</summary>
        public string Context { get; }

        /// <summary>Gets the human-readable diagnostic message.</summary>
        public string Message { get; }

        /// <summary>Creates an informational diagnostic.</summary>
        public static Diagnostic Info(string context, string message)
        {
            return new Diagnostic(DiagnosticSeverity.Info, context, message);
        }

        /// <summary>Creates a warning diagnostic.</summary>
        public static Diagnostic Warning(string context, string message)
        {
            return new Diagnostic(DiagnosticSeverity.Warning, context, message);
        }
    }

    /// <summary>
    /// Describes a path candidate observed while resolving a DestList entry.
    /// </summary>
    public sealed class PathSource
    {
        internal PathSource(string source, string value)
        {
            Source = source ?? string.Empty;
            Value = value ?? string.Empty;
        }

        /// <summary>Gets the source that produced this path candidate.</summary>
        public string Source { get; }

        /// <summary>Gets the path candidate value.</summary>
        public string Value { get; }
    }

    /// <summary>
    /// Explorer-oriented DestList entry helpers.
    /// </summary>
    public static class DestListEntries
    {
        /// <summary>Returns all parsed entries.</summary>
        public static IReadOnlyList<DestListEntry> Entries(DestList destList)
        {
            if (destList == null)
                throw new ArgumentNullException(nameof(destList));

            return destList.Entries.ToList().AsReadOnly();
        }

        /// <summary>Returns entries likely visible in Explorer Quick Access.</summary>
        public static IReadOnlyList<DestListEntry> QuickAccessEntries(DestList destList, int normalSlotCount)
        {
            if (destList == null)
                throw new ArgumentNullException(nameof(destList));

            switch (destList.Version)
            {
                case 4:
                    return QuickAccessEntriesV4(destList.Entries).AsReadOnly();
                case 6:
                    return QuickAccessEntriesV6(destList.Entries, normalSlotCount).AsReadOnly();
                default:
                    return destList.Entries.ToList().AsReadOnly();
            }
        }

        /// <summary>Returns entries likely visible in Explorer Quick Access using four normal slots.</summary>
        public static IReadOnlyList<DestListEntry> VisibleEntries(DestList destList)
        {
            return QuickAccessEntries(destList, 4);
        }

        private static List<DestListEntry> QuickAccessEntriesV6(IReadOnlyList<DestListEntry> entries, int normalSlotCount)
        {
            var pinned = entries
                .Where(entry => entry.PinOrder.HasValue)
                .OrderBy(entry => entry.PinOrder ?? int.MaxValue)
                .ToList();
            var usedPaths = new HashSet<string>(pinned.Select(entry => VisibleEntryPathKey(entry.Path)), StringComparer.Ordinal);

            var normal = entries
                .Where(entry => !entry.PinOrder.HasValue && entry.RecentRank >= 0 && entry.RecentRank < normalSlotCount)
                .OrderByDescending(entry => entry.RecentRank)
                .Where(entry => usedPaths.Add(VisibleEntryPathKey(entry.Path)))
                .ToList();

            pinned.AddRange(normal);
            return pinned;
        }

        private static List<DestListEntry> QuickAccessEntriesV4(IReadOnlyList<DestListEntry> entries)
        {
            return entries.Any(entry => entry.PinOrder.HasValue)
                ? FrequentFolderEntriesV4(entries)
                : RecentFileEntriesV4(entries);
        }

        private static List<DestListEntry> FrequentFolderEntriesV4(IReadOnlyList<DestListEntry> entries)
        {
            var pinned = entries
                .Where(entry => entry.PinOrder.HasValue)
                .OrderBy(entry => entry.PinOrder ?? int.MaxValue)
                .ToList();
            var usedPaths = new HashSet<string>(pinned.Select(entry => VisibleEntryPathKey(entry.Path)), StringComparer.Ordinal);

            var candidates = entries
                .Where(entry => !entry.PinOrder.HasValue && entry.AccessCount > 1 && entry.RecentRank >= 0)
                .OrderBy(entry => entry.RecentRank)
                .ThenByDescending(entry => entry.LastInteractionFileTime ?? 0UL)
                .ThenByDescending(entry => entry.EntryNumber)
                .Where(entry => usedPaths.Add(VisibleEntryPathKey(entry.Path)))
                .Take(4)
                .ToList();

            pinned.AddRange(candidates);
            return pinned;
        }

        private static List<DestListEntry> RecentFileEntriesV4(IReadOnlyList<DestListEntry> entries)
        {
            var visible = new List<DestListEntry>();
            var usedPaths = new HashSet<string>(StringComparer.Ordinal);
            var backingFiles = new List<DestListEntry>();

            foreach (var entry in entries)
            {
                if (entry.AccessCount == 0)
                    continue;

                if (IsAutomaticDestinationsPath(entry.Path))
                {
                    backingFiles.Add(entry);
                    continue;
                }

                if (usedPaths.Add(VisibleEntryPathKey(entry.Path)))
                    visible.Add(entry);
            }

            var bestBackingFile = backingFiles
                .OrderBy(entry => entry.AccessCount)
                .ThenBy(entry => entry.LastInteractionFileTime ?? 0UL)
                .ThenBy(entry => entry.EntryNumber)
                .LastOrDefault();
            if (bestBackingFile != null && usedPaths.Add(VisibleEntryPathKey(bestBackingFile.Path)))
                visible.Add(bestBackingFile);

            return visible;
        }

        private static bool IsAutomaticDestinationsPath(string path)
        {
            return (path ?? string.Empty).EndsWith(".automaticDestinations-ms", StringComparison.OrdinalIgnoreCase);
        }

        private static string VisibleEntryPathKey(string path)
        {
            return (path ?? string.Empty)
                .Replace('/', '\\')
                .TrimEnd('\\')
                .ToUpperInvariant();
        }
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
