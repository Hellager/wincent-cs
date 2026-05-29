using System;
using System.Collections.Generic;

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
}
