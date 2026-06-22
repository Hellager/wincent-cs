using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace Wincent
{
    internal interface IDestListMetadataReader
    {
        AutomaticDestinations ParseFile(string path);

        AutomaticDestinations ParseBytes(byte[] data);
    }

    internal sealed class DefaultDestListMetadataReader : IDestListMetadataReader
    {
        public AutomaticDestinations ParseFile(string path)
        {
            return DestListMetadataParser.ParseFile(path);
        }

        public AutomaticDestinations ParseBytes(byte[] data)
        {
            return DestListMetadataParser.ParseBytes(data);
        }
    }

    internal sealed class NoOpDestListMetadataReader : IDestListMetadataReader
    {
        public AutomaticDestinations ParseFile(string path)
        {
            throw new InvalidOperationException("DestList metadata is disabled for this instance.");
        }

        public AutomaticDestinations ParseBytes(byte[] data)
        {
            throw new InvalidOperationException("DestList metadata is disabled for this instance.");
        }
    }

    internal static class DestListMetadataParser
    {
        public static AutomaticDestinations ParseFile(string path)
        {
            if (path == null)
                throw new ArgumentNullException(nameof(path));

            return ParseBytes(File.ReadAllBytes(path), path);
        }

        public static AutomaticDestinations ParseBytes(byte[] data)
        {
            return ParseBytes(data, null);
        }

        private static AutomaticDestinations ParseBytes(byte[] data, string filePath)
        {
            if (data == null)
                throw new ArgumentNullException(nameof(data));

            CompoundFile cfb = CompoundFile.Parse(data, filePath);
            DestList destList = ParseDestList(cfb, filePath);
            var cfbInfo = new CfbInfo(
                cfb.SectorSize,
                cfb.MiniSectorSize,
                cfb.MiniCutoffSize,
                cfb.DirectoryEntries
                    .Select(entry => new CfbDirectoryEntry(
                        entry.Name,
                        entry.RawObjectType,
                        MapObjectType(entry.RawObjectType),
                        entry.StartSector,
                        entry.StreamSize))
                    .ToList());

            return new AutomaticDestinations(cfbInfo, destList);
        }

        internal static string RecentFilesDestPath()
        {
            return Path.Combine(
                new WindowsRecentFolder(new DefaultNativeMethods()).GetPath(),
                "AutomaticDestinations",
                "5f7b5f1e01b83767.automaticDestinations-ms");
        }

        internal static string FrequentFoldersDestPath()
        {
            return Path.Combine(
                new WindowsRecentFolder(new DefaultNativeMethods()).GetPath(),
                "AutomaticDestinations",
                "f01b4d95cf55d32a.automaticDestinations-ms");
        }

        private static DestList ParseDestList(CompoundFile cfb, string filePath)
        {
            byte[] destList = cfb.Stream("DestList");
            if (destList == null)
                throw new DestListParseException(filePath, null, "DestList stream not found.");

            if (destList.Length < 32)
            {
                return new DestList
                {
                    Version = 0,
                    DeclaredEntryCount = 0,
                    PinnedEntryCount = 0,
                    HeaderCounterRaw = 0,
                    HeaderCounterF32 = 0.0f,
                    LastEntryId = 0,
                    LastEntryNumber = 0,
                    LastEntryNumberReserved = 0,
                    LastRevisionNumber = 0,
                    LastRevisionNumberReserved = 0,
                    AddDeleteActionCount = 0,
                    Entries = Array.Empty<DestListEntry>(),
                    Diagnostics = new[]
                    {
                        Diagnostic.Warning("destlist", $"DestList stream is too small: {destList.Length} bytes")
                    }
                };
            }

            uint version = ReadUInt32(destList, 0, filePath);
            if (version != 1 && version != 3 && version != 4 && version != 6)
                throw new DestListUnsupportedVersionException(filePath, 0, version);

            int declaredEntryCount = ToInt32Safe(ReadUInt32(destList, 4, filePath), filePath, 4);
            uint pinnedEntryCount = ReadUInt32(destList, 8, filePath);
            uint headerCounterRaw = ReadUInt32(destList, 0x0c, filePath);
            float headerCounterF32 = BitConverter.ToSingle(BitConverter.GetBytes(headerCounterRaw), 0);
            ulong lastEntryId = ReadUInt64(destList, 0x10, filePath);
            uint lastEntryNumber = unchecked((uint)lastEntryId);
            uint lastEntryNumberReserved = ReadUInt32(destList, 0x14, filePath);
            ulong addDeleteActionCount = ReadUInt64(destList, 0x18, filePath);

            int offset = 32;
            var entries = new List<DestListEntry>();
            var diagnostics = new List<Diagnostic>();
            for (int i = 0; i < declaredEntryCount; i++)
            {
                DestListEntry entry = ParseDestListEntry(cfb, destList, version, i, offset, filePath);
                if (entry == null)
                {
                    if (i == 0)
                    {
                        throw new DestListParseException(
                            filePath,
                            offset,
                            $"DestList truncated before first entry (declared {declaredEntryCount}).");
                    }

                    diagnostics.Add(Diagnostic.Warning(
                        "destlist.entry",
                        $"stopped parsing at entry {i}, offset 0x{offset:x}; declared {declaredEntryCount}, parsed {entries.Count}"));
                    break;
                }

                offset = checked(offset + entry.EntryLength);
                entries.Add(entry);
            }

            return new DestList
            {
                Version = version,
                DeclaredEntryCount = declaredEntryCount,
                PinnedEntryCount = pinnedEntryCount,
                HeaderCounterRaw = headerCounterRaw,
                HeaderCounterF32 = headerCounterF32,
                LastEntryId = lastEntryId,
                LastEntryNumber = lastEntryNumber,
                LastEntryNumberReserved = lastEntryNumberReserved,
                LastRevisionNumber = unchecked((uint)addDeleteActionCount),
                LastRevisionNumberReserved = unchecked((uint)(addDeleteActionCount >> 32)),
                AddDeleteActionCount = addDeleteActionCount,
                Entries = entries.AsReadOnly(),
                Diagnostics = diagnostics.AsReadOnly()
            };
        }

        private static DestListEntry ParseDestListEntry(
            CompoundFile cfb,
            byte[] destList,
            uint version,
            int mruPosition,
            int offset,
            string filePath)
        {
            if (version == 1)
                return ParseDestListEntryV1(cfb, destList, mruPosition, offset, filePath);

            return ParseDestListEntryV2OrLater(cfb, destList, mruPosition, offset, filePath);
        }

        private static DestListEntry ParseDestListEntryV1(
            CompoundFile cfb,
            byte[] destList,
            int mruPosition,
            int offset,
            string filePath)
        {
            if (offset + 0x72 > destList.Length)
                return null;

            uint entryNumber = ReadUInt32(destList, offset + 0x58, filePath);
            uint entryNumberReserved = ReadUInt32(destList, offset + 0x5c, filePath);
            float score = BitConverter.ToSingle(BitConverter.GetBytes(ReadUInt32(destList, offset + 0x60, filePath)), 0);
            ulong? lastInteractionFiletime = ReadOptionalUInt64(destList, offset + 0x64);
            int pinStatus = ReadInt32(destList, offset + 0x6c, filePath);
            int pathChars = checked((int)ReadUInt16(destList, offset + 0x70, filePath));
            int pathStart = offset + 0x72;
            int pathEnd = checked(pathStart + pathChars * 2);
            if (pathEnd > destList.Length)
                return null;

            return BuildEntry(
                cfb,
                destList,
                offset,
                pathEnd - offset,
                mruPosition,
                entryNumber,
                entryNumberReserved,
                SliceBytes(destList, pathStart, pathEnd - pathStart),
                pinStatus,
                -1,
                0,
                score,
                lastInteractionFiletime,
                null,
                null,
                null,
                filePath);
        }

        private static DestListEntry ParseDestListEntryV2OrLater(
            CompoundFile cfb,
            byte[] destList,
            int mruPosition,
            int offset,
            string filePath)
        {
            if (offset + 0x82 > destList.Length)
                return null;

            uint entryNumber = ReadUInt32(destList, offset + 0x58, filePath);
            uint entryNumberReserved = ReadUInt32(destList, offset + 0x5c, filePath);
            float score = BitConverter.ToSingle(BitConverter.GetBytes(ReadUInt32(destList, offset + 0x60, filePath)), 0);
            ulong? lastInteractionFiletime = ReadOptionalUInt64(destList, offset + 0x64);
            int pinStatus = ReadInt32(destList, offset + 0x6c, filePath);
            int recentRank = ReadInt32(destList, offset + 0x70, filePath);
            uint accessCount = ReadUInt32(destList, offset + 0x74, filePath);
            uint? reserved78 = ReadOptionalUInt32(destList, offset + 0x78);
            uint? reserved7c = ReadOptionalUInt32(destList, offset + 0x7c);
            int pathChars = checked((int)ReadUInt16(destList, offset + 0x80, filePath));
            int pathStart = offset + 0x82;
            int pathEnd = checked(pathStart + pathChars * 2);
            if (pathEnd > destList.Length || pathEnd + 4 > destList.Length)
                return null;

            uint spsSize = ReadUInt32(destList, pathEnd, filePath);
            int spsSizeInt = ToInt32Safe(spsSize, filePath, offset);
            int entryEnd;
            try
            {
                entryEnd = checked(pathEnd + 4 + spsSizeInt);
            }
            catch (OverflowException)
            {
                throw new DestListParseException(
                    filePath,
                    offset,
                    $"Entry end overflow: pathEnd={pathEnd}, spsSize={spsSizeInt}.");
            }
            if (entryEnd > destList.Length)
                return null;

            return BuildEntry(
                cfb,
                destList,
                offset,
                entryEnd - offset,
                mruPosition,
                entryNumber,
                entryNumberReserved,
                SliceBytes(destList, pathStart, pathEnd - pathStart),
                pinStatus,
                recentRank,
                accessCount,
                score,
                lastInteractionFiletime,
                spsSize,
                reserved78,
                reserved7c,
                filePath);
        }

        private static DestListEntry BuildEntry(
            CompoundFile cfb,
            byte[] destList,
            int entryOffset,
            int entryLength,
            int mruPosition,
            uint entryNumber,
            uint entryNumberReserved,
            byte[] rawPathBytes,
            int pinStatus,
            int recentRank,
            uint accessCount,
            float score,
            ulong? lastInteractionFiletime,
            uint? spsSize,
            uint? reserved78,
            uint? reserved7c,
            string filePath)
        {
            string rawPath = DecodeUtf16Lossy(rawPathBytes);
            string streamName = entryNumber.ToString("x", CultureInfo.InvariantCulture);
            ResolvedPath resolved = ResolvePath(cfb, streamName, rawPath);
            DateTimeOffset? lastInteractionTime = FileTimeToDateTimeOffset(lastInteractionFiletime);

            return new DestListEntry
            {
                EntryOffset = entryOffset,
                EntryLength = entryLength,
                MruPosition = mruPosition,
                EntryId = ((ulong)entryNumberReserved << 32) | entryNumber,
                EntryNumber = entryNumber,
                EntryNumberUnknown = entryNumberReserved,
                EntryNumberReserved = entryNumberReserved,
                Hostname = DecodeHostname(destList, entryOffset + 0x48),
                VolumeDroid = FormatGuidFromLittleEndianBytes(destList, entryOffset + 0x08),
                FileDroid = FormatGuidFromLittleEndianBytes(destList, entryOffset + 0x18),
                VolumeBirthDroid = FormatGuidFromLittleEndianBytes(destList, entryOffset + 0x28),
                FileBirthDroid = FormatGuidFromLittleEndianBytes(destList, entryOffset + 0x38),
                FileDroidMac = MacFromDroidBytes(destList, entryOffset + 0x18),
                StreamName = streamName,
                RawPath = rawPath,
                Path = resolved.BestPath,
                PinStatus = pinStatus,
                PinOrder = pinStatus >= 0 ? (int?)pinStatus : null,
                IsPinned = pinStatus >= 0,
                Rank = recentRank,
                RecentRank = recentRank,
                Count = accessCount,
                AccessCount = accessCount,
                Score = score,
                LastAccessFileTime = lastInteractionFiletime,
                LastInteractionFileTime = lastInteractionFiletime,
                LastAccessTime = lastInteractionTime,
                LastInteractionTime = lastInteractionTime,
                SerializedPropertyStoreSize = spsSize,
                Reserved78 = reserved78,
                Reserved7C = reserved7c,
                PathSources = resolved.PathSources.AsReadOnly(),
                Warnings = resolved.Warnings.AsReadOnly()
            };
        }

        private static ResolvedPath ResolvePath(CompoundFile cfb, string streamName, string rawPath)
        {
            var pathSources = new List<PathSource>
            {
                new PathSource("destlist.raw_path", rawPath)
            };
            var warnings = new List<Diagnostic>();
            string safeRawPath = rawPath ?? string.Empty;
            bool needsLinkResolution =
                StartsWithKnownFolder(safeRawPath) ||
                safeRawPath.StartsWith("::", StringComparison.Ordinal);

            if (!needsLinkResolution)
                return new ResolvedPath(safeRawPath, pathSources, warnings);

            byte[] linkBytes = cfb.Stream(streamName);
            if (linkBytes == null)
            {
                warnings.Add(Diagnostic.Warning("destlist.link", $"missing Shell Link stream `{streamName}`"));
                warnings.Add(Diagnostic.Warning("destlist.path", $"could not resolve Shell Link path for `{safeRawPath}`"));
                return new ResolvedPath(safeRawPath, pathSources, warnings);
            }

            string linkPath = ParseLnkLocalPath(linkBytes);
            if (!string.IsNullOrWhiteSpace(linkPath))
            {
                pathSources.Add(new PathSource("lnk.linkinfo.local", linkPath));
                return new ResolvedPath(linkPath, pathSources, warnings);
            }

            warnings.Add(Diagnostic.Warning("destlist.path", $"could not resolve Shell Link path for `{safeRawPath}`"));
            return new ResolvedPath(safeRawPath, pathSources, warnings);
        }

        private static bool StartsWithKnownFolder(string rawPath)
        {
            return rawPath != null &&
                   rawPath.StartsWith("knownfolder:", StringComparison.OrdinalIgnoreCase);
        }

        private sealed class ResolvedPath
        {
            public ResolvedPath(string bestPath, List<PathSource> pathSources, List<Diagnostic> warnings)
            {
                BestPath = bestPath ?? string.Empty;
                PathSources = pathSources ?? new List<PathSource>();
                Warnings = warnings ?? new List<Diagnostic>();
            }

            public string BestPath { get; }

            public List<PathSource> PathSources { get; }

            public List<Diagnostic> Warnings { get; }
        }

        private static CfbObjectType MapObjectType(byte rawObjectType)
        {
            switch (rawObjectType)
            {
                case 1:
                    return CfbObjectType.Storage;
                case 2:
                    return CfbObjectType.Stream;
                case 5:
                    return CfbObjectType.Root;
                default:
                    return CfbObjectType.Unknown;
            }
        }

        private static DateTimeOffset? FileTimeToDateTimeOffset(ulong? filetime)
        {
            if (!filetime.HasValue)
                return null;

            try
            {
                return new DateTimeOffset(DateTime.FromFileTimeUtc(checked((long)filetime.Value)), TimeSpan.Zero);
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }
        }

        private static string ParseLnkLocalPath(byte[] data)
        {
            if (data == null || data.Length < 0x4c || ReadUInt32(data, 0, null) != 0x4c)
                return null;

            uint flags = ReadUInt32(data, 0x14, null);
            int offset = 0x4c;

            if ((flags & 0x1) != 0)
            {
                int idListSize = checked((int)ReadUInt16(data, offset, null));
                offset = checked(offset + 2 + idListSize);
            }

            if ((flags & 0x2) == 0 || offset + 28 > data.Length)
                return null;

            int linkInfoStart = offset;
            int linkInfoSize = ToInt32Safe(ReadUInt32(data, linkInfoStart, null), null, linkInfoStart);
            int linkInfoHeaderSize = ToInt32Safe(ReadUInt32(data, linkInfoStart + 4, null), null, linkInfoStart + 4);
            int linkInfoEnd = CheckedAdd(linkInfoStart, linkInfoSize, null, linkInfoStart);
            if (linkInfoSize < 28 || linkInfoEnd > data.Length)
                return null;

            int localBaseOffset = ToInt32Safe(ReadUInt32(data, linkInfoStart + 16, null), null, linkInfoStart + 16);
            int commonSuffixOffset = ToInt32Safe(ReadUInt32(data, linkInfoStart + 24, null), null, linkInfoStart + 24);
            int localBaseUnicodeOffset = linkInfoHeaderSize >= 0x24 && linkInfoStart + 32 <= data.Length
                ? ToInt32Safe(ReadUInt32(data, linkInfoStart + 28, null), null, linkInfoStart + 28)
                : 0;
            int commonSuffixUnicodeOffset = linkInfoHeaderSize >= 0x24 && linkInfoStart + 36 <= data.Length
                ? ToInt32Safe(ReadUInt32(data, linkInfoStart + 32, null), null, linkInfoStart + 32)
                : 0;

            string basePath = ReadUtf16ZStringInLinkInfo(data, linkInfoStart, linkInfoSize, localBaseUnicodeOffset)
                ?? ReadCStringInLinkInfo(data, linkInfoStart, linkInfoSize, localBaseOffset);
            string suffixPath = ReadUtf16ZStringInLinkInfo(data, linkInfoStart, linkInfoSize, commonSuffixUnicodeOffset)
                ?? ReadCStringInLinkInfo(data, linkInfoStart, linkInfoSize, commonSuffixOffset);

            if (!string.IsNullOrEmpty(basePath) && !string.IsNullOrEmpty(suffixPath))
                return JoinWindowsPath(basePath, suffixPath);

            if (LooksLikeWindowsPath(basePath))
                return basePath;

            if (LooksLikeWindowsPath(suffixPath))
                return suffixPath;

            return null;
        }

        private static string ReadCStringInLinkInfo(byte[] data, int linkInfoStart, int linkInfoSize, int relativeOffset)
        {
            if (relativeOffset == 0 || relativeOffset >= linkInfoSize)
                return null;

            int absoluteOffset = CheckedAdd(linkInfoStart, relativeOffset, null, linkInfoStart);
            int linkInfoEnd = CheckedAdd(linkInfoStart, linkInfoSize, null, linkInfoStart);
            return ReadCString(data, absoluteOffset, linkInfoEnd);
        }

        private static string ReadUtf16ZStringInLinkInfo(byte[] data, int linkInfoStart, int linkInfoSize, int relativeOffset)
        {
            if (relativeOffset == 0 || relativeOffset >= linkInfoSize)
                return null;

            int absoluteOffset = CheckedAdd(linkInfoStart, relativeOffset, null, linkInfoStart);
            int linkInfoEnd = CheckedAdd(linkInfoStart, linkInfoSize, null, linkInfoStart);
            return ReadUtf16ZString(data, absoluteOffset, linkInfoEnd);
        }

        private static string JoinWindowsPath(string basePath, string suffixPath)
        {
            if (LooksLikeWindowsPath(suffixPath))
                return suffixPath;

            if (basePath.EndsWith("\\", StringComparison.Ordinal) || suffixPath.Length == 0)
                return basePath + suffixPath;

            return basePath + "\\" + suffixPath;
        }

        private static bool LooksLikeWindowsPath(string value)
        {
            if (string.IsNullOrEmpty(value) || value.Length < 3)
                return false;

            return value[1] == ':' && (value[2] == '\\' || value[2] == '/');
        }

        private static string ReadCString(byte[] data, int offset, int limit)
        {
            if (offset >= data.Length || offset >= limit)
                return null;

            int end = offset;
            while (end < data.Length && end < limit && data[end] != 0)
                end++;

            if (end >= data.Length || end >= limit)
                return null;

            return Encoding.ASCII.GetString(data, offset, end - offset);
        }

        private static string DecodeHostname(byte[] data, int offset)
        {
            if (data == null || offset < 0 || offset + 16 > data.Length)
                return string.Empty;

            string ascii = DecodeAsciiNulPadded(data, offset, 16);
            if (!string.IsNullOrEmpty(ascii))
                return ascii;

            return DecodeUtf16Lossy(data, offset, 16).TrimEnd('\0');
        }

        private static string DecodeAsciiNulPadded(byte[] data, int offset, int length)
        {
            int end = offset;
            int limit = Math.Min(data.Length, offset + length);
            while (end < limit && data[end] != 0)
                end++;

            return Encoding.UTF8.GetString(data, offset, end - offset).Trim();
        }

        private static string FormatGuidFromLittleEndianBytes(byte[] data, int offset)
        {
            if (data == null || offset < 0 || offset + 16 > data.Length)
                return string.Empty;

            return new Guid(SliceBytes(data, offset, 16)).ToString("D");
        }

        private static string MacFromDroidBytes(byte[] data, int offset)
        {
            if (data == null || offset < 0 || offset + 16 > data.Length)
                return string.Empty;

            return string.Join(
                ":",
                Enumerable.Range(offset + 10, 6)
                    .Select(index => data[index].ToString("x2", CultureInfo.InvariantCulture)));
        }

        private static string ReadUtf16ZString(byte[] data, int offset, int limit)
        {
            if (offset + 1 >= data.Length || offset + 1 >= limit)
                return null;

            int end = offset;
            while (end + 1 < data.Length && end + 1 < limit)
            {
                if (data[end] == 0 && data[end + 1] == 0)
                    break;
                end += 2;
            }

            if (end == offset || end + 1 >= data.Length || end + 1 >= limit)
                return null;

            return DecodeUtf16Lossy(data, offset, end - offset);
        }

        private static string DecodeUtf16Lossy(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
                return string.Empty;

            return DecodeUtf16Lossy(bytes, 0, bytes.Length);
        }

        private static string DecodeUtf16Lossy(byte[] bytes, int offset, int length)
        {
            if (bytes == null || length <= 0)
                return string.Empty;

            int usableLength = length - length % 2;
            if (usableLength == 0)
                return string.Empty;

            return Encoding.Unicode.GetString(bytes, offset, usableLength);
        }

        private static byte[] SliceBytes(byte[] data, int offset, int length)
        {
            var slice = new byte[length];
            Buffer.BlockCopy(data, offset, slice, 0, length);
            return slice;
        }

        private static ushort ReadUInt16(byte[] data, int offset, string filePath)
        {
            if (offset < 0 || offset + 2 > data.Length)
                throw new DestListParseException(filePath, offset, $"Unexpected end of data at offset {offset}.");

            return BitConverter.ToUInt16(data, offset);
        }

        private static uint ReadUInt32(byte[] data, int offset, string filePath)
        {
            if (offset < 0 || offset + 4 > data.Length)
                throw new DestListParseException(filePath, offset, $"Unexpected end of data at offset {offset}.");

            return BitConverter.ToUInt32(data, offset);
        }

        private static ulong ReadUInt64(byte[] data, int offset, string filePath)
        {
            if (offset < 0 || offset + 8 > data.Length)
                throw new DestListParseException(filePath, offset, $"Unexpected end of data at offset {offset}.");

            return BitConverter.ToUInt64(data, offset);
        }

        private static int ReadInt32(byte[] data, int offset, string filePath)
        {
            if (offset < 0 || offset + 4 > data.Length)
                throw new DestListParseException(filePath, offset, $"Unexpected end of data at offset {offset}.");

            return BitConverter.ToInt32(data, offset);
        }

        private static ulong? ReadOptionalUInt64(byte[] data, int offset)
        {
            if (offset < 0 || offset + 8 > data.Length)
                return null;

            return BitConverter.ToUInt64(data, offset);
        }

        private static uint? ReadOptionalUInt32(byte[] data, int offset)
        {
            if (offset < 0 || offset + 4 > data.Length)
                return null;

            return BitConverter.ToUInt32(data, offset);
        }

        private static int ToInt32Safe(ulong value, string filePath, long? offset)
        {
            try
            {
                return checked((int)value);
            }
            catch (OverflowException)
            {
                throw new DestListParseException(filePath, offset, $"Value too large for int32: {value}.");
            }
        }

        private static int ToInt32Safe(uint value, string filePath, long? offset)
        {
            try
            {
                return checked((int)value);
            }
            catch (OverflowException)
            {
                throw new DestListParseException(filePath, offset, $"Value too large for int32: {value}.");
            }
        }

        private static int CheckedAdd(int a, int b, string filePath, long? offset)
        {
            try
            {
                return checked(a + b);
            }
            catch (OverflowException)
            {
                throw new DestListParseException(filePath, offset, $"Addition overflow: {a} + {b}.");
            }
        }

        private static int SectorDataOffset(uint sectorId, int sectorSize)
        {
            int id = ToInt32Safe(sectorId, null, sectorId);
            try
            {
                return checked((id + 1) * sectorSize);
            }
            catch (OverflowException)
            {
                throw new DestListParseException(null, sectorId, $"Sector offset overflow for sector {sectorId} (sector size {sectorSize}).");
            }
        }

        private sealed class CompoundFile
        {
            private const uint EndOfChain = 0xFFFF_FFFE;
            private const uint FreeSector = 0xFFFF_FFFF;

            private readonly byte[] _data;
            private readonly List<DirectoryEntry> _directory;
            private readonly byte[] _rootStream;
            private readonly List<uint> _fat;
            private readonly List<uint> _miniFat;

            private CompoundFile(
                byte[] data,
                int sectorSize,
                int miniSectorSize,
                uint miniCutoffSize,
                List<uint> fat,
                List<uint> miniFat,
                List<DirectoryEntry> directory,
                byte[] rootStream)
            {
                _data = data;
                SectorSize = sectorSize;
                MiniSectorSize = miniSectorSize;
                MiniCutoffSize = miniCutoffSize;
                _fat = fat;
                _miniFat = miniFat;
                _directory = directory;
                _rootStream = rootStream;
            }

            public int SectorSize { get; }

            public int MiniSectorSize { get; }

            public uint MiniCutoffSize { get; }

            internal List<DirectoryEntry> DirectoryEntries => _directory;

            public static CompoundFile Parse(byte[] data, string filePath)
            {
                if (data.Length < 512)
                    throw new DestListParseException(filePath, null, "file is too small for a CFB header");

                byte[] magic = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
                if (!data.Take(8).SequenceEqual(magic))
                    throw new DestListParseException(filePath, null, "not an OLE Compound File Binary");

                int sectorSizeShift = checked((int)ReadUInt16(data, 0x1e, filePath));
                int sectorSize = 1 << sectorSizeShift;
                if (sectorSize != 512 && sectorSize != 4096)
                    throw new DestListParseException(filePath, null, "unsupported CFB sector size");

                int miniSectorSizeShift = checked((int)ReadUInt16(data, 0x20, filePath));
                int miniSectorSize = 1 << miniSectorSizeShift;
                if (miniSectorSize != 64)
                    throw new DestListParseException(filePath, null, "unsupported CFB mini sector size");

                uint firstDirSector = ReadUInt32(data, 0x30, filePath);
                uint miniCutoffSize = ReadUInt32(data, 0x38, filePath);
                uint firstMiniFatSector = ReadUInt32(data, 0x3c, filePath);
                uint numMiniFatSectors = ReadUInt32(data, 0x40, filePath);
                uint firstDifatSector = ReadUInt32(data, 0x44, filePath);
                uint numDifatSectors = ReadUInt32(data, 0x48, filePath);

                var fatSectorIds = ReadDifat(data, sectorSize, firstDifatSector, numDifatSectors);
                var fat = ReadFat(data, sectorSize, fatSectorIds);

                byte[] directoryStream = ReadRegularStream(data, sectorSize, fat, firstDirSector, filePath);
                List<DirectoryEntry> directory = ParseDirectory(directoryStream);
                DirectoryEntry root = directory.FirstOrDefault(entry => entry.ObjectType == 5);
                if (root == null)
                    throw new DestListParseException(filePath, null, "root storage entry not found");

                byte[] rootStream = ReadRegularStreamSized(
                    data,
                    sectorSize,
                    fat,
                    root.StartSector,
                    ToInt32Safe(root.StreamSize, filePath, null),
                    filePath);

                List<uint> miniFat;
                if (firstMiniFatSector == FreeSector || numMiniFatSectors == 0)
                {
                    miniFat = new List<uint>();
                }
                else
                {
                    int numSectors = ToInt32Safe(numMiniFatSectors, filePath, 0x40);
                    int miniFatSize;
                    try
                    {
                        miniFatSize = checked(numSectors * sectorSize);
                    }
                    catch (OverflowException)
                    {
                        throw new DestListParseException(
                            filePath,
                            0x40,
                            $"Mini FAT size overflow: {numSectors} sectors * {sectorSize}.");
                    }

                    byte[] miniFatBytes = ReadRegularStreamSized(
                        data,
                        sectorSize,
                        fat,
                        firstMiniFatSector,
                        miniFatSize,
                        filePath);

                    miniFat = new List<uint>();
                    for (int i = 0; i + 3 < miniFatBytes.Length; i += 4)
                        miniFat.Add(BitConverter.ToUInt32(miniFatBytes, i));
                }

                return new CompoundFile(data, sectorSize, miniSectorSize, miniCutoffSize, fat, miniFat, directory, rootStream);
            }

            public byte[] Stream(string name)
            {
                DirectoryEntry entry = _directory.FirstOrDefault(
                    e => e.ObjectType == 2 && string.Equals(e.Name, name, StringComparison.OrdinalIgnoreCase));
                if (entry == null)
                    return null;

                int size = ToInt32Safe(entry.StreamSize, null, null);
                if (size < MiniCutoffSize)
                    return ReadMiniStream(entry.StartSector, size);

                return ReadRegularStreamSized(_data, SectorSize, _fat, entry.StartSector, size, null);
            }

            private byte[] ReadMiniStream(uint startSector, int size)
            {
                if (size == 0)
                    return Array.Empty<byte>();

                List<uint> chain = SectorChain(_miniFat, startSector);
                var outBytes = new List<byte>(size);
                foreach (uint miniSector in chain)
                {
                    int miniSectorInt = ToInt32Safe(miniSector, null, miniSector);
                    int offset;
                    try
                    {
                        offset = checked(miniSectorInt * MiniSectorSize);
                    }
                    catch (OverflowException)
                    {
                        throw new DestListParseException(
                            null,
                            miniSectorInt,
                            $"Mini sector offset overflow: {miniSectorInt} * {MiniSectorSize}.");
                    }
                    int end = CheckedAdd(offset, MiniSectorSize, null, miniSector);
                    if (offset < 0 || end > _rootStream.Length)
                        throw new DestListParseException(null, offset, $"mini sector {miniSector} is out of bounds");

                    byte[] sector = SliceBytes(_rootStream, offset, MiniSectorSize);
                    outBytes.AddRange(sector);
                    if (outBytes.Count >= size)
                        break;
                }

                if (outBytes.Count > size)
                    outBytes.RemoveRange(size, outBytes.Count - size);

                return outBytes.ToArray();
            }

            private static List<uint> ReadDifat(byte[] data, int sectorSize, uint firstDifatSector, uint numDifatSectors)
            {
                var fatSectorIds = new List<uint>();
                for (int index = 0; index < 109; index++)
                {
                    uint sectorId = BitConverter.ToUInt32(data, 0x4c + index * 4);
                    if (sectorId != FreeSector)
                        fatSectorIds.Add(sectorId);
                }

                uint nextDifatSector = firstDifatSector;
                for (int i = 0; i < numDifatSectors; i++)
                {
                    if (nextDifatSector == FreeSector || nextDifatSector == EndOfChain)
                        break;

                    byte[] sector = SectorSlice(data, sectorSize, nextDifatSector);
                    int entriesPerSector = sectorSize / 4;
                    for (int index = 0; index < entriesPerSector - 1; index++)
                    {
                        uint sectorId = BitConverter.ToUInt32(sector, index * 4);
                        if (sectorId != FreeSector)
                            fatSectorIds.Add(sectorId);
                    }

                    nextDifatSector = BitConverter.ToUInt32(sector, (entriesPerSector - 1) * 4);
                }

                return fatSectorIds;
            }

            private static List<uint> ReadFat(byte[] data, int sectorSize, List<uint> fatSectorIds)
            {
                var fat = new List<uint>();
                foreach (uint sectorId in fatSectorIds)
                {
                    byte[] sector = SectorSlice(data, sectorSize, sectorId);
                    for (int i = 0; i + 3 < sector.Length; i += 4)
                        fat.Add(BitConverter.ToUInt32(sector, i));
                }

                return fat;
            }

            private static List<uint> SectorChain(List<uint> fat, uint startSector)
            {
                if (startSector == FreeSector || startSector == EndOfChain)
                    return new List<uint>();

                var chain = new List<uint>();
                var seen = new HashSet<uint>();
                uint sector = startSector;

                while (true)
                {
                    if (sector == EndOfChain)
                        break;

                    if (sector == FreeSector)
                        throw new DestListParseException(null, null, $"invalid sector marker {sector:#x} in chain");

                    int index = ToInt32Safe(sector, null, sector);
                    if (index >= fat.Count)
                        throw new DestListParseException(null, null, $"sector {sector} is outside the FAT");

                    if (!seen.Add(sector))
                        throw new DestListParseException(null, null, $"loop detected in sector chain at sector {sector}");

                    chain.Add(sector);
                    sector = fat[index];
                }

                return chain;
            }

            private static byte[] ReadRegularStreamSized(
                byte[] data,
                int sectorSize,
                List<uint> fat,
                uint startSector,
                int size,
                string filePath)
            {
                byte[] stream = ReadRegularStream(data, sectorSize, fat, startSector, filePath);
                if (stream.Length > size)
                    Array.Resize(ref stream, size);
                return stream;
            }

            private static byte[] ReadRegularStream(
                byte[] data,
                int sectorSize,
                List<uint> fat,
                uint startSector,
                string filePath)
            {
                List<uint> chain = SectorChain(fat, startSector);
                var outBytes = new List<byte>();
                foreach (uint sectorId in chain)
                    outBytes.AddRange(StreamSectorSlice(data, sectorSize, sectorId));

                return outBytes.ToArray();
            }

            private static byte[] SectorSlice(byte[] data, int sectorSize, uint sectorId)
            {
                int offset = SectorDataOffset(sectorId, sectorSize);
                int end = CheckedAdd(offset, sectorSize, null, sectorId);
                if (offset < 0 || end > data.Length)
                    throw new DestListParseException(null, null, $"sector {sectorId} is out of bounds");

                return SliceBytes(data, offset, sectorSize);
            }

            private static byte[] StreamSectorSlice(byte[] data, int sectorSize, uint sectorId)
            {
                int offset = SectorDataOffset(sectorId, sectorSize);
                if (offset >= data.Length)
                    throw new DestListParseException(null, null, $"sector {sectorId} is out of bounds");

                int end = Math.Min(CheckedAdd(offset, sectorSize, null, sectorId), data.Length);
                return SliceBytes(data, offset, end - offset);
            }

            private static List<DirectoryEntry> ParseDirectory(byte[] directoryStream)
            {
                var entries = new List<DirectoryEntry>();
                for (int offset = 0; offset + 128 <= directoryStream.Length; offset += 128)
                {
                    byte objectType = directoryStream[offset + 66];
                    if (objectType == 0)
                        continue;

                    ushort nameLength = BitConverter.ToUInt16(directoryStream, offset + 64);
                    string name = string.Empty;
                    if (nameLength >= 2 && nameLength <= 64)
                    {
                        name = DecodeUtf16Lossy(directoryStream, offset, nameLength - 2);
                    }

                    entries.Add(new DirectoryEntry
                    {
                        Name = name,
                        RawObjectType = objectType,
                        StartSector = BitConverter.ToUInt32(directoryStream, offset + 116),
                        StreamSize = BitConverter.ToUInt64(directoryStream, offset + 120)
                    });
                }

                return entries;
            }

            internal sealed class DirectoryEntry
            {
                public string Name { get; set; }

                public byte RawObjectType { get; set; }

                public uint StartSector { get; set; }

                public ulong StreamSize { get; set; }

                public byte ObjectType => RawObjectType;
            }
        }
    }
}
