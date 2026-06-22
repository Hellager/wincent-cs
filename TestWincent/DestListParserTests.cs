using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Text;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class DestListParserTests
    {
        private static readonly byte[] VolumeDroidBytes =
        {
            0x00, 0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77,
            0x88, 0x99, 0xaa, 0xbb, 0xcc, 0xdd, 0xee, 0xff
        };

        private static readonly byte[] FileDroidBytes =
        {
            0x10, 0x32, 0x54, 0x76, 0x98, 0xba, 0xdc, 0xfe,
            0x00, 0x11, 0xaa, 0xbb, 0xcc, 0xdd, 0xee, 0x01
        };

        private static readonly byte[] VolumeBirthDroidBytes =
        {
            0x01, 0x23, 0x45, 0x67, 0x89, 0xab, 0xcd, 0xef,
            0x10, 0x20, 0x30, 0x40, 0x50, 0x60, 0x70, 0x80
        };

        private static readonly byte[] FileBirthDroidBytes =
        {
            0x89, 0x67, 0x45, 0x23, 0x01, 0xef, 0xcd, 0xab,
            0x98, 0x76, 0x54, 0x32, 0x10, 0x00, 0xff, 0xee
        };

        public TestContext TestContext { get; set; }

        [TestMethod]
        public void ParseBytes_RejectsWrongMagic()
        {
            var data = new byte[512];

            var exception = Assert.ThrowsException<DestListParseException>(
                () => DestListMetadataParser.ParseBytes(data));

            StringAssert.Contains(exception.Details, "OLE Compound File Binary");
        }

        [TestMethod]
        public void ParseBytes_RejectsTruncatedFile()
        {
            var data = new byte[100];

            Assert.ThrowsException<DestListParseException>(
                () => DestListMetadataParser.ParseBytes(data));
        }

        [TestMethod]
        public void ParseBytes_NullData_ThrowsArgumentNullException()
        {
            Assert.ThrowsException<ArgumentNullException>(
                () => DestListMetadataParser.ParseBytes(null));
        }

        [TestMethod]
        public void ParseBytes_ParsesMinimalVersion4DestList()
        {
            byte[] data = BuildMinimalCfbWithDestList(@"C:\Test\file.txt");

            var parsed = DestListMetadataParser.ParseBytes(data);
            var entry = parsed.DestList.Entries[0];

            Assert.AreEqual(4u, parsed.DestList.Version);
            Assert.AreEqual(1, parsed.DestList.DeclaredEntryCount);
            Assert.AreEqual(2u, parsed.DestList.PinnedEntryCount);
            Assert.AreEqual(0x3fc00000u, parsed.DestList.HeaderCounterRaw);
            Assert.AreEqual(1.5f, parsed.DestList.HeaderCounterF32);
            Assert.AreEqual(0x000000010000002aUL, parsed.DestList.LastEntryId);
            Assert.AreEqual(42u, parsed.DestList.LastEntryNumber);
            Assert.AreEqual(100UL, parsed.DestList.AddDeleteActionCount);
            Assert.AreEqual(42u, entry.EntryNumber);
            Assert.AreEqual(99u, entry.EntryNumberReserved);
            Assert.AreEqual(0x000000630000002aUL, entry.EntryId);
            Assert.AreEqual("HOST", entry.Hostname);
            Assert.AreEqual("33221100-5544-7766-8899-aabbccddeeff", entry.VolumeDroid);
            Assert.AreEqual("76543210-ba98-fedc-0011-aabbccddee01", entry.FileDroid);
            Assert.AreEqual("67452301-ab89-efcd-1020-304050607080", entry.VolumeBirthDroid);
            Assert.AreEqual("23456789-ef01-abcd-9876-54321000ffee", entry.FileBirthDroid);
            Assert.AreEqual("aa:bb:cc:dd:ee:01", entry.FileDroidMac);
            Assert.AreEqual("2a", entry.StreamName);
            Assert.AreEqual(@"C:\Test\file.txt", entry.RawPath);
            Assert.AreEqual(@"C:\Test\file.txt", entry.Path);
            Assert.AreEqual("destlist.raw_path", entry.PathSources.Single().Source);
            Assert.AreEqual(@"C:\Test\file.txt", entry.PathSources.Single().Value);
            Assert.AreEqual(0, entry.Warnings.Count);
            Assert.AreEqual(-1, entry.PinStatus);
            Assert.IsFalse(entry.PinOrder.HasValue);
            Assert.IsFalse(entry.IsPinned);
            Assert.AreEqual(7, entry.RecentRank);
            Assert.AreEqual(3u, entry.Count);
            Assert.AreEqual(3u, entry.AccessCount);
            Assert.AreEqual(1.5f, entry.Score);
            Assert.AreEqual(132537600000000000UL, entry.LastInteractionFileTime);
            Assert.AreEqual(0x11111111u, entry.Reserved78);
            Assert.AreEqual(0x22222222u, entry.Reserved7C);
            Assert.AreEqual(0u, entry.SerializedPropertyStoreSize);
            Assert.IsTrue(entry.LastInteractionTime.HasValue);
            Assert.AreEqual(0, parsed.DestList.Diagnostics.Count);
            Assert.IsTrue(parsed.CfbInfo.DirectoryEntries.Any(e => e.Name == "DestList" && e.ObjectType == CfbObjectType.Stream));
        }

        [TestMethod]
        public void ParseBytes_KnownFolderWithoutLinkStream_RecordsPathDiagnostics()
        {
            byte[] data = BuildMinimalCfbWithDestList("knownfolder:{guid}");

            var parsed = DestListMetadataParser.ParseBytes(data);
            var entry = parsed.DestList.Entries.Single();

            Assert.AreEqual("knownfolder:{guid}", entry.Path);
            Assert.AreEqual(1, entry.PathSources.Count);
            Assert.AreEqual("destlist.raw_path", entry.PathSources[0].Source);
            Assert.AreEqual(2, entry.Warnings.Count);
            Assert.IsTrue(entry.Warnings.Any(w => w.Context == "destlist.link"));
            Assert.IsTrue(entry.Warnings.Any(w => w.Context == "destlist.path"));
        }

        [TestMethod]
        public void ParseBytes_TruncatedAfterFirstEntry_RecordsDiagnostic()
        {
            byte[] data = BuildMinimalCfbWithDestList(
                new[]
                {
                    new DestListEntrySpec(@"C:\One.txt", 42, 0, -1, 0, 1, 1),
                    new DestListEntrySpec(@"C:\Two.txt", 43, 0, -1, 1, 1, 1)
                },
                declaredEntryCount: 3);

            var parsed = DestListMetadataParser.ParseBytes(data);

            Assert.AreEqual(3, parsed.DestList.DeclaredEntryCount);
            Assert.AreEqual(2, parsed.DestList.Entries.Count);
            Assert.AreEqual(1, parsed.DestList.Diagnostics.Count);
            Assert.AreEqual(DiagnosticSeverity.Warning, parsed.DestList.Diagnostics[0].Severity);
            Assert.AreEqual("destlist.entry", parsed.DestList.Diagnostics[0].Context);
        }

        [TestMethod]
        public void VisibleEntries_V4RecentFiles_FiltersHiddenDedupesAndKeepsBestBackingFile()
        {
            byte[] data = BuildMinimalCfbWithDestList(
                new[]
                {
                    new DestListEntrySpec(@"C:\Visible.txt", 1, 0, -1, 3, 2, 10),
                    new DestListEntrySpec(@"c:/visible.txt", 2, 0, -1, 2, 4, 11),
                    new DestListEntrySpec(@"C:\Hidden.txt", 3, 0, -1, 1, 0, 12),
                    new DestListEntrySpec(@"C:\Low.automaticDestinations-ms", 4, 0, -1, 4, 1, 1),
                    new DestListEntrySpec(@"C:\High.automaticDestinations-ms", 5, 0, -1, 5, 3, 3)
                });

            var visible = DestListEntries.VisibleEntries(DestListMetadataParser.ParseBytes(data).DestList);

            CollectionAssert.AreEqual(
                new[] { @"C:\Visible.txt", @"C:\High.automaticDestinations-ms" },
                visible.Select(entry => entry.Path).ToArray());
        }

        [TestMethod]
        public void VisibleEntries_V4FrequentFolders_OrdersPinnedAndTopFourFrequentCandidates()
        {
            byte[] data = BuildMinimalCfbWithDestList(
                new[]
                {
                    new DestListEntrySpec(@"C:\Pinned2", 1, 0, 1, 10, 1, 1),
                    new DestListEntrySpec(@"C:\Pinned1", 2, 0, 0, 9, 1, 1),
                    new DestListEntrySpec(@"C:\A", 3, 0, -1, 0, 2, 1),
                    new DestListEntrySpec(@"C:\B", 4, 0, -1, 1, 3, 2),
                    new DestListEntrySpec(@"C:\C", 5, 0, -1, 2, 4, 3),
                    new DestListEntrySpec(@"C:\D", 6, 0, -1, 3, 5, 4),
                    new DestListEntrySpec(@"C:\E", 7, 0, -1, 4, 6, 5),
                    new DestListEntrySpec(@"c:/a", 8, 0, -1, 0, 8, 6)
                });

            var visible = DestListEntries.VisibleEntries(DestListMetadataParser.ParseBytes(data).DestList);

            CollectionAssert.AreEqual(
                new[] { @"C:\Pinned1", @"C:\Pinned2", @"c:/a", @"C:\B", @"C:\C", @"C:\D" },
                visible.Select(entry => entry.Path).ToArray());
        }

        [TestMethod]
        public void VisibleEntries_V6_UsesPinnedThenReverseRankedNormalSlots()
        {
            byte[] data = BuildMinimalCfbWithDestList(
                new[]
                {
                    new DestListEntrySpec(@"C:\Pinned", 1, 0, 0, 10, 1, 1),
                    new DestListEntrySpec(@"C:\Rank0", 2, 0, -1, 0, 1, 1),
                    new DestListEntrySpec(@"C:\Rank1", 3, 0, -1, 1, 1, 1),
                    new DestListEntrySpec(@"C:\Rank4", 4, 0, -1, 4, 1, 1),
                    new DestListEntrySpec(@"c:/pinned", 5, 0, -1, 1, 1, 2)
                },
                version: 6);

            var visible = DestListEntries.VisibleEntries(DestListMetadataParser.ParseBytes(data).DestList);

            CollectionAssert.AreEqual(
                new[] { @"C:\Pinned", @"C:\Rank1", @"C:\Rank0" },
                visible.Select(entry => entry.Path).ToArray());
        }

        [TestMethod]
        [TestCategory("Integration")]
        [Ignore("Integration test; reads the current user's real Recent Files automaticDestinations file.")]
        public void Integration_ParseRecentFilesMetadata()
        {
            string path = DestListMetadataParser.RecentFilesDestPath();
            var parsed = DestListMetadataParser.ParseFile(path);

            Assert.IsNotNull(parsed.DestList);
            OutputDestList("Recent Files", path, parsed);
        }

        [TestMethod]
        [TestCategory("Integration")]
        [Ignore("Integration test; reads the current user's real Frequent Folders automaticDestinations file.")]
        public void Integration_ParseFrequentFoldersMetadata()
        {
            string path = DestListMetadataParser.FrequentFoldersDestPath();
            var parsed = DestListMetadataParser.ParseFile(path);

            Assert.IsNotNull(parsed.DestList);
            OutputDestList("Frequent Folders", path, parsed);
        }

        private void OutputDestList(string label, string path, AutomaticDestinations parsed)
        {
            var dest = parsed.DestList;
            var cfb = parsed.CfbInfo;

            TestContext.WriteLine("");
            TestContext.WriteLine($"=== {label} ===");
            TestContext.WriteLine($"  File          : {path}");
            TestContext.WriteLine($"  CFB sector    : {cfb.SectorSize} / mini {cfb.MiniSectorSize} / cutoff {cfb.MiniCutoffSize}");
            TestContext.WriteLine($"  Dir entries   : {cfb.DirectoryEntries.Count}");
            TestContext.WriteLine($"  Version       : {dest.Version}");
            TestContext.WriteLine($"  Entries       : {dest.DeclaredEntryCount} declared ({dest.Entries.Count} parsed)");
            TestContext.WriteLine($"  Pinned        : {dest.PinnedEntryCount}");
            TestContext.WriteLine($"  Last rev      : {dest.LastRevisionNumber}");

            for (int i = 0; i < dest.Entries.Count; i++)
            {
                var e = dest.Entries[i];
                TestContext.WriteLine(
                    $"  [{i}] {(e.IsPinned ? "PIN " : "    ")}{e.Path,-60} rank={e.RecentRank,4} access={e.AccessCount,4} score={e.Score:F2} pinOrder={e.PinOrder?.ToString() ?? "-"} time={e.LastInteractionTime:yyyy-MM-dd HH:mm}");
            }

            TestContext.WriteLine($"=== {dest.Entries.Count} entries total ===");
        }

        [TestMethod]
        public void ParseBytes_RejectsOversizedStreamSize_ThrowsDestListParseException()
        {
            byte[] data = BuildCfbWithInflatedStreamSize();

            var exception = Assert.ThrowsException<DestListParseException>(
                () => DestListMetadataParser.ParseBytes(data));

            StringAssert.Contains(exception.Details, "int32");
        }

        [TestMethod]
        public void ParseBytes_TruncatedDestListStream_ReturnsDiagnostic()
        {
            byte[] file = new byte[512 + 512 * 4];
            file[0] = 0xD0; file[1] = 0xCF; file[2] = 0x11; file[3] = 0xE0;
            file[4] = 0xA1; file[5] = 0xB1; file[6] = 0x1A; file[7] = 0xE1;
            WriteUInt16(file, 0x1e, 9);
            WriteUInt16(file, 0x20, 6);
            WriteUInt32(file, 0x30, 0);
            WriteUInt32(file, 0x38, 0);
            WriteUInt32(file, 0x3c, 0xFFFFFFFF);
            WriteUInt32(file, 0x40, 0);
            WriteUInt32(file, 0x44, 0xFFFFFFFF);
            WriteUInt32(file, 0x48, 0);
            for (int i = 0; i < 109; i++)
                WriteUInt32(file, 0x4c + i * 4, 0xFFFFFFFF);
            WriteUInt32(file, 0x4c, 3);

            int directoryOffset = 512;
            WriteDirectoryEntry(file, directoryOffset, "Root Entry", 5, 1, 0);
            WriteDirectoryEntry(file, directoryOffset + 128, "DestList", 2, 2, 10);

            byte[] shortStream = new byte[10];
            Array.Copy(shortStream, 0, file, 512 + 512 * 2, 10);

            int fatOffset = 512 + 512 * 3;
            for (int i = 0; i < 128; i++)
                WriteUInt32(file, fatOffset + i * 4, 0xFFFFFFFF);
            WriteUInt32(file, fatOffset, 0xFFFFFFFE);
            WriteUInt32(file, fatOffset + 4, 0xFFFFFFFE);
            WriteUInt32(file, fatOffset + 8, 0xFFFFFFFE);
            WriteUInt32(file, fatOffset + 12, 0xFFFFFFFD);

            var parsed = DestListMetadataParser.ParseBytes(file);

            Assert.AreEqual(0u, parsed.DestList.Version);
            Assert.AreEqual(0, parsed.DestList.Entries.Count);
            Assert.AreEqual(1, parsed.DestList.Diagnostics.Count);
            Assert.AreEqual(DiagnosticSeverity.Warning, parsed.DestList.Diagnostics[0].Severity);
            Assert.AreEqual("destlist", parsed.DestList.Diagnostics[0].Context);
            StringAssert.Contains(parsed.DestList.Diagnostics[0].Message, "too small");
        }

        private static byte[] BuildMinimalCfbWithDestList(string path)
        {
            byte[] destList = BuildDestList(path);
            return BuildMinimalCfbWithDestList(destList);
        }

        private static byte[] BuildMinimalCfbWithDestList(
            DestListEntrySpec[] entries,
            uint version = 4,
            int? declaredEntryCount = null)
        {
            byte[] destList = BuildDestList(entries, version, declaredEntryCount);
            return BuildMinimalCfbWithDestList(destList);
        }

        private static byte[] BuildMinimalCfbWithDestList(byte[] destList)
        {
            const int sectorSize = 512;
            int destListSectors = (destList.Length + sectorSize - 1) / sectorSize;
            uint directorySector = 0;
            uint rootSector = 1;
            uint destListStartSector = 2;
            uint fatSector = (uint)(2 + destListSectors);
            byte[] file = new byte[sectorSize + sectorSize * (int)(fatSector + 1)];

            file[0] = 0xD0;
            file[1] = 0xCF;
            file[2] = 0x11;
            file[3] = 0xE0;
            file[4] = 0xA1;
            file[5] = 0xB1;
            file[6] = 0x1A;
            file[7] = 0xE1;
            WriteUInt16(file, 0x1e, 9);
            WriteUInt16(file, 0x20, 6);
            WriteUInt32(file, 0x30, 0);
            WriteUInt32(file, 0x38, 0);
            WriteUInt32(file, 0x3c, 0xFFFFFFFF);
            WriteUInt32(file, 0x40, 0);
            WriteUInt32(file, 0x44, 0xFFFFFFFF);
            WriteUInt32(file, 0x48, 0);
            for (int i = 0; i < 109; i++)
                WriteUInt32(file, 0x4c + i * 4, 0xFFFFFFFF);
            WriteUInt32(file, 0x4c, fatSector);

            int directoryOffset = 512;
            WriteDirectoryEntry(file, directoryOffset, "Root Entry", 5, rootSector, 0);
            WriteDirectoryEntry(file, directoryOffset + 128, "DestList", 2, destListStartSector, (ulong)destList.Length);

            Array.Copy(destList, 0, file, 512 + 512 * 2, destList.Length);

            int fatOffset = 512 + 512 * (int)fatSector;
            for (int i = 0; i < 128; i++)
                WriteUInt32(file, fatOffset + i * 4, 0xFFFFFFFF);
            WriteUInt32(file, fatOffset + (int)directorySector * 4, 0xFFFFFFFE);
            WriteUInt32(file, fatOffset + (int)rootSector * 4, 0xFFFFFFFE);
            for (int i = 0; i < destListSectors; i++)
            {
                uint sector = destListStartSector + (uint)i;
                uint next = i == destListSectors - 1 ? 0xFFFFFFFEu : sector + 1;
                WriteUInt32(file, fatOffset + (int)sector * 4, next);
            }
            WriteUInt32(file, fatOffset + (int)fatSector * 4, 0xFFFFFFFD);

            return file;
        }

        private static byte[] BuildCfbWithInflatedStreamSize()
        {
            byte[] destList = BuildDestList(@"C:\Test\file.txt");
            byte[] file = new byte[512 + 512 * 4];

            file[0] = 0xD0; file[1] = 0xCF; file[2] = 0x11; file[3] = 0xE0;
            file[4] = 0xA1; file[5] = 0xB1; file[6] = 0x1A; file[7] = 0xE1;
            WriteUInt16(file, 0x1e, 9);
            WriteUInt16(file, 0x20, 6);
            WriteUInt32(file, 0x30, 0);
            WriteUInt32(file, 0x38, 0);
            WriteUInt32(file, 0x3c, 0xFFFFFFFF);
            WriteUInt32(file, 0x40, 0);
            WriteUInt32(file, 0x44, 0xFFFFFFFF);
            WriteUInt32(file, 0x48, 0);
            for (int i = 0; i < 109; i++)
                WriteUInt32(file, 0x4c + i * 4, 0xFFFFFFFF);
            WriteUInt32(file, 0x4c, 3);

            int directoryOffset = 512;
            WriteDirectoryEntry(file, directoryOffset, "Root Entry", 5, 1, 0);
            WriteDirectoryEntry(file, directoryOffset + 128, "DestList", 2, 2, ((ulong)int.MaxValue) + 1);

            Array.Copy(destList, 0, file, 512 + 512 * 2, destList.Length);

            int fatOffset = 512 + 512 * 3;
            for (int i = 0; i < 128; i++)
                WriteUInt32(file, fatOffset + i * 4, 0xFFFFFFFF);
            WriteUInt32(file, fatOffset, 0xFFFFFFFE);
            WriteUInt32(file, fatOffset + 4, 0xFFFFFFFE);
            WriteUInt32(file, fatOffset + 8, 0xFFFFFFFE);
            WriteUInt32(file, fatOffset + 12, 0xFFFFFFFD);

            return file;
        }

        private static byte[] BuildDestList(string path)
        {
            return BuildDestList(new[] { new DestListEntrySpec(path, 42, 99, -1, 7, 3, 132537600000000000UL) });
        }

        private static byte[] BuildDestList(DestListEntrySpec[] entries, uint version = 4, int? declaredEntryCount = null)
        {
            int totalEntryBytes = entries.Sum(entry => 0x82 + Encoding.Unicode.GetByteCount(entry.Path) + 4);
            byte[] data = new byte[32 + totalEntryBytes];
            WriteUInt32(data, 0, version);
            WriteUInt32(data, 4, (uint)(declaredEntryCount ?? entries.Length));
            WriteUInt32(data, 8, 2);
            WriteUInt32(data, 0x0c, BitConverter.ToUInt32(BitConverter.GetBytes(1.5f), 0));
            WriteUInt64(data, 0x10, 0x000000010000002aUL);
            WriteUInt64(data, 0x18, 100);

            int offset = 32;
            foreach (var entry in entries)
            {
                byte[] pathBytes = Encoding.Unicode.GetBytes(entry.Path);
                WriteBytes(data, offset + 0x08, VolumeDroidBytes);
                WriteBytes(data, offset + 0x18, FileDroidBytes);
                WriteBytes(data, offset + 0x28, VolumeBirthDroidBytes);
                WriteBytes(data, offset + 0x38, FileBirthDroidBytes);
                WriteAsciiPadded(data, offset + 0x48, "HOST", 16);
                WriteUInt32(data, offset + 0x58, entry.EntryNumber);
                WriteUInt32(data, offset + 0x5c, entry.EntryNumberUnknown);
                WriteUInt32(data, offset + 0x60, BitConverter.ToUInt32(BitConverter.GetBytes(1.5f), 0));
                WriteUInt64(data, offset + 0x64, entry.LastInteractionFileTime);
                WriteInt32(data, offset + 0x6c, entry.PinStatus);
                WriteInt32(data, offset + 0x70, entry.RecentRank);
                WriteUInt32(data, offset + 0x74, entry.AccessCount);
                WriteUInt32(data, offset + 0x78, 0x11111111);
                WriteUInt32(data, offset + 0x7c, 0x22222222);
                WriteUInt16(data, offset + 0x80, (ushort)entry.Path.Length);
                Array.Copy(pathBytes, 0, data, offset + 0x82, pathBytes.Length);
                WriteUInt32(data, offset + 0x82 + pathBytes.Length, 0);
                offset += 0x82 + pathBytes.Length + 4;
            }

            return data;
        }

        private static void WriteBytes(byte[] data, int offset, byte[] value)
        {
            Array.Copy(value, 0, data, offset, value.Length);
        }

        private static void WriteAsciiPadded(byte[] data, int offset, string value, int length)
        {
            byte[] valueBytes = Encoding.ASCII.GetBytes(value);
            Array.Copy(valueBytes, 0, data, offset, Math.Min(valueBytes.Length, length));
        }

        private sealed class DestListEntrySpec
        {
            public DestListEntrySpec(
                string path,
                uint entryNumber,
                uint entryNumberUnknown,
                int pinStatus,
                int recentRank,
                uint accessCount,
                ulong lastInteractionFileTime)
            {
                Path = path;
                EntryNumber = entryNumber;
                EntryNumberUnknown = entryNumberUnknown;
                PinStatus = pinStatus;
                RecentRank = recentRank;
                AccessCount = accessCount;
                LastInteractionFileTime = lastInteractionFileTime;
            }

            public string Path { get; }

            public uint EntryNumber { get; }

            public uint EntryNumberUnknown { get; }

            public int PinStatus { get; }

            public int RecentRank { get; }

            public uint AccessCount { get; }

            public ulong LastInteractionFileTime { get; }
        }

        private static void WriteDirectoryEntry(byte[] data, int offset, string name, byte objectType, uint startSector, ulong streamSize)
        {
            byte[] nameBytes = Encoding.Unicode.GetBytes(name + "\0");
            Array.Copy(nameBytes, 0, data, offset, nameBytes.Length);
            WriteUInt16(data, offset + 64, (ushort)nameBytes.Length);
            data[offset + 66] = objectType;
            WriteUInt32(data, offset + 116, startSector);
            WriteUInt64(data, offset + 120, streamSize);
        }

        private static void WriteUInt16(byte[] data, int offset, ushort value)
        {
            Array.Copy(BitConverter.GetBytes(value), 0, data, offset, 2);
        }

        private static void WriteUInt32(byte[] data, int offset, uint value)
        {
            Array.Copy(BitConverter.GetBytes(value), 0, data, offset, 4);
        }

        private static void WriteUInt64(byte[] data, int offset, ulong value)
        {
            Array.Copy(BitConverter.GetBytes(value), 0, data, offset, 8);
        }

        private static void WriteInt32(byte[] data, int offset, int value)
        {
            Array.Copy(BitConverter.GetBytes(value), 0, data, offset, 4);
        }
    }
}
