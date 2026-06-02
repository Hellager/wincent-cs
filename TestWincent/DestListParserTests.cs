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
            Assert.AreEqual(0x0000000100000000UL, parsed.DestList.LastEntryId);
            Assert.AreEqual(42u, entry.EntryNumber);
            Assert.AreEqual(99u, entry.EntryNumberReserved);
            Assert.AreEqual("2a", entry.StreamName);
            Assert.AreEqual(@"C:\Test\file.txt", entry.RawPath);
            Assert.AreEqual(@"C:\Test\file.txt", entry.Path);
            Assert.AreEqual(-1, entry.PinStatus);
            Assert.IsFalse(entry.PinOrder.HasValue);
            Assert.IsFalse(entry.IsPinned);
            Assert.AreEqual(7, entry.RecentRank);
            Assert.AreEqual(3u, entry.AccessCount);
            Assert.AreEqual(1.5f, entry.Score);
            Assert.AreEqual(0u, entry.SerializedPropertyStoreSize);
            Assert.IsTrue(entry.LastInteractionTime.HasValue);
            Assert.IsTrue(parsed.CfbInfo.DirectoryEntries.Any(e => e.Name == "DestList" && e.ObjectType == CfbObjectType.Stream));
        }

        [TestMethod]
        [Ignore("Integration test; reads the current user's real Recent Files automaticDestinations file.")]
        public void Integration_ParseRecentFilesMetadata()
        {
            var parsed = DestListMetadataParser.ParseFile(DestListMetadataParser.RecentFilesDestPath());

            Assert.IsNotNull(parsed.DestList);
        }

        [TestMethod]
        [Ignore("Integration test; reads the current user's real Frequent Folders automaticDestinations file.")]
        public void Integration_ParseFrequentFoldersMetadata()
        {
            var parsed = DestListMetadataParser.ParseFile(DestListMetadataParser.FrequentFoldersDestPath());

            Assert.IsNotNull(parsed.DestList);
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
        public void ParseBytes_RejectsTruncatedDestListStream_ThrowsDestListParseException()
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

            var exception = Assert.ThrowsException<DestListParseException>(
                () => DestListMetadataParser.ParseBytes(file));

            StringAssert.Contains(exception.Details, "too short");
        }

        private static byte[] BuildMinimalCfbWithDestList(string path)
        {
            byte[] destList = BuildDestList(path);
            byte[] file = new byte[512 + 512 * 4];

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
            WriteUInt32(file, 0x4c, 3);

            int directoryOffset = 512;
            WriteDirectoryEntry(file, directoryOffset, "Root Entry", 5, 1, 0);
            WriteDirectoryEntry(file, directoryOffset + 128, "DestList", 2, 2, (ulong)destList.Length);

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
            byte[] pathBytes = Encoding.Unicode.GetBytes(path);
            int entryLength = 0x82 + pathBytes.Length + 4;
            byte[] data = new byte[32 + entryLength];
            WriteUInt32(data, 0, 4);
            WriteUInt32(data, 4, 1);
            WriteUInt64(data, 8, 0x0000000100000000UL);
            WriteUInt32(data, 0x10, 42);
            WriteUInt32(data, 0x14, 0);
            WriteUInt32(data, 0x18, 100);
            WriteUInt32(data, 0x1c, 0);

            int offset = 32;
            WriteUInt32(data, offset + 0x58, 42);
            WriteUInt32(data, offset + 0x5c, 99);
            WriteUInt32(data, offset + 0x60, BitConverter.ToUInt32(BitConverter.GetBytes(1.5f), 0));
            WriteUInt64(data, offset + 0x64, 132537600000000000UL);
            WriteInt32(data, offset + 0x6c, -1);
            WriteInt32(data, offset + 0x70, 7);
            WriteUInt32(data, offset + 0x74, 3);
            WriteUInt16(data, offset + 0x80, (ushort)path.Length);
            Array.Copy(pathBytes, 0, data, offset + 0x82, pathBytes.Length);
            WriteUInt32(data, offset + 0x82 + pathBytes.Length, 0);
            return data;
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
