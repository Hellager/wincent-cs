using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class RecentLinksCleanerTests
    {
        [TestMethod]
        public void DeleteForTarget_DeletesOnlyMatchingShortcutFiles()
        {
            string directory = CreateTempDirectory();
            try
            {
                string matching = Path.Combine(directory, "matching.lnk");
                string other = Path.Combine(directory, "other.LNK");
                string broken = Path.Combine(directory, "broken.lnk");
                string nonShortcut = Path.Combine(directory, "matching.txt");
                File.WriteAllText(matching, "matching");
                File.WriteAllText(other, "other");
                File.WriteAllText(broken, "broken");
                File.WriteAllText(nonShortcut, "txt");

                var resolver = new FakeShortcutTargetResolver
                {
                    Targets =
                    {
                        [matching] = @"C:\Work\Report.docx",
                        [other] = @"C:\Work\Other.docx",
                    }
                };
                var cleaner = new RecentLinksCleaner(
                    new FakeRecentFolder(directory),
                    resolver,
                    new DefaultRecentLinkFileSystem());

                var deleted = cleaner.DeleteForTarget("c:/work/report.docx", TimeSpan.FromSeconds(1));

                CollectionAssert.AreEqual(new[] { matching }, new List<string>(deleted));
                Assert.IsFalse(File.Exists(matching));
                Assert.IsTrue(File.Exists(other));
                Assert.IsTrue(File.Exists(broken));
                Assert.IsTrue(File.Exists(nonShortcut));
            }
            finally
            {
                Directory.Delete(directory, recursive: true);
            }
        }

        [TestMethod]
        public void DeleteForTarget_DeleteFailurePropagates()
        {
            var fileSystem = new FakeRecentLinkFileSystem
            {
                Files = { @"C:\Recent\matching.lnk" },
                DeleteException = new IOException("delete failed"),
            };
            var resolver = new FakeShortcutTargetResolver
            {
                Targets =
                {
                    [@"C:\Recent\matching.lnk"] = @"C:\Work\Report.docx",
                }
            };
            var cleaner = new RecentLinksCleaner(
                new FakeRecentFolder(@"C:\Recent"),
                resolver,
                fileSystem);

            Assert.ThrowsException<IOException>(
                () => cleaner.DeleteForTarget(@"C:\Work\Report.docx", TimeSpan.FromSeconds(1)));
        }

        [TestMethod]
        public void DeleteForTarget_PassesFullTimeoutToEachResolverCall()
        {
            var fileSystem = new FakeRecentLinkFileSystem
            {
                Files = { @"C:\Recent\a.lnk", @"C:\Recent\b.lnk", @"C:\Recent\c.lnk" },
            };
            var resolver = new FakeShortcutTargetResolver
            {
                Targets =
                {
                    [@"C:\Recent\a.lnk"] = @"C:\Work\Report.docx",
                    [@"C:\Recent\b.lnk"] = @"C:\Work\Report.docx",
                    [@"C:\Recent\c.lnk"] = @"C:\Work\Report.docx",
                }
            };
            var cleaner = new RecentLinksCleaner(
                new FakeRecentFolder(@"C:\Recent"),
                resolver,
                fileSystem);

            cleaner.DeleteForTarget(@"C:\Work\Report.docx", TimeSpan.FromSeconds(10));

            Assert.AreEqual(3, resolver.ReceivedTimeouts.Count);
            foreach (var receivedTimeout in resolver.ReceivedTimeouts)
            {
                Assert.AreEqual(TimeSpan.FromSeconds(10), receivedTimeout);
            }
        }

        [TestMethod]
        public void DeleteForTarget_TimeoutExceptionPropagates()
        {
            var fileSystem = new FakeRecentLinkFileSystem
            {
                Files = { @"C:\Recent\a.lnk", @"C:\Recent\b.lnk" },
            };
            var resolver = new FakeShortcutTargetResolver
            {
                Targets =
                {
                    [@"C:\Recent\a.lnk"] = @"C:\Work\Report.docx",
                    [@"C:\Recent\b.lnk"] = @"C:\Work\Report.docx",
                },
                ThrowTimeoutOnPath = @"C:\Recent\a.lnk",
            };
            var cleaner = new RecentLinksCleaner(
                new FakeRecentFolder(@"C:\Recent"),
                resolver,
                fileSystem);

            Assert.ThrowsException<TimeoutException>(
                () => cleaner.DeleteForTarget(@"C:\Work\Report.docx", TimeSpan.FromSeconds(1)));

            Assert.AreEqual(1, resolver.CallCount, "Only the first link should have been resolved before timeout propagated.");
        }

        [TestMethod]
        public void DeleteForTarget_DoesNotUseOverallStopwatchTimeout()
        {
            var fileSystem = new FakeRecentLinkFileSystem
            {
                Files = { @"C:\Recent\a.lnk", @"C:\Recent\b.lnk" },
            };
            var resolver = new FakeShortcutTargetResolver
            {
                Targets =
                {
                    [@"C:\Recent\a.lnk"] = @"C:\Work\Report.docx",
                    [@"C:\Recent\b.lnk"] = @"C:\Work\Report.docx",
                },
                ConsumedTime = TimeSpan.FromMilliseconds(200),
            };
            var cleaner = new RecentLinksCleaner(
                new FakeRecentFolder(@"C:\Recent"),
                resolver,
                fileSystem);

            var deleted = cleaner.DeleteForTarget(@"C:\Work\Report.docx", TimeSpan.FromMilliseconds(10));

            CollectionAssert.AreEqual(
                new[] { @"C:\Recent\a.lnk", @"C:\Recent\b.lnk" },
                new List<string>(deleted));
            Assert.AreEqual(2, resolver.CallCount, "Each link should get its own resolver timeout budget.");
            CollectionAssert.AreEqual(
                new[] { TimeSpan.FromMilliseconds(10), TimeSpan.FromMilliseconds(10) },
                resolver.ReceivedTimeouts);
        }

        [TestMethod]
        public void IsShortcutFile_IsCaseInsensitive()
        {
            Assert.IsTrue(RecentLinksCleaner.IsShortcutFile(@"C:\Recent\a.lnk"));
            Assert.IsTrue(RecentLinksCleaner.IsShortcutFile(@"C:\Recent\a.LNK"));
            Assert.IsFalse(RecentLinksCleaner.IsShortcutFile(@"C:\Recent\a.txt"));
            Assert.IsFalse(RecentLinksCleaner.IsShortcutFile(@"C:\Recent\a"));
        }

        [TestMethod]
        public void ParseShellLinkSummary_InvalidHeader_ReturnsNull()
        {
            Assert.IsNull(ShellLinkParser.ParseShellLinkSummary(Encoding.UTF8.GetBytes("not a link")));
        }

        [TestMethod]
        public void ResolveBytes_RelativePathOnly_UsesRelativePathAndHeaderAttributes()
        {
            var lnk = MinimalLnk();
            WriteUInt32(lnk, 0x18, FileAttributeDirectory);

            var resolution = ShellLinkParser.ResolveBytes(lnk.ToArray(), TimeSpan.Zero);

            Assert.IsNotNull(resolution);
            Assert.AreEqual("relative-target", resolution.Path);
            Assert.AreEqual(true, resolution.IsDirectory);
        }

        [TestMethod]
        public void ParseShellLinkSummary_ReadsUnicodeLocalBaseAndSuffix()
        {
            var lnk = LnkWithLinkInfo(
                localBase: @"C:\Users\Alice",
                commonSuffix: @"Documents\report.docx",
                networkBase: null,
                attributes: FileAttributeArchive);

            var summary = ShellLinkParser.ParseShellLinkSummary(lnk.ToArray());

            Assert.IsNotNull(summary);
            Assert.AreEqual(@"C:\Users\Alice\Documents\report.docx", summary.TargetPath);
            Assert.AreEqual(FileAttributeArchive, summary.FileAttributes);
            Assert.IsFalse(summary.TargetIsNetwork);
        }

        [TestMethod]
        public void ParseShellLinkSummary_ReadsUncNetworkPath()
        {
            var lnk = LnkWithLinkInfo(
                localBase: null,
                commonSuffix: @"Share\report.docx",
                networkBase: @"\\server\team",
                attributes: FileAttributeArchive);

            var summary = ShellLinkParser.ParseShellLinkSummary(lnk.ToArray());

            Assert.IsNotNull(summary);
            Assert.AreEqual(@"\\server\team\Share\report.docx", summary.TargetPath);
            Assert.IsTrue(summary.TargetIsNetwork);
        }

        [TestMethod]
        public void DetermineTargetIsDirectory_UsesAttributesWithoutMetadataProbe()
        {
            Assert.AreEqual(
                true,
                ShellLinkTargetResolver.DetermineTargetIsDirectory(
                    @"\\server\share\folder",
                    FileAttributeDirectory,
                    targetIsNetwork: true,
                    timeout: TimeSpan.Zero));
            Assert.AreEqual(
                false,
                ShellLinkTargetResolver.DetermineTargetIsDirectory(
                    @"\\server\share\file.txt",
                    FileAttributeArchive,
                    targetIsNetwork: true,
                    timeout: TimeSpan.Zero));
        }

        [TestMethod]
        public void DetermineTargetIsDirectory_LocalTarget_UsesFileSystemMetadata()
        {
            string directory = CreateTempDirectory();
            try
            {
                string folder = Path.Combine(directory, "folder");
                Directory.CreateDirectory(folder);

                Assert.AreEqual(
                    true,
                    ShellLinkTargetResolver.DetermineTargetIsDirectory(
                        folder,
                        attributes: 0,
                        targetIsNetwork: false,
                        timeout: TimeSpan.Zero));
            }
            finally
            {
                Directory.Delete(directory, recursive: true);
            }
        }

        private static string CreateTempDirectory()
        {
            string path = Path.Combine(Path.GetTempPath(), "WincentRecentLinksCleanerTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(path);
            return path;
        }

        private sealed class FakeRecentFolder : IWindowsRecentFolder
        {
            private readonly string _path;

            public FakeRecentFolder(string path)
            {
                _path = path;
            }

            public string GetPath()
            {
                return _path;
            }
        }

        private sealed class FakeShortcutTargetResolver : IShortcutTargetResolver
        {
            public Dictionary<string, string> Targets { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            public List<TimeSpan> ReceivedTimeouts { get; } = new List<TimeSpan>();

            public string ThrowTimeoutOnPath { get; set; }

            public TimeSpan? ConsumedTime { get; set; }

            public int CallCount { get; private set; }

            public string ResolveTarget(string shortcutPath, TimeSpan timeout)
            {
                CallCount++;
                ReceivedTimeouts.Add(timeout);

                if (ThrowTimeoutOnPath != null &&
                    string.Equals(shortcutPath, ThrowTimeoutOnPath, StringComparison.OrdinalIgnoreCase))
                {
                    throw new TimeoutException("Simulated timeout.");
                }

                if (ConsumedTime.HasValue)
                {
                    var target = DateTime.UtcNow + ConsumedTime.Value;
                    while (DateTime.UtcNow < target)
                    {
                    }
                }

                return Targets.TryGetValue(shortcutPath, out var targetPath) ? targetPath : null;
            }

            public ShortcutResolution Resolve(string shortcutPath, TimeSpan timeout)
            {
                string target = ResolveTarget(shortcutPath, timeout);
                return target == null ? null : new ShortcutResolution(target, null);
            }
        }

        private sealed class FakeRecentLinkFileSystem : IRecentLinkFileSystem
        {
            public List<string> Files { get; } = new List<string>();

            public Exception DeleteException { get; set; }

            public IEnumerable<string> EnumerateFiles(string directory)
            {
                return Files;
            }

            public void DeleteFile(string path)
            {
                if (DeleteException != null)
                    throw DeleteException;
            }
        }

        private const uint FileAttributeDirectory = 0x00000010;
        private const uint FileAttributeArchive = 0x00000020;
        private const uint ShellLinkHeaderSize = 0x4c;
        private const uint HasLinkInfo = 0x00000002;
        private const uint HasRelativePath = 0x00000008;
        private const uint IsUnicode = 0x00000080;
        private const uint VolumeIdAndLocalBasePath = 0x00000001;
        private const uint CommonNetworkRelativeLinkAndPathSuffix = 0x00000002;

        private static readonly byte[] LinkClsid =
        {
            0x01, 0x14, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
            0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46
        };

        private static List<byte> MinimalLnk()
        {
            var lnk = new List<byte>(new byte[ShellLinkHeaderSize]);
            WriteUInt32(lnk, 0, ShellLinkHeaderSize);
            for (int i = 0; i < LinkClsid.Length; i++)
                lnk[4 + i] = LinkClsid[i];
            WriteUInt32(lnk, 0x14, HasRelativePath | IsUnicode);
            WriteLnkString(lnk, "relative-target");
            return lnk;
        }

        private static List<byte> LnkWithLinkInfo(string localBase, string commonSuffix, string networkBase, uint attributes)
        {
            var lnk = new List<byte>(new byte[ShellLinkHeaderSize]);
            WriteUInt32(lnk, 0, ShellLinkHeaderSize);
            for (int i = 0; i < LinkClsid.Length; i++)
                lnk[4 + i] = LinkClsid[i];
            WriteUInt32(lnk, 0x14, HasLinkInfo | IsUnicode);
            WriteUInt32(lnk, 0x18, attributes);

            int linkInfoStart = lnk.Count;
            AddZeros(lnk, 0x24);
            uint flags = 0;
            uint localBaseOffset = 0;
            uint networkOffset = 0;
            uint commonSuffixOffset = 0;

            if (localBase != null)
            {
                flags |= VolumeIdAndLocalBasePath;
                localBaseOffset = (uint)(lnk.Count - linkInfoStart);
                WriteUtf16Z(lnk, localBase);
            }

            if (networkBase != null)
            {
                flags |= CommonNetworkRelativeLinkAndPathSuffix;
                networkOffset = (uint)(lnk.Count - linkInfoStart);
                WriteNetworkLink(lnk, networkBase);
            }

            if (commonSuffix != null)
            {
                commonSuffixOffset = (uint)(lnk.Count - linkInfoStart);
                WriteUtf16Z(lnk, commonSuffix);
            }

            uint size = (uint)(lnk.Count - linkInfoStart);
            WriteUInt32(lnk, linkInfoStart, size);
            WriteUInt32(lnk, linkInfoStart + 4, 0x24);
            WriteUInt32(lnk, linkInfoStart + 8, flags);
            WriteUInt32(lnk, linkInfoStart + 16, localBaseOffset);
            WriteUInt32(lnk, linkInfoStart + 20, networkOffset);
            WriteUInt32(lnk, linkInfoStart + 24, commonSuffixOffset);
            WriteUInt32(lnk, linkInfoStart + 28, localBaseOffset);
            WriteUInt32(lnk, linkInfoStart + 32, commonSuffixOffset);
            return lnk;
        }

        private static void WriteNetworkLink(List<byte> data, string netName)
        {
            int start = data.Count;
            AddZeros(data, 0x18);
            uint netNameOffset = 0x18;
            WriteUtf16Z(data, netName);
            uint size = (uint)(data.Count - start);
            WriteUInt32(data, start, size);
            WriteUInt32(data, start + 8, netNameOffset);
            WriteUInt32(data, start + 20, netNameOffset);
        }

        private static void WriteLnkString(List<byte> data, string value)
        {
            var words = Encoding.Unicode.GetBytes(value);
            ushort chars = (ushort)(words.Length / 2);
            data.AddRange(BitConverter.GetBytes(chars));
            data.AddRange(words);
        }

        private static void WriteUtf16Z(List<byte> data, string value)
        {
            data.AddRange(Encoding.Unicode.GetBytes(value));
            data.AddRange(BitConverter.GetBytes((ushort)0));
        }

        private static void WriteUInt32(List<byte> data, int offset, uint value)
        {
            byte[] bytes = BitConverter.GetBytes(value);
            for (int i = 0; i < bytes.Length; i++)
                data[offset + i] = bytes[i];
        }

        private static void AddZeros(List<byte> data, int count)
        {
            for (int i = 0; i < count; i++)
                data.Add(0);
        }
    }
}
