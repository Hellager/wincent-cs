using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
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
        public void DeleteForTarget_PassesDecreasingRemainingTime()
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
                },
                ConsumedTime = TimeSpan.FromMilliseconds(5),
            };
            var cleaner = new RecentLinksCleaner(
                new FakeRecentFolder(@"C:\Recent"),
                resolver,
                fileSystem);

            cleaner.DeleteForTarget(@"C:\Work\Report.docx", TimeSpan.FromSeconds(10));

            Assert.AreEqual(3, resolver.ReceivedTimeouts.Count);
            for (int i = 1; i < resolver.ReceivedTimeouts.Count; i++)
            {
                Assert.IsTrue(
                    resolver.ReceivedTimeouts[i] < resolver.ReceivedTimeouts[i - 1],
                    $"Timeout at index {i} ({resolver.ReceivedTimeouts[i]}) should be less than at index {i - 1} ({resolver.ReceivedTimeouts[i - 1]})");
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
        public void DeleteForTarget_ExhaustedRemainingTimeThrowsTimeoutException()
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

            var ex = Assert.ThrowsException<TimeoutException>(
                () => cleaner.DeleteForTarget(@"C:\Work\Report.docx", TimeSpan.FromMilliseconds(10)));

            Assert.IsTrue(ex.Message.Contains("timed out"), "Exception message should indicate timeout.");
            Assert.AreEqual(1, resolver.CallCount, "Only the first link should be resolved before time is exhausted.");
        }

        [TestMethod]
        public void IsShortcutFile_IsCaseInsensitive()
        {
            Assert.IsTrue(RecentLinksCleaner.IsShortcutFile(@"C:\Recent\a.lnk"));
            Assert.IsTrue(RecentLinksCleaner.IsShortcutFile(@"C:\Recent\a.LNK"));
            Assert.IsFalse(RecentLinksCleaner.IsShortcutFile(@"C:\Recent\a.txt"));
            Assert.IsFalse(RecentLinksCleaner.IsShortcutFile(@"C:\Recent\a"));
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
    }
}
