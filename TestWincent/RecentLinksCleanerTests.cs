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

            public string ResolveTarget(string shortcutPath, TimeSpan timeout)
            {
                return Targets.TryGetValue(shortcutPath, out var target) ? target : null;
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
