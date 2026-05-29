using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class ExperimentalDestListRemovalTests
    {
        [TestMethod]
        public void RemoveEntryPathsByRebuild_RejectsEmptyTargets()
        {
            var engine = CreateEngine();

            Assert.ThrowsException<ArgumentException>(
                () => engine.RemoveEntryPathsByRebuild(
                    AutomaticDestinationsKind.RecentFiles,
                    Array.Empty<string>(),
                    new ExperimentalRemoveOptions()));
        }

        [TestMethod]
        public void RemoveEntryPathsByRebuild_DeletesDestinationAndMatchingShortcuts()
        {
            var reader = new StubDestListReader();
            reader.Responses.Enqueue(CreateDestinations(@"C:\Target\a.txt", @"C:\Other\b.txt"));
            reader.Responses.Enqueue(CreateDestinations(@"C:\Other\b.txt"));
            var recentLinks = new StubRecentLinkFileSystem(new[]
            {
                @"C:\Recent\a.lnk",
                @"C:\Recent\b.lnk",
                @"C:\Recent\note.txt"
            });
            var resolver = new StubShortcutTargetResolver(new Dictionary<string, string>
            {
                [@"C:\Recent\a.lnk"] = @"C:\Target\a.txt",
                [@"C:\Recent\b.lnk"] = @"C:\Other\b.txt"
            });
            var fileSystem = new StubExperimentalRemovalFileSystem();
            var engine = CreateEngine(reader, recentLinks, resolver, fileSystem);

            var report = engine.RemoveEntryPathsByRebuild(
                AutomaticDestinationsKind.RecentFiles,
                new[] { @"C:\Target\a.txt" },
                new ExperimentalRemoveOptions());

            Assert.IsTrue(report.Success);
            Assert.IsTrue(report.DestinationDeleted);
            Assert.IsTrue(report.Rebuilt);
            CollectionAssert.AreEqual(new[] { @"C:\Target\a.txt" }, new List<string>(report.MatchingPathsBefore));
            CollectionAssert.AreEqual(new[] { @"C:\Recent\a.lnk" }, new List<string>(report.DeletedShortcutPaths));
            Assert.AreEqual(0, report.MissingShortcutTargetPaths.Count);
            Assert.AreEqual(0, report.RemainingPathsAfterRebuild.Count);
            CollectionAssert.AreEqual(new[] { @"C:\recent.automaticDestinations-ms" }, fileSystem.DeletedFiles);
        }

        [TestMethod]
        public void RemoveEntryPathsByRebuild_ReportsRemainingPathsWhenRebuildStillContainsTarget()
        {
            var reader = new StubDestListReader();
            reader.Responses.Enqueue(CreateDestinations(@"C:\Target\a.txt"));
            reader.Responses.Enqueue(CreateDestinations(@"C:\Target\a.txt"));
            var engine = CreateEngine(reader);

            var report = engine.RemoveEntryPathsByRebuild(
                AutomaticDestinationsKind.RecentFiles,
                new[] { @"C:\Target\a.txt" },
                new ExperimentalRemoveOptions());

            Assert.IsFalse(report.Success);
            CollectionAssert.AreEqual(new[] { @"C:\Target\a.txt" }, new List<string>(report.RemainingPathsAfterRebuild));
        }

        [TestMethod]
        public void RemoveEntryPathsByRebuild_MissingDestination_ThrowsFileNotFoundException()
        {
            var fileSystem = new StubExperimentalRemovalFileSystem { FileExistsValue = false };
            var engine = CreateEngine(fileSystem: fileSystem);

            Assert.ThrowsException<System.IO.FileNotFoundException>(
                () => engine.RemoveEntryPathsByRebuild(
                    AutomaticDestinationsKind.RecentFiles,
                    new[] { @"C:\Target\a.txt" },
                    new ExperimentalRemoveOptions()));
        }

        private static ExperimentalDestListRemovalEngine CreateEngine(
            StubDestListReader reader = null,
            StubRecentLinkFileSystem recentLinks = null,
            StubShortcutTargetResolver resolver = null,
            StubExperimentalRemovalFileSystem fileSystem = null)
        {
            return new ExperimentalDestListRemovalEngine(
                new StubDataFiles(),
                new StubRecentFolder(@"C:\Recent"),
                resolver ?? new StubShortcutTargetResolver(new Dictionary<string, string>()),
                recentLinks ?? new StubRecentLinkFileSystem(Array.Empty<string>()),
                new StubExplorerRefresher(),
                reader ?? new StubDestListReader(CreateDestinations()),
                new StubDelay(),
                fileSystem ?? new StubExperimentalRemovalFileSystem());
        }

        private static AutomaticDestinations CreateDestinations(params string[] paths)
        {
            var destList = new DestList
            {
                Entries = Array.ConvertAll(paths, path => new DestListEntry { Path = path })
            };
            return new AutomaticDestinations(
                new CfbInfo(512, 64, 4096, Array.Empty<CfbDirectoryEntry>()),
                destList);
        }

        private sealed class StubDataFiles : IQuickAccessDataFiles
        {
            public DateTime GetModifiedTimeForScript(PSScript scriptType) => DateTime.Now;

            public void RemoveRecentFile()
            {
            }

            public string RecentFilesPath => @"C:\recent.automaticDestinations-ms";

            public string FrequentFoldersPath => @"C:\frequent.automaticDestinations-ms";
        }

        private sealed class StubRecentFolder : IWindowsRecentFolder
        {
            private readonly string _path;

            public StubRecentFolder(string path)
            {
                _path = path;
            }

            public string GetPath() => _path;
        }

        private sealed class StubShortcutTargetResolver : IShortcutTargetResolver
        {
            private readonly Dictionary<string, string> _targets;

            public StubShortcutTargetResolver(Dictionary<string, string> targets)
            {
                _targets = targets;
            }

            public string ResolveTarget(string shortcutPath, TimeSpan timeout)
            {
                string target;
                return _targets.TryGetValue(shortcutPath, out target) ? target : null;
            }
        }

        private sealed class StubRecentLinkFileSystem : IRecentLinkFileSystem
        {
            private readonly IEnumerable<string> _files;

            public StubRecentLinkFileSystem(IEnumerable<string> files)
            {
                _files = files;
            }

            public List<string> DeletedFiles { get; } = new List<string>();

            public IEnumerable<string> EnumerateFiles(string directory) => _files;

            public void DeleteFile(string path)
            {
                DeletedFiles.Add(path);
            }
        }

        private sealed class StubExplorerRefresher : IExplorerRefresher
        {
            public void Refresh(TimeSpan timeout)
            {
            }
        }

        private sealed class StubDestListReader : IDestListMetadataReader
        {
            public StubDestListReader()
            {
            }

            public StubDestListReader(AutomaticDestinations response)
            {
                Responses.Enqueue(response);
            }

            public Queue<AutomaticDestinations> Responses { get; } = new Queue<AutomaticDestinations>();

            public AutomaticDestinations ParseFile(string path)
            {
                return Responses.Count > 0 ? Responses.Dequeue() : CreateDestinations();
            }

            public AutomaticDestinations ParseBytes(byte[] data)
            {
                throw new NotSupportedException();
            }
        }

        private sealed class StubDelay : IExperimentalRemovalDelay
        {
            public void Sleep(TimeSpan delay)
            {
            }
        }

        private sealed class StubExperimentalRemovalFileSystem : IExperimentalRemovalFileSystem
        {
            public List<string> DeletedFiles { get; } = new List<string>();

            public bool FileExistsValue { get; set; } = true;

            public bool FileExists(string path) => FileExistsValue;

            public void DeleteFile(string path)
            {
                DeletedFiles.Add(path);
            }
        }
    }
}
