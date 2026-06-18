using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class QuickAccessRestoreTests
    {
        [TestMethod]
        public void RestoreRecentFilesDefaults_DefaultCleanupDeletesOnlyFileShortcuts()
        {
            var links = new StubRecentLinkFileSystem(new[]
            {
                @"C:\Recent\file.lnk",
                @"C:\Recent\dir.lnk",
                @"C:\Recent\unknown.lnk",
                @"C:\Recent\unresolved.lnk",
                @"C:\Recent\note.txt"
            });
            var resolver = new StubShortcutResolver(new Dictionary<string, ShortcutResolution>
            {
                [@"C:\Recent\file.lnk"] = new ShortcutResolution(@"C:\Data\a.txt", false),
                [@"C:\Recent\dir.lnk"] = new ShortcutResolution(@"C:\Data", true),
                [@"C:\Recent\unknown.lnk"] = new ShortcutResolution(@"C:\Missing", null),
                [@"C:\Recent\unresolved.lnk"] = null
            });
            var engine = CreateEngine(links, resolver);
            TimeSpan? clearTimeout = null;

            var report = engine.RestoreRecentFilesDefaults(
                new RestoreDefaultsOptions(),
                timeout => clearTimeout = timeout);

            Assert.IsTrue(report.Success);
            Assert.IsTrue(report.RecentFilesCleared);
            Assert.IsNull(report.Error);
            CollectionAssert.AreEqual(new[] { @"C:\Recent\file.lnk" }, links.DeletedFiles);
            CollectionAssert.AreEqual(new[] { @"C:\Recent\file.lnk" }, report.DeletedLnkPaths.ToList());
            Assert.AreEqual(TimeSpan.FromSeconds(10), clearTimeout.Value);
        }

        [TestMethod]
        public void RestoreRecentFilesDefaults_DeepCleanupDeletesUnknownAndUnresolvedShortcuts()
        {
            var links = new StubRecentLinkFileSystem(new[]
            {
                @"C:\Recent\file.lnk",
                @"C:\Recent\dir.lnk",
                @"C:\Recent\unknown.lnk",
                @"C:\Recent\unresolved.lnk"
            });
            var resolver = new StubShortcutResolver(new Dictionary<string, ShortcutResolution>
            {
                [@"C:\Recent\file.lnk"] = new ShortcutResolution(@"C:\Data\a.txt", false),
                [@"C:\Recent\dir.lnk"] = new ShortcutResolution(@"C:\Data", true),
                [@"C:\Recent\unknown.lnk"] = new ShortcutResolution(@"C:\Missing", null),
                [@"C:\Recent\unresolved.lnk"] = null
            });
            var engine = CreateEngine(links, resolver);

            var report = engine.RestoreRecentFilesDefaults(
                new RestoreDefaultsOptions { DeepLnkCleanup = true },
                _ => { });

            Assert.IsTrue(report.Success);
            CollectionAssert.AreEqual(
                new[]
                {
                    @"C:\Recent\file.lnk",
                    @"C:\Recent\unknown.lnk",
                    @"C:\Recent\unresolved.lnk"
                },
                links.DeletedFiles);
        }

        [TestMethod]
        public void RestoreFrequentFoldersDefaults_DeepCleanupDeletesDirectoryUnknownAndUnresolvedShortcuts()
        {
            var links = new StubRecentLinkFileSystem(new[]
            {
                @"C:\Recent\file.lnk",
                @"C:\Recent\dir.lnk",
                @"C:\Recent\unknown.lnk",
                @"C:\Recent\unresolved.lnk"
            });
            var resolver = new StubShortcutResolver(new Dictionary<string, ShortcutResolution>
            {
                [@"C:\Recent\file.lnk"] = new ShortcutResolution(@"C:\Data\a.txt", false),
                [@"C:\Recent\dir.lnk"] = new ShortcutResolution(@"C:\Data", true),
                [@"C:\Recent\unknown.lnk"] = new ShortcutResolution(@"C:\Missing", null),
                [@"C:\Recent\unresolved.lnk"] = null
            });
            var fileSystem = new StubFileSystem();
            var reader = new StubDestListReader(CreateDestinations("knownfolder:{desktop}"));
            var engine = CreateEngine(links, resolver, fileSystem, reader);
            var refreshCalls = new List<TimeSpan>();

            var report = engine.RestoreFrequentFoldersDefaults(
                new RestoreDefaultsOptions { DeepLnkCleanup = true, RebuildDelay = TimeSpan.Zero },
                refreshCalls.Add);

            Assert.IsTrue(report.Success);
            Assert.IsTrue(report.BackingFileDeleted);
            Assert.IsTrue(report.Rebuilt);
            Assert.IsNull(report.Error);
            CollectionAssert.AreEqual(
                new[]
                {
                    @"C:\Recent\dir.lnk",
                    @"C:\Recent\unknown.lnk",
                    @"C:\Recent\unresolved.lnk"
                },
                links.DeletedFiles);
            CollectionAssert.AreEqual(new[] { @"C:\frequent.automaticDestinations-ms" }, fileSystem.DeletedFiles);
            CollectionAssert.AreEqual(new[] { TimeSpan.FromSeconds(10) }, refreshCalls);
        }

        [TestMethod]
        public void RestoreFrequentFoldersDefaults_RebuildTimeoutCapturedInReport()
        {
            var fileSystem = new StubFileSystem { FileExistsValue = false };
            var engine = CreateEngine(fileSystem: fileSystem);

            var report = engine.RestoreFrequentFoldersDefaults(
                new RestoreDefaultsOptions
                {
                    RebuildDelay = TimeSpan.Zero,
                    RebuildPollTimeout = TimeSpan.FromMilliseconds(1)
                },
                _ => { });

            Assert.IsFalse(report.Success);
            Assert.IsTrue(report.BackingFileDeleted);
            Assert.IsFalse(report.Rebuilt);
            Assert.IsInstanceOfType(report.Error, typeof(TimeoutException));
        }

        [TestMethod]
        public void RestoreFrequentFoldersDefaults_NonDefaultRawPathRunsSingleCleanupCycle()
        {
            var fileSystem = new StubFileSystem();
            var reader = new StubDestListReader(
                CreateDestinations(@"C:\Projects", "knownfolder:{desktop}"),
                CreateDestinations("knownfolder:{desktop}"));
            var engine = CreateEngine(fileSystem: fileSystem, reader: reader);
            var refreshCalls = new List<TimeSpan>();

            var report = engine.RestoreFrequentFoldersDefaults(
                new RestoreDefaultsOptions { RebuildDelay = TimeSpan.Zero },
                refreshCalls.Add);

            Assert.IsTrue(report.Success);
            Assert.IsNotNull(report.RawPathRemoveReport);
            CollectionAssert.AreEqual(new[] { @"C:\Projects" }, report.RawPathRemoveReport.RequestedRawPaths.ToList());
            CollectionAssert.AreEqual(Array.Empty<string>(), report.RawPathRemoveReport.RemainingNonDefaultRawPaths.ToList());
            CollectionAssert.AreEqual(Array.Empty<string>(), report.NonDefaultRawPaths.ToList());
            Assert.AreEqual(2, fileSystem.DeletedFiles.Count);
            Assert.AreEqual(2, refreshCalls.Count);
        }

        [TestMethod]
        public void RestoreDefaults_AllUsesRestoreEngineAndClearsCache()
        {
            var executor = new Mock<IScriptExecutor>(MockBehavior.Strict);
            var engine = new StubRestoreEngine
            {
                Report = new RestoreDefaultsReport(
                    new RecentRestoreReport(new[] { @"C:\Recent\a.lnk" }, true, null),
                    new FrequentRestoreReport(Array.Empty<string>(), true, true, Array.Empty<string>(), null, null))
            };
            executor.Setup(e => e.ClearCache());
            executor.Setup(e => e.Dispose());
            using (var manager = CreateManager(executor.Object, engine))
            {
                var options = new RestoreDefaultsOptions { RebuildDelay = TimeSpan.Zero };

                var report = manager.RestoreDefaults(QuickAccess.All, options);

                Assert.AreSame(engine.Report, report);
                Assert.AreEqual(QuickAccess.All, engine.Target.Value);
                Assert.AreSame(options, engine.Options);
            }

            executor.Verify(e => e.ClearCache(), Times.Once);
        }

        private static QuickAccessRestoreEngine CreateEngine(
            StubRecentLinkFileSystem links = null,
            StubShortcutResolver resolver = null,
            StubFileSystem fileSystem = null,
            StubDestListReader reader = null)
        {
            return new QuickAccessRestoreEngine(
                new StubDataFiles(),
                new StubRecentFolder(@"C:\Recent"),
                resolver ?? new StubShortcutResolver(new Dictionary<string, ShortcutResolution>()),
                links ?? new StubRecentLinkFileSystem(Array.Empty<string>()),
                fileSystem ?? new StubFileSystem(),
                reader ?? new StubDestListReader(CreateDestinations("knownfolder:{desktop}")),
                new StubDelay());
        }

        private static QuickAccessManager CreateManager(IScriptExecutor executor, IQuickAccessRestoreEngine restoreEngine)
        {
            return new QuickAccessManager(
                executor,
                TimeSpan.FromSeconds(10),
                new StubFileSystem(),
                Mock.Of<INativeMethods>(),
                new StubDataFiles(),
                RetryPolicy.Standard,
                new PowerShellFallbackNativeQuery(),
                Mock.Of<IQuickAccessNativeMutation>(),
                Mock.Of<IExplorerRefresher>(),
                new NoOpRecentLinksCleaner(),
                new NoOpQuickAccessLockFactory(),
                new NoOpQuickAccessVisibility(),
                new NoOpDestListMetadataReader(),
                restoreEngine);
        }

        private static AutomaticDestinations CreateDestinations(params string[] rawPaths)
        {
            return new AutomaticDestinations(
                new CfbInfo(512, 64, 4096, Array.Empty<CfbDirectoryEntry>()),
                new DestList
                {
                    Entries = rawPaths
                        .Select(rawPath => new DestListEntry { RawPath = rawPath, Path = rawPath })
                        .ToList()
                });
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

        private sealed class StubShortcutResolver : IShortcutResolutionResolver
        {
            private readonly Dictionary<string, ShortcutResolution> _resolutions;

            public StubShortcutResolver(Dictionary<string, ShortcutResolution> resolutions)
            {
                _resolutions = resolutions;
            }

            public ShortcutResolution Resolve(string shortcutPath, TimeSpan timeout)
            {
                ShortcutResolution resolution;
                return _resolutions.TryGetValue(shortcutPath, out resolution) ? resolution : null;
            }

            public string ResolveTarget(string shortcutPath, TimeSpan timeout)
            {
                return Resolve(shortcutPath, timeout)?.Path;
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

        private sealed class StubFileSystem : IFileSystemOperations
        {
            public bool FileExistsValue { get; set; } = true;

            public List<string> DeletedFiles { get; } = new List<string>();

            public bool FileExists(string path) => FileExistsValue;

            public bool DirectoryExists(string path) => true;

            public void DeleteFile(string path)
            {
                DeletedFiles.Add(path);
            }
        }

        private sealed class StubDestListReader : IDestListMetadataReader
        {
            public StubDestListReader(params AutomaticDestinations[] responses)
            {
                foreach (var response in responses)
                    Responses.Enqueue(response);
            }

            public Queue<AutomaticDestinations> Responses { get; } = new Queue<AutomaticDestinations>();

            public AutomaticDestinations ParseFile(string path)
            {
                if (Responses.Count == 0)
                    throw new IOException("No DestList response was queued.");

                return Responses.Dequeue();
            }

            public AutomaticDestinations ParseBytes(byte[] data)
            {
                throw new NotSupportedException();
            }
        }

        private sealed class StubDelay : IQuickAccessRestoreDelay
        {
            public void Sleep(TimeSpan delay)
            {
            }
        }

        private sealed class StubRestoreEngine : IQuickAccessRestoreEngine
        {
            public QuickAccess? Target { get; private set; }

            public RestoreDefaultsOptions Options { get; private set; }

            public RestoreDefaultsReport Report { get; set; }

            public RestoreDefaultsReport RestoreDefaults(
                QuickAccess target,
                RestoreDefaultsOptions options,
                Action<TimeSpan> clearRecentFiles,
                Action<TimeSpan> refreshExplorer)
            {
                Target = target;
                Options = options;
                return Report;
            }

            public RecentRestoreReport RestoreRecentFilesDefaults(
                RestoreDefaultsOptions options,
                Action<TimeSpan> clearRecentFiles)
            {
                throw new NotSupportedException();
            }

            public FrequentRestoreReport RestoreFrequentFoldersDefaults(
                RestoreDefaultsOptions options,
                Action<TimeSpan> refreshExplorer)
            {
                throw new NotSupportedException();
            }
        }
    }
}
