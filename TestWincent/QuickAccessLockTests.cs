using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class QuickAccessLockTests
    {
        [TestMethod]
        public void Factory_RecentFiles_OpensRecentBackingFileAndCapturesShortcutSnapshot()
        {
            var dataFiles = new StubDataFiles();
            var recentFolder = new StubRecentFolder(@"C:\Recent");
            var fileSystem = new StubRecentLinkFileSystem(new[]
            {
                @"C:\Recent\a.lnk",
                @"C:\Recent\b.txt"
            });
            var opener = new StubHandleOpener();
            var factory = new QuickAccessLockFactory(dataFiles, recentFolder, fileSystem, opener);

            using (var quickAccessLock = factory.Lock(QuickAccessLockTarget.RecentFiles))
            {
                Assert.AreEqual(QuickAccessLockTarget.RecentFiles, quickAccessLock.Target);
                Assert.AreEqual(@"C:\Recent", quickAccessLock.RecentFolder);
                Assert.AreEqual(1, quickAccessLock.LockedFileCount);
                CollectionAssert.AreEqual(new[] { @"C:\recent.automaticDestinations-ms" }, opener.OpenedPaths);
                CollectionAssert.AreEqual(new[] { @"C:\Recent\a.lnk" }, quickAccessLock.InitialShortcutPaths.ToList());
            }
        }

        [TestMethod]
        public void Factory_FrequentFolders_OpensFrequentBackingFileAndStillCapturesRecentSnapshot()
        {
            var opener = new StubHandleOpener();
            var factory = new QuickAccessLockFactory(
                new StubDataFiles(),
                new StubRecentFolder(@"C:\Recent"),
                new StubRecentLinkFileSystem(new[] { @"C:\Recent\folder.lnk" }),
                opener);

            using (var quickAccessLock = factory.Lock(QuickAccessLockTarget.FrequentFolders))
            {
                Assert.AreEqual(QuickAccessLockTarget.FrequentFolders, quickAccessLock.Target);
                Assert.AreEqual(1, quickAccessLock.LockedFileCount);
                CollectionAssert.AreEqual(new[] { @"C:\frequent.automaticDestinations-ms" }, opener.OpenedPaths);
                CollectionAssert.AreEqual(new[] { @"C:\Recent\folder.lnk" }, quickAccessLock.InitialShortcutPaths.ToList());
            }
        }

        [TestMethod]
        public void Factory_All_OpensBothBackingFiles()
        {
            var opener = new StubHandleOpener();
            var factory = new QuickAccessLockFactory(
                new StubDataFiles(),
                new StubRecentFolder(@"C:\Recent"),
                new StubRecentLinkFileSystem(Array.Empty<string>()),
                opener);

            using (var quickAccessLock = factory.Lock(QuickAccessLockTarget.All))
            {
                Assert.AreEqual(QuickAccessLockTarget.All, quickAccessLock.Target);
                Assert.AreEqual(2, quickAccessLock.LockedFileCount);
                CollectionAssert.AreEqual(
                    new[] { @"C:\recent.automaticDestinations-ms", @"C:\frequent.automaticDestinations-ms" },
                    opener.OpenedPaths);
            }
        }

        [TestMethod]
        public void Factory_All_WhenSecondOpenFails_DisposesFirstHandle()
        {
            var opener = new StubHandleOpener
            {
                ThrowOnPath = @"C:\frequent.automaticDestinations-ms"
            };
            var factory = new QuickAccessLockFactory(
                new StubDataFiles(),
                new StubRecentFolder(@"C:\Recent"),
                new StubRecentLinkFileSystem(Array.Empty<string>()),
                opener);

            Assert.ThrowsException<IOException>(() => factory.Lock(QuickAccessLockTarget.All));

            Assert.AreEqual(1, opener.Handles.Count);
            Assert.IsTrue(opener.Handles[0].IsDisposed);
        }

        [TestMethod]
        public void Unlock_ReturnsSnapshotDiffAndDisposesHandles()
        {
            var handle = new StubHandle();
            var fileSystem = new StubRecentLinkFileSystem(new[] { @"C:\Recent\a.lnk", @"C:\Recent\c.lnk" });
            var quickAccessLock = new QuickAccessLock(
                QuickAccessLockTarget.RecentFiles,
                @"C:\Recent",
                new[] { @"C:\Recent\a.lnk", @"C:\Recent\b.lnk" },
                new[] { handle },
                fileSystem);

            var report = quickAccessLock.Unlock();

            Assert.IsTrue(handle.IsDisposed);
            CollectionAssert.AreEqual(new[] { @"C:\Recent\a.lnk", @"C:\Recent\b.lnk" }, report.InitialShortcutPaths.ToList());
            CollectionAssert.AreEqual(new[] { @"C:\Recent\a.lnk", @"C:\Recent\c.lnk" }, report.CurrentShortcutPaths.ToList());
            CollectionAssert.AreEqual(new[] { @"C:\Recent\c.lnk" }, report.NewShortcutPaths.ToList());
            CollectionAssert.AreEqual(new[] { @"C:\Recent\b.lnk" }, report.DisappearedShortcutPaths.ToList());
            CollectionAssert.AreEqual(Array.Empty<string>(), report.DeletedShortcutPaths.ToList());
        }

        [TestMethod]
        public void Unlock_WithCleanup_DeletesNewShortcutsAndReportsFailures()
        {
            var fileSystem = new StubRecentLinkFileSystem(new[] { @"C:\Recent\new.lnk", @"C:\Recent\failing.lnk" })
            {
                DeleteFailurePath = @"C:\Recent\failing.lnk"
            };
            var quickAccessLock = new QuickAccessLock(
                QuickAccessLockTarget.All,
                @"C:\Recent",
                Array.Empty<string>(),
                new[] { new StubHandle(), new StubHandle() },
                fileSystem);

            var report = quickAccessLock.Unlock(new QuickAccessUnlockOptions { CleanupNewRecentLinks = true });

            CollectionAssert.AreEqual(new[] { @"C:\Recent\new.lnk" }, fileSystem.DeletedPaths);
            CollectionAssert.AreEqual(new[] { @"C:\Recent\new.lnk", @"C:\Recent\failing.lnk" }, report.CurrentShortcutPaths.ToList());
            CollectionAssert.AreEqual(new[] { @"C:\Recent\new.lnk", @"C:\Recent\failing.lnk" }, report.NewShortcutPaths.ToList());
            CollectionAssert.AreEqual(new[] { @"C:\Recent\failing.lnk" }, report.RemainingNewShortcutPaths.ToList());
            CollectionAssert.AreEqual(Array.Empty<string>(), report.DisappearedShortcutPaths.ToList());
            CollectionAssert.AreEqual(new[] { @"C:\Recent\new.lnk" }, report.DeletedShortcutPaths.ToList());
            Assert.AreEqual(1, report.FailedShortcutDeletions.Count);
            Assert.AreEqual(@"C:\Recent\failing.lnk", report.FailedShortcutDeletions[0].Path);
            Assert.IsInstanceOfType(report.FailedShortcutDeletions[0].Error, typeof(IOException));
        }

        [TestMethod]
        public void Unlock_WithCleanup_ReportsDeletedShortcutsSeparatelyFromDisappearedShortcuts()
        {
            var fileSystem = new StubRecentLinkFileSystem(new[] { @"C:\Recent\existing.lnk", @"C:\Recent\new.lnk" });
            var quickAccessLock = new QuickAccessLock(
                QuickAccessLockTarget.RecentFiles,
                @"C:\Recent",
                new[] { @"C:\Recent\existing.lnk", @"C:\Recent\gone.lnk" },
                new[] { new StubHandle() },
                fileSystem);

            var report = quickAccessLock.Unlock(new QuickAccessUnlockOptions { CleanupNewRecentLinks = true });

            CollectionAssert.AreEqual(new[] { @"C:\Recent\new.lnk" }, fileSystem.DeletedPaths);
            CollectionAssert.AreEqual(new[] { @"C:\Recent\new.lnk" }, report.NewShortcutPaths.ToList());
            CollectionAssert.AreEqual(Array.Empty<string>(), report.RemainingNewShortcutPaths.ToList());
            CollectionAssert.AreEqual(new[] { @"C:\Recent\gone.lnk" }, report.DisappearedShortcutPaths.ToList());
            CollectionAssert.AreEqual(new[] { @"C:\Recent\new.lnk" }, report.DeletedShortcutPaths.ToList());
            Assert.AreEqual(0, report.FailedShortcutDeletions.Count);
        }

        [TestMethod]
        public void Dispose_DoesNotEnumerateCurrentSnapshotOrDeleteShortcuts()
        {
            var handle = new StubHandle();
            var fileSystem = new StubRecentLinkFileSystem(Array.Empty<string>())
            {
                ThrowOnEnumerate = true
            };
            var quickAccessLock = new QuickAccessLock(
                QuickAccessLockTarget.RecentFiles,
                @"C:\Recent",
                Array.Empty<string>(),
                new[] { handle },
                fileSystem);

            quickAccessLock.Dispose();

            Assert.IsTrue(handle.IsDisposed);
            Assert.AreEqual(0, fileSystem.DeletedPaths.Count);
        }

        [TestMethod]
        public void Dispose_CalledConcurrently_DisposesHandlesOnce()
        {
            var handle = new StubHandle();
            var quickAccessLock = new QuickAccessLock(
                QuickAccessLockTarget.RecentFiles,
                @"C:\Recent",
                Array.Empty<string>(),
                new[] { handle },
                new StubRecentLinkFileSystem(Array.Empty<string>()));

            Parallel.For(0, 64, _ => quickAccessLock.Dispose());

            Assert.AreEqual(1, handle.DisposeCount);
        }

        [TestMethod]
        public void Unlock_WhenCurrentSnapshotEnumerationFails_DisposesHandles()
        {
            var handle = new StubHandle();
            var fileSystem = new StubRecentLinkFileSystem(Array.Empty<string>())
            {
                ThrowOnEnumerate = true
            };
            var quickAccessLock = new QuickAccessLock(
                QuickAccessLockTarget.RecentFiles,
                @"C:\Recent",
                Array.Empty<string>(),
                new[] { handle },
                fileSystem);

            Assert.ThrowsException<IOException>(() => quickAccessLock.Unlock());

            Assert.IsTrue(handle.IsDisposed, "Handles should be released even when snapshot enumeration fails.");
            Assert.ThrowsException<ObjectDisposedException>(() => quickAccessLock.Unlock());
        }

        [TestMethod]
        public void Unlock_AfterUnlock_ThrowsObjectDisposedException()
        {
            var quickAccessLock = new QuickAccessLock(
                QuickAccessLockTarget.RecentFiles,
                @"C:\Recent",
                Array.Empty<string>(),
                new[] { new StubHandle() },
                new StubRecentLinkFileSystem(Array.Empty<string>()));

            quickAccessLock.Unlock();

            Assert.ThrowsException<ObjectDisposedException>(() => quickAccessLock.Unlock());
        }

        [TestMethod]
        public void Unlock_AfterDispose_ThrowsObjectDisposedException()
        {
            var quickAccessLock = new QuickAccessLock(
                QuickAccessLockTarget.RecentFiles,
                @"C:\Recent",
                Array.Empty<string>(),
                new[] { new StubHandle() },
                new StubRecentLinkFileSystem(Array.Empty<string>()));

            quickAccessLock.Dispose();

            Assert.ThrowsException<ObjectDisposedException>(() => quickAccessLock.Unlock());
        }

        [TestMethod]
        public void Unlock_NullOptions_ThrowsArgumentNullException()
        {
            var quickAccessLock = new QuickAccessLock(
                QuickAccessLockTarget.RecentFiles,
                @"C:\Recent",
                Array.Empty<string>(),
                new[] { new StubHandle() },
                new StubRecentLinkFileSystem(Array.Empty<string>()));

            Assert.ThrowsException<ArgumentNullException>(() => quickAccessLock.Unlock(null));
        }

        private sealed class StubDataFiles : IQuickAccessDataFiles
        {
            public DateTime GetModifiedTimeForScript(PSScript scriptType)
            {
                return DateTime.Now;
            }

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

            public string GetPath()
            {
                return _path;
            }
        }

        private sealed class StubRecentLinkFileSystem : IRecentLinkFileSystem
        {
            private readonly IEnumerable<string> _files;

            public StubRecentLinkFileSystem(IEnumerable<string> files)
            {
                _files = files;
            }

            public bool ThrowOnEnumerate { get; set; }

            public string DeleteFailurePath { get; set; }

            public List<string> DeletedPaths { get; } = new List<string>();

            public IEnumerable<string> EnumerateFiles(string directory)
            {
                if (ThrowOnEnumerate)
                    throw new IOException("enumerate failed");

                return _files;
            }

            public void DeleteFile(string path)
            {
                if (path == DeleteFailurePath)
                    throw new IOException("delete failed");

                DeletedPaths.Add(path);
            }
        }

        private sealed class StubHandleOpener : IQuickAccessBackingFileHandleOpener
        {
            public List<string> OpenedPaths { get; } = new List<string>();

            public List<StubHandle> Handles { get; } = new List<StubHandle>();

            public string ThrowOnPath { get; set; }

            public IQuickAccessBackingFileHandle Open(string path)
            {
                if (path == ThrowOnPath)
                    throw new IOException("open failed");

                OpenedPaths.Add(path);
                var handle = new StubHandle();
                Handles.Add(handle);
                return handle;
            }
        }

        private sealed class StubHandle : IQuickAccessBackingFileHandle
        {
            private int _disposeCount;

            public bool IsDisposed => DisposeCount != 0;

            public int DisposeCount => Volatile.Read(ref _disposeCount);

            public void Dispose()
            {
                Interlocked.Increment(ref _disposeCount);
            }
        }
    }
}
