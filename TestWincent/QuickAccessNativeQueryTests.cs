using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class QuickAccessNativeQueryTests
    {
        [TestMethod]
        public void ForTarget_All_ThrowsArgumentOutOfRangeException()
        {
            Assert.ThrowsException<ArgumentOutOfRangeException>(
                () => QuickAccessNativeQueryMapping.ForTarget(QuickAccess.All));
        }

        [TestMethod]
        public void ForTarget_RecentFiles_UsesQuickAccessNamespaceAndFilesOnlyFilter()
        {
            var spec = QuickAccessNativeQueryMapping.ForTarget(QuickAccess.RecentFiles);

            Assert.AreEqual(ShellNamespaces.QuickAccess, spec.Namespace);
            Assert.AreEqual(QuickAccessNativeQueryFilter.FilesOnly, spec.Filter);
        }

        [TestMethod]
        public void ForTarget_FrequentFolders_UsesFrequentFoldersNamespaceAndFoldersOnlyFilter()
        {
            var spec = QuickAccessNativeQueryMapping.ForTarget(QuickAccess.FrequentFolders);

            Assert.AreEqual(ShellNamespaces.FrequentFolders, spec.Namespace);
            Assert.AreEqual(QuickAccessNativeQueryFilter.FoldersOnly, spec.Filter);
        }

        [TestMethod]
        public void ForTarget_InvalidTarget_ThrowsArgumentOutOfRangeException()
        {
            Assert.ThrowsException<ArgumentOutOfRangeException>(
                () => QuickAccessNativeQueryMapping.ForTarget((QuickAccess)999));
        }

        [TestMethod]
        public void ShouldKeep_AppliesFilter()
        {
            Assert.IsTrue(QuickAccessNativeQueryMapping.ShouldKeep(new FakeShellItem { IsFolder = false }, QuickAccessNativeQueryFilter.All));
            Assert.IsTrue(QuickAccessNativeQueryMapping.ShouldKeep(new FakeShellItem { IsFolder = false }, QuickAccessNativeQueryFilter.FilesOnly));
            Assert.IsFalse(QuickAccessNativeQueryMapping.ShouldKeep(new FakeShellItem { IsFolder = true }, QuickAccessNativeQueryFilter.FilesOnly));
            Assert.IsTrue(QuickAccessNativeQueryMapping.ShouldKeep(new FakeShellItem { IsFolder = true }, QuickAccessNativeQueryFilter.FoldersOnly));
            Assert.IsFalse(QuickAccessNativeQueryMapping.ShouldKeep(new FakeShellItem { IsFolder = false }, QuickAccessNativeQueryFilter.FoldersOnly));
        }

        [TestMethod]
        public void MergeRecentAndFrequent_PreservesRecentFirstAndDeduplicatesPaths()
        {
            var result = QuickAccessQueryMerger.MergeRecentAndFrequent(
                new[] { @"C:\Recent.txt", @"C:\Shared" },
                new[] { @"c:/shared/", @"C:\Folder" });

            CollectionAssert.AreEqual(
                new[] { @"C:\Recent.txt", @"C:\Shared", @"C:\Folder" },
                result.ToList());
        }

        [TestMethod]
        public void MergeRecentAndFrequent_DeduplicatesLargeSyntheticInputs()
        {
            var recent = Enumerable.Range(0, 500)
                .Select(i => i % 2 == 0 ? @"C:\Shared" : $@"C:\Recent{i}.txt")
                .ToList();
            var frequent = Enumerable.Range(0, 500)
                .Select(i => i % 2 == 0 ? @"c:/shared/" : $@"C:\Folder{i}")
                .ToList();

            var result = QuickAccessQueryMerger.MergeRecentAndFrequent(recent, frequent);

            Assert.AreEqual(501, result.Count);
            Assert.AreEqual(@"C:\Shared", result[0]);
            Assert.AreEqual(1, result.Count(path => WindowsPathComparer.Equals(path, @"C:\Shared")));
        }

        [TestMethod]
        public void GetItems_All_MergesRecentAndFrequentFolders()
        {
            var nativeMethods = new Moq.Mock<INativeMethods>(Moq.MockBehavior.Strict);
            nativeMethods.Setup(n => n.CoInitializeEx(IntPtr.Zero, 2))
                .Returns(0);
            nativeMethods.Setup(n => n.CoUninitialize());
            var shellApplication = new FakeShellApplication();
            shellApplication.Folders[ShellNamespaces.QuickAccess] = new FakeFolder(
                new FakeShellItem { IsFolder = false, Path = @"C:\Recent.txt" },
                new FakeShellItem { IsFolder = true, Path = @"C:\PinnedFromQuickAccess" },
                new FakeShellItem { IsFolder = false, Path = @"C:\Shared" });
            shellApplication.Folders[ShellNamespaces.FrequentFolders] = new FakeFolder(
                new FakeShellItem { IsFolder = true, Path = @"c:/shared/" },
                new FakeShellItem { IsFolder = true, Path = @"C:\Folder" });
            var query = new ShellQuickAccessNativeQuery(
                nativeMethods.Object,
                new FakeShellApplicationFactory(shellApplication));

            var result = query.GetItems(QuickAccess.All, TimeSpan.FromSeconds(1));

            CollectionAssert.AreEqual(
                new[] { @"C:\Recent.txt", @"C:\Shared", @"C:\Folder" },
                result.ToList());
            nativeMethods.Verify(n => n.CoUninitialize(), Moq.Times.Once);
        }

        [TestMethod]
        public void GetItems_SkipsEmptyMismatchedAndFailingItems()
        {
            var nativeMethods = new Moq.Mock<INativeMethods>(Moq.MockBehavior.Strict);
            nativeMethods.Setup(n => n.CoInitializeEx(IntPtr.Zero, 2))
                .Returns(0);
            nativeMethods.Setup(n => n.CoUninitialize());
            var shellApplication = new FakeShellApplication();
            shellApplication.Folders[ShellNamespaces.FrequentFolders] = new FakeFolder(
                new FakeShellItem { IsFolder = true, Path = @"C:\Folder" },
                new FakeShellItem { IsFolder = false, Path = @"C:\File.txt" },
                new FakeShellItem { IsFolder = true, Path = " " },
                new ThrowingShellItem());
            var query = new ShellQuickAccessNativeQuery(
                nativeMethods.Object,
                new FakeShellApplicationFactory(shellApplication));

            var result = query.GetItems(QuickAccess.FrequentFolders, TimeSpan.FromSeconds(1));

            CollectionAssert.AreEqual(new[] { @"C:\Folder" }, result.ToList());
            nativeMethods.Verify(n => n.CoUninitialize(), Moq.Times.Once);
        }

        [TestMethod]
        [TestCategory("Integration")]
        [Ignore("Depends on current Windows user Explorer state; enable manually to compare native query with PowerShell fallback.")]
        public void Integration_NativeAndPowerShellQuery_ReturnCompatibleShapes()
        {
            var nativeQuery = new ShellQuickAccessNativeQuery(new DefaultNativeMethods());
            var executor = new ScriptExecutor();

            Console.WriteLine("========== QuickAccess.All ==========");
            AssertCompatibleShape("All",
                nativeQuery.GetItems(QuickAccess.All, TimeSpan.FromSeconds(10)),
                ExecutePowerShellQuery(executor, PSScript.QueryQuickAccess));
            Console.WriteLine("========== QuickAccess.RecentFiles ==========");
            AssertCompatibleShape("RecentFiles",
                nativeQuery.GetItems(QuickAccess.RecentFiles, TimeSpan.FromSeconds(10)),
                ExecutePowerShellQuery(executor, PSScript.QueryRecentFile));
            Console.WriteLine("========== QuickAccess.FrequentFolders ==========");
            AssertCompatibleShape("FrequentFolders",
                nativeQuery.GetItems(QuickAccess.FrequentFolders, TimeSpan.FromSeconds(10)),
                ExecutePowerShellQuery(executor, PSScript.QueryFrequentFolder));
        }

        private static IReadOnlyList<string> ExecutePowerShellQuery(ScriptExecutor executor, PSScript script)
        {
            return executor.ExecutePSScriptWithCache(script, null, 10).GetAwaiter().GetResult();
        }

        private static void AssertCompatibleShape(string label, IReadOnlyList<string> nativeResult, IReadOnlyList<string> powerShellResult)
        {
            Console.WriteLine($"--- Native query ({label}) [{nativeResult.Count} items] ---");
            foreach (var path in nativeResult)
                Console.WriteLine($"  {path}");

            Console.WriteLine($"--- PowerShell query ({label}) [{powerShellResult.Count} items] ---");
            foreach (var path in powerShellResult)
                Console.WriteLine($"  {path}");

            Assert.IsNotNull(nativeResult);
            Assert.IsNotNull(powerShellResult);
            CollectionAssert.AllItemsAreNotNull(nativeResult.ToList());
            CollectionAssert.AllItemsAreNotNull(powerShellResult.ToList());
        }

        public sealed class FakeShellItem
        {
            public bool IsFolder { get; set; }

            public string Path { get; set; }
        }

        public sealed class ThrowingShellItem
        {
            public bool IsFolder
            {
                get { return true; }
            }

            public string Path
            {
                get { throw new InvalidOperationException("Path unavailable."); }
            }
        }

        public sealed class FakeShellApplicationFactory : IShellApplicationFactory
        {
            private readonly FakeShellApplication _shellApplication;

            public FakeShellApplicationFactory(FakeShellApplication shellApplication)
            {
                _shellApplication = shellApplication;
            }

            public object CreateShellApplication()
            {
                return _shellApplication;
            }
        }

        public sealed class FakeShellApplication
        {
            public Dictionary<string, FakeFolder> Folders { get; } = new Dictionary<string, FakeFolder>();

            public object Namespace(string name)
            {
                FakeFolder folder;
                return Folders.TryGetValue(name, out folder) ? folder : null;
            }
        }

        public sealed class FakeFolder
        {
            private readonly object[] _items;

            public FakeFolder(params object[] items)
            {
                _items = items;
            }

            public FakeFolderItems Items()
            {
                return new FakeFolderItems(_items);
            }
        }

        public sealed class FakeFolderItems
        {
            private readonly object[] _items;

            public FakeFolderItems(object[] items)
            {
                _items = items;
            }

            public int Count
            {
                get { return _items.Length; }
            }

            public object Item(int index)
            {
                return _items[index];
            }
        }
    }
}
