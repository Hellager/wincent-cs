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
        public void ForTarget_All_UsesQuickAccessNamespaceWithoutFilter()
        {
            var spec = QuickAccessNativeQueryMapping.ForTarget(QuickAccess.All);

            Assert.AreEqual(ShellNamespaces.QuickAccess, spec.Namespace);
            Assert.AreEqual(QuickAccessNativeQueryFilter.All, spec.Filter);
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

            var result = query.GetItems(QuickAccess.FrequentFolders);

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

            AssertCompatibleShape(nativeQuery.GetItems(QuickAccess.All), ExecutePowerShellQuery(executor, PSScript.QueryQuickAccess));
            AssertCompatibleShape(nativeQuery.GetItems(QuickAccess.RecentFiles), ExecutePowerShellQuery(executor, PSScript.QueryRecentFile));
            AssertCompatibleShape(nativeQuery.GetItems(QuickAccess.FrequentFolders), ExecutePowerShellQuery(executor, PSScript.QueryFrequentFolder));
        }

        private static IReadOnlyList<string> ExecutePowerShellQuery(ScriptExecutor executor, PSScript script)
        {
            return executor.ExecutePSScriptWithCache(script, null, 10).GetAwaiter().GetResult();
        }

        private static void AssertCompatibleShape(IReadOnlyList<string> nativeResult, IReadOnlyList<string> powerShellResult)
        {
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
