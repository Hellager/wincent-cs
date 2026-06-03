using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.Linq;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class QuickAccessNativeMutationTests
    {
        [TestMethod]
        public void UnpinFrequentFolder_UnpinnedWin10_PinsThenUnpinsFromFrequentNamespace()
        {
            var shellApplication = new FakeShellApplication();
            var frequentItem = shellApplication.AddFrequentFolder(@"C:\Folder", false);
            var mutation = CreateMutation(shellApplication, false);

            mutation.UnpinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(5));

            Assert.IsFalse(shellApplication.ContainsFrequentFolder(@"C:\Folder"));
            CollectionAssert.AreEqual(new[] { "pintohome" }, shellApplication.SelfItem(@"C:\Folder").InvokedVerbs.ToList());
            CollectionAssert.AreEqual(new[] { "unpinfromhome", "unpinfromhome" }, frequentItem.InvokedVerbs.ToList());
        }

        [TestMethod]
        public void UnpinFrequentFolder_UnpinnedWin11_TogglesSelfTwice()
        {
            var shellApplication = new FakeShellApplication();
            var frequentItem = shellApplication.AddFrequentFolder(@"C:\Folder", false);
            var mutation = CreateMutation(shellApplication, true);

            mutation.UnpinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(5));

            Assert.IsFalse(shellApplication.ContainsFrequentFolder(@"C:\Folder"));
            CollectionAssert.AreEqual(
                new[] { "pintohome", "pintohome" },
                shellApplication.SelfItem(@"C:\Folder").InvokedVerbs.ToList());
            CollectionAssert.AreEqual(new[] { "unpinfromhome" }, frequentItem.InvokedVerbs.ToList());
        }

        [TestMethod]
        public void UnpinFrequentFolder_PinnedFolderRemovedByDirectUnpin_DoesNotPinFirst()
        {
            var shellApplication = new FakeShellApplication();
            var frequentItem = shellApplication.AddFrequentFolder(@"C:\Folder", true);
            var mutation = CreateMutation(shellApplication, false);

            mutation.UnpinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(5));

            Assert.IsFalse(shellApplication.ContainsFrequentFolder(@"C:\Folder"));
            CollectionAssert.AreEqual(new[] { "unpinfromhome" }, frequentItem.InvokedVerbs.ToList());
            CollectionAssert.AreEqual(Array.Empty<string>(), shellApplication.SelfItem(@"C:\Folder").InvokedVerbs.ToList());
        }

        [TestMethod]
        public void UnpinFrequentFolder_VerificationStillContainsTarget_Throws()
        {
            var shellApplication = new FakeShellApplication { IgnoreRemoveVerbs = true };
            shellApplication.AddFrequentFolder(@"C:\Folder", false);
            var mutation = CreateMutation(shellApplication, false);

            Assert.ThrowsException<InvalidOperationException>(
                () => mutation.UnpinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(5)));
        }

        private static ShellQuickAccessNativeMutation CreateMutation(
            FakeShellApplication shellApplication,
            bool isWindows11OrLater)
        {
            var nativeMethods = new Mock<INativeMethods>(MockBehavior.Strict);
            nativeMethods.Setup(n => n.CoInitializeEx(IntPtr.Zero, 2)).Returns(0);
            nativeMethods.Setup(n => n.CoUninitialize());

            return new ShellQuickAccessNativeMutation(
                nativeMethods.Object,
                new FakeShellApplicationFactory(shellApplication),
                () => isWindows11OrLater,
                TimeSpan.FromMilliseconds(1));
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
            private readonly List<FakeFrequentItem> _frequentItems = new List<FakeFrequentItem>();
            private readonly Dictionary<string, FakeSelfItem> _selfItems =
                new Dictionary<string, FakeSelfItem>(StringComparer.OrdinalIgnoreCase);

            public bool IgnoreRemoveVerbs { get; set; }

            public FakeFrequentItem AddFrequentFolder(string path, bool isPinned)
            {
                var item = new FakeFrequentItem(this, path, isPinned);
                _frequentItems.Add(item);
                SelfItem(path);
                return item;
            }

            public bool ContainsFrequentFolder(string path)
            {
                return _frequentItems.Any(item => WindowsPathComparer.Equals(item.Path, path));
            }

            public FakeFrequentItem FindFrequentFolder(string path)
            {
                return _frequentItems.FirstOrDefault(item => WindowsPathComparer.Equals(item.Path, path));
            }

            public FakeSelfItem SelfItem(string path)
            {
                FakeSelfItem item;
                if (!_selfItems.TryGetValue(path, out item))
                {
                    item = new FakeSelfItem(this, path);
                    _selfItems.Add(path, item);
                }

                return item;
            }

            public object Namespace(string name)
            {
                if (name == ShellNamespaces.FrequentFolders)
                    return new FakeFolder(_frequentItems);

                if (_selfItems.ContainsKey(name))
                    return new FakeSelfFolder(SelfItem(name));

                return null;
            }

            public void PinFrequentFolder(string path)
            {
                var existing = _frequentItems.FirstOrDefault(item => WindowsPathComparer.Equals(item.Path, path));
                if (existing != null)
                {
                    existing.IsPinned = true;
                    return;
                }

                _frequentItems.Add(new FakeFrequentItem(this, path, true));
            }

            public void RemoveFrequentFolder(string path)
            {
                if (IgnoreRemoveVerbs)
                    return;

                _frequentItems.RemoveAll(item => WindowsPathComparer.Equals(item.Path, path));
            }
        }

        public sealed class FakeFolder
        {
            private readonly List<FakeFrequentItem> _items;

            public FakeFolder(List<FakeFrequentItem> items)
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
            private readonly List<FakeFrequentItem> _items;

            public FakeFolderItems(List<FakeFrequentItem> items)
            {
                _items = items;
            }

            public int Count
            {
                get { return _items.Count; }
            }

            public object Item(int index)
            {
                return _items[index];
            }
        }

        public sealed class FakeSelfFolder
        {
            public FakeSelfFolder(FakeSelfItem self)
            {
                Self = self;
            }

            public FakeSelfItem Self { get; }
        }

        public sealed class FakeFrequentItem
        {
            private readonly FakeShellApplication _shellApplication;

            public FakeFrequentItem(FakeShellApplication shellApplication, string path, bool isPinned)
            {
                _shellApplication = shellApplication;
                Path = path;
                IsPinned = isPinned;
            }

            public string Path { get; }

            public bool IsPinned { get; set; }

            public List<string> InvokedVerbs { get; } = new List<string>();

            public void InvokeVerb(string verb)
            {
                InvokedVerbs.Add(verb);
                if (verb == "unpinfromhome" && IsPinned)
                    _shellApplication.RemoveFrequentFolder(Path);
            }
        }

        public sealed class FakeSelfItem
        {
            private readonly FakeShellApplication _shellApplication;

            public FakeSelfItem(FakeShellApplication shellApplication, string path)
            {
                _shellApplication = shellApplication;
                Path = path;
            }

            public string Path { get; }

            public List<string> InvokedVerbs { get; } = new List<string>();

            public void InvokeVerb(string verb)
            {
                InvokedVerbs.Add(verb);
                if (verb != "pintohome")
                    return;

                if (!_shellApplication.ContainsFrequentFolder(Path))
                {
                    _shellApplication.PinFrequentFolder(Path);
                    return;
                }

                var item = _shellApplication.FindFrequentFolder(Path);
                if (item != null && item.IsPinned)
                    _shellApplication.RemoveFrequentFolder(Path);
                else if (item != null)
                    item.IsPinned = true;
            }
        }
    }
}
