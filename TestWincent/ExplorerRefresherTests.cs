using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class ExplorerRefresherTests
    {
        [TestMethod]
        public void IsRecentAccessLocation_MatchesQuickAccessGuid()
        {
            var location = new ExplorerLocation(
                string.Empty,
                "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}");

            Assert.IsTrue(ShellExplorerRefresher.IsRecentAccessLocation(location));
        }

        [TestMethod]
        public void IsRecentAccessLocation_MatchesHomeGuid()
        {
            var location = new ExplorerLocation(
                string.Empty,
                "shell:::{f874310e-b6b7-47dc-bc84-b9e6b38f5903}");

            Assert.IsTrue(ShellExplorerRefresher.IsRecentAccessLocation(location));
        }

        [TestMethod]
        public void IsRecentAccessLocation_DoesNotUseLocalizedName()
        {
            var location = new ExplorerLocation("Quick Access", string.Empty);

            Assert.IsFalse(ShellExplorerRefresher.IsRecentAccessLocation(location));
        }

        [TestMethod]
        public void IsProbableExplorerLocation_MatchesFileUrl()
        {
            var location = new ExplorerLocation("Documents", "file:///C:/Users/Test/Documents");

            Assert.IsTrue(ShellExplorerRefresher.IsProbableExplorerLocation(location));
        }

        [TestMethod]
        public void IsProbableExplorerLocation_MatchesNamedWindowWithEmptyUrl()
        {
            var location = new ExplorerLocation("Home", string.Empty);

            Assert.IsTrue(ShellExplorerRefresher.IsProbableExplorerLocation(location));
        }

        [TestMethod]
        public void IsProbableExplorerLocation_RejectsWebUrl()
        {
            var location = new ExplorerLocation("Example", "https://example.com");

            Assert.IsFalse(ShellExplorerRefresher.IsProbableExplorerLocation(location));
        }

        [TestMethod]
        public void Refresh_RefreshesQuickAccessWindowBeforeOtherExplorerWindows()
        {
            var quickAccess = new FakeWindow("Home", "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}");
            var documents = new FakeWindow("Documents", "file:///C:/Users/Test/Documents");
            var refresher = CreateRefresher(quickAccess, documents);

            refresher.Refresh(TimeSpan.FromSeconds(10));

            Assert.AreEqual(1, quickAccess.RefreshCalls);
            Assert.AreEqual(0, documents.RefreshCalls);
        }

        [TestMethod]
        public void Refresh_WhenQuickAccessWindowRefreshFails_ThrowsWithoutRefreshingBroaderExplorerWindows()
        {
            var quickAccess = new FakeWindow(
                "Home",
                "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}",
                refreshThrows: true,
                documentRefreshThrows: true);
            var documents = new FakeWindow("Documents", "file:///C:/Users/Test/Documents");
            var refresher = CreateRefresher(quickAccess, documents);

            Assert.ThrowsException<QuickAccessOperationException>(
                () => refresher.Refresh(TimeSpan.FromSeconds(10)));

            Assert.AreEqual(1, quickAccess.RefreshCalls);
            Assert.AreEqual(1, quickAccess.Document.RefreshCalls);
            Assert.AreEqual(0, documents.RefreshCalls);
        }

        [TestMethod]
        public void Refresh_WhenNoQuickAccessWindow_RefreshesProbableExplorerWindow()
        {
            var browser = new FakeWindow("Browser", "https://example.com");
            var documents = new FakeWindow("Documents", "file:///C:/Users/Test/Documents");
            var refresher = CreateRefresher(browser, documents);

            refresher.Refresh(TimeSpan.FromSeconds(10));

            Assert.AreEqual(0, browser.RefreshCalls);
            Assert.AreEqual(1, documents.RefreshCalls);
        }

        [TestMethod]
        [TestCategory("Integration")]
        [Ignore("Performs a real Explorer refresh via COM Shell.Application; enable manually to verify current system behavior.")]
        public void Integration_RefreshRealExplorerWindows()
        {
            var refresher = new ShellExplorerRefresher(new DefaultNativeMethods());
            var stopwatch = Stopwatch.StartNew();

            refresher.Refresh(TimeSpan.FromSeconds(10));

            Console.WriteLine($"Refresh completed in {stopwatch.Elapsed.TotalSeconds:0.###} seconds.");
        }

        private static ShellExplorerRefresher CreateRefresher(params FakeWindow[] windows)
        {
            var nativeMethods = new Mock<INativeMethods>(MockBehavior.Strict);
            nativeMethods.Setup(n => n.CoInitializeEx(It.IsAny<IntPtr>(), It.IsAny<uint>())).Returns(0);
            nativeMethods.Setup(n => n.CoUninitialize());

            return new ShellExplorerRefresher(
                nativeMethods.Object,
                new FakeShellApplicationFactory(windows));
        }

        public sealed class FakeShellApplicationFactory : IShellApplicationFactory
        {
            private readonly FakeShellApplication _shellApplication;

            public FakeShellApplicationFactory(IEnumerable<FakeWindow> windows)
            {
                _shellApplication = new FakeShellApplication(windows);
            }

            public object CreateShellApplication()
            {
                return _shellApplication;
            }
        }

        public sealed class FakeShellApplication
        {
            private readonly FakeWindows _windows;

            public FakeShellApplication(IEnumerable<FakeWindow> windows)
            {
                _windows = new FakeWindows(windows);
            }

            public FakeWindows Windows()
            {
                return _windows;
            }
        }

        public sealed class FakeWindows
        {
            private readonly List<FakeWindow> _windows;

            public FakeWindows(IEnumerable<FakeWindow> windows)
            {
                _windows = new List<FakeWindow>(windows);
            }

            public int Count => _windows.Count;

            public FakeWindow Item(int index)
            {
                return _windows[index];
            }
        }

        public sealed class FakeWindow
        {
            private readonly bool _refreshThrows;

            public FakeWindow(
                string locationName,
                string locationUrl,
                bool refreshThrows = false,
                bool documentRefreshThrows = false)
            {
                LocationName = locationName;
                LocationURL = locationUrl;
                _refreshThrows = refreshThrows;
                Document = new FakeDocument(documentRefreshThrows);
            }

            public string LocationName { get; }

            public string LocationURL { get; }

            public FakeDocument Document { get; }

            public int RefreshCalls { get; private set; }

            public void Refresh()
            {
                RefreshCalls++;
                if (_refreshThrows)
                    throw new InvalidOperationException("refresh failed");
            }
        }

        public sealed class FakeDocument
        {
            private readonly bool _refreshThrows;

            public FakeDocument(bool refreshThrows)
            {
                _refreshThrows = refreshThrows;
            }

            public int RefreshCalls { get; private set; }

            public void Refresh()
            {
                RefreshCalls++;
                if (_refreshThrows)
                    throw new InvalidOperationException("document refresh failed");
            }
        }
    }
}
