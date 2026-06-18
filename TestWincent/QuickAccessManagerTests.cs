using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class QuickAccessManagerTests
    {
        private Mock<IScriptExecutor> _executor;
        private Mock<IFileSystemOperations> _fileSystem;
        private Mock<INativeMethods> _nativeMethods;
        private Mock<IQuickAccessDataFiles> _dataFiles;
        private Mock<IQuickAccessNativeMutation> _nativeMutation;
        private Mock<IExplorerRefresher> _explorerRefresher;
        private Mock<IRecentLinksCleaner> _recentLinksCleaner;
        private QuickAccessManager _manager;

        [TestInitialize]
        public void SetUp()
        {
            _executor = new Mock<IScriptExecutor>(MockBehavior.Strict);
            _fileSystem = new Mock<IFileSystemOperations>(MockBehavior.Strict);
            _nativeMethods = new Mock<INativeMethods>(MockBehavior.Strict);
            _dataFiles = new Mock<IQuickAccessDataFiles>(MockBehavior.Strict);
            _nativeMutation = new Mock<IQuickAccessNativeMutation>(MockBehavior.Strict);
            _explorerRefresher = new Mock<IExplorerRefresher>(MockBehavior.Strict);
            _recentLinksCleaner = new Mock<IRecentLinksCleaner>(MockBehavior.Strict);

            _executor.Setup(e => e.ExecutePSScriptWithCache(It.IsAny<PSScript>(), It.IsAny<string>(), It.IsAny<int>()))
                .ReturnsAsync(new List<string>());
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(It.IsAny<PSScript>(), It.IsAny<string>(), It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(0, string.Empty, string.Empty));
            _executor.Setup(e => e.ClearCache());
            _executor.Setup(e => e.Dispose());

            _fileSystem.Setup(f => f.FileExists(It.IsAny<string>())).Returns(true);
            _fileSystem.Setup(f => f.DirectoryExists(It.IsAny<string>())).Returns(true);
            _fileSystem.Setup(f => f.DeleteFile(It.IsAny<string>()));

            _nativeMethods.Setup(n => n.CoInitializeEx(It.IsAny<IntPtr>(), It.IsAny<uint>())).Returns(0);
            _nativeMethods.Setup(n => n.CoUninitialize());
            _nativeMethods.Setup(n => n.SHAddToRecentDocs(It.IsAny<uint>(), It.IsAny<IntPtr>()));
            _nativeMethods.Setup(n => n.CoTaskMemFree(It.IsAny<IntPtr>()));

            IntPtr recentFolder = Marshal.StringToHGlobalUni(@"C:\Users\Test\AppData\Roaming\Microsoft\Windows\Recent");
            _nativeMethods.Setup(n => n.SHGetKnownFolderPath(
                    It.IsAny<Guid>(),
                    It.IsAny<uint>(),
                    It.IsAny<IntPtr>(),
                    out It.Ref<IntPtr>.IsAny))
                .Callback(new SHGetKnownFolderPathCallback((Guid id, uint flags, IntPtr token, out IntPtr path) =>
                {
                    path = recentFolder;
                }))
                .Returns(0);

            _dataFiles.Setup(d => d.RemoveRecentFile());
            _dataFiles.Setup(d => d.GetModifiedTimeForScript(It.IsAny<PSScript>())).Returns(DateTime.Now);
            _dataFiles.Setup(d => d.RecentFilesPath).Returns(@"C:\recent.automaticDestinations-ms");
            _dataFiles.Setup(d => d.FrequentFoldersPath).Returns(@"C:\frequent.automaticDestinations-ms");

            _manager = new QuickAccessManager(
                _executor.Object,
                TimeSpan.FromSeconds(10),
                _fileSystem.Object,
                _nativeMethods.Object,
                _dataFiles.Object,
                RetryPolicy.Standard,
                new PowerShellFallbackNativeQuery(),
                _nativeMutation.Object,
                _explorerRefresher.Object,
                _recentLinksCleaner.Object);
        }

        [TestCleanup]
        public void CleanUp()
        {
            _manager.Dispose();
        }

        private delegate void SHGetKnownFolderPathCallback(Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr ppszPath);

        [TestMethod]
        public void GetItems_All_UsesMergedRecentAndFrequentFallback()
        {
            var expected = new List<string> { @"C:\a.txt", @"C:\shared", @"C:\folder" };
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { @"C:\a.txt", @"C:\shared" });
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string> { @"c:/shared/", @"C:\folder" });

            var result = _manager.GetItems(QuickAccess.All);

            CollectionAssert.AreEqual(expected, result.ToList());
            _executor.Verify(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10), Times.Once);
            _executor.Verify(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10), Times.Once);
            _executor.Verify(e => e.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null, 10), Times.Never);
        }

        [TestMethod]
        public void GetItems_NativeSuccess_DoesNotUsePowerShellFallback()
        {
            var nativeQuery = new Mock<IQuickAccessNativeQuery>(MockBehavior.Strict);
            nativeQuery.Setup(n => n.GetItems(QuickAccess.RecentFiles, TimeSpan.FromSeconds(10)))
                .Returns(new List<string> { @"C:\native.txt" });
            _manager.Dispose();
            _manager = CreateManager(nativeQuery.Object);

            var result = _manager.GetItems(QuickAccess.RecentFiles);

            CollectionAssert.AreEqual(new[] { @"C:\native.txt" }, result.ToList());
            _executor.Verify(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, It.IsAny<int>()), Times.Never);
            nativeQuery.Verify(n => n.GetItems(QuickAccess.RecentFiles, TimeSpan.FromSeconds(10)), Times.Once);
        }

        [TestMethod]
        public void GetItemPaths_NativeSuccess_ReturnsPathObjects()
        {
            var nativeQuery = new Mock<IQuickAccessNativeQuery>(MockBehavior.Strict);
            nativeQuery.Setup(n => n.GetItems(QuickAccess.RecentFiles, TimeSpan.FromSeconds(10)))
                .Returns(new List<string> { @"C:\native.txt", @"C:\missing.txt" });
            _manager.Dispose();
            _manager = CreateManager(nativeQuery.Object);

            var result = _manager.GetItemPaths(QuickAccess.RecentFiles);

            CollectionAssert.AreEqual(
                new[] { @"C:\native.txt", @"C:\missing.txt" },
                result.Select(path => path.FullName).ToList());
            Assert.AreEqual(@"C:\native.txt", result[0].ToString());
            nativeQuery.Verify(n => n.GetItems(QuickAccess.RecentFiles, TimeSpan.FromSeconds(10)), Times.Once);
            _executor.Verify(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, It.IsAny<int>()), Times.Never);
        }

        [TestMethod]
        public void QuickAccessPath_EqualityUsesOrdinalPathString()
        {
            var path = new QuickAccessPath(@"C:\Folder");

            Assert.AreEqual(new QuickAccessPath(@"C:\Folder"), path);
            Assert.AreNotEqual(new QuickAccessPath(@"c:\folder"), path);
            Assert.AreEqual(@"C:\Folder", path.FullName);
        }

        [TestMethod]
        public void GetItems_NativeFailure_FallsBackToPowerShell()
        {
            var nativeQuery = new Mock<IQuickAccessNativeQuery>(MockBehavior.Strict);
            nativeQuery.Setup(n => n.GetItems(QuickAccess.FrequentFolders, It.IsAny<TimeSpan>()))
                .Throws(new InvalidOperationException("native failed"));
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string> { @"C:\fallback" });
            _manager.Dispose();
            _manager = CreateManager(nativeQuery.Object);

            var result = _manager.GetItems(QuickAccess.FrequentFolders);

            CollectionAssert.AreEqual(new[] { @"C:\fallback" }, result.ToList());
            _executor.Verify(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10), Times.Once);
        }

        [TestMethod]
        public void GetItems_TransientFallbackFailure_RetriesNativeQuery()
        {
            var nativeQuery = new Mock<IQuickAccessNativeQuery>(MockBehavior.Strict);
            nativeQuery.SetupSequence(n => n.GetItems(QuickAccess.FrequentFolders, TimeSpan.FromSeconds(10)))
                .Throws(new InvalidOperationException("Failed to open shell namespace: shell:::broken"))
                .Returns(new List<string> { @"C:\native-after-retry" });
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .Throws(CreatePowerShellException(PowerShellOperation.QueryFrequentFolders, PowerShellErrorKind.Timeout, "operation timed out"));
            _manager.Dispose();
            _manager = CreateManager(
                nativeQuery.Object,
                new RetryPolicy(1, TimeSpan.FromMilliseconds(1), TimeSpan.FromMilliseconds(1), 1.0, false));

            var result = _manager.GetItems(QuickAccess.FrequentFolders);

            CollectionAssert.AreEqual(new[] { @"C:\native-after-retry" }, result.ToList());
            nativeQuery.Verify(n => n.GetItems(QuickAccess.FrequentFolders, TimeSpan.FromSeconds(10)), Times.Exactly(2));
            _executor.Verify(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10), Times.Once);
        }

        [TestMethod]
        public void GetItems_NativeAndPowerShellFailure_ThrowsPowerShellExecutionException()
        {
            var nativeQuery = new Mock<IQuickAccessNativeQuery>(MockBehavior.Strict);
            nativeQuery.Setup(n => n.GetItems(QuickAccess.All, It.IsAny<TimeSpan>()))
                .Throws(new InvalidOperationException("native failed"));
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .Throws(CreatePowerShellException(PowerShellOperation.QueryRecentFiles, PowerShellErrorKind.ProcessFailed, "fallback failed"));
            _manager.Dispose();
            _manager = CreateManager(nativeQuery.Object);

            var exception = Assert.ThrowsException<PowerShellExecutionException>(
                () => _manager.GetItems(QuickAccess.All));

            Assert.AreEqual(PowerShellOperation.QueryRecentFiles, exception.Operation);
        }

        [TestMethod]
        public void QuickAccess_EnumValues_PreserveLegacyNumericValues()
        {
            Assert.AreEqual(0, (int)QuickAccess.All);
            Assert.AreEqual(1, (int)QuickAccess.RecentFiles);
            Assert.AreEqual(2, (int)QuickAccess.FrequentFolders);
        }

        [TestMethod]
        public void ContainsItem_UsesCaseSensitiveSubstring()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { @"C:\CaseSensitive.txt" });

            Assert.IsTrue(_manager.ContainsItem("Sensitive", QuickAccess.RecentFiles));
            Assert.IsFalse(_manager.ContainsItem("sensitive", QuickAccess.RecentFiles));
        }

        [TestMethod]
        public void ContainsItem_NativeResult_UsesCaseSensitiveSubstring()
        {
            var nativeQuery = new Mock<IQuickAccessNativeQuery>(MockBehavior.Strict);
            nativeQuery.Setup(n => n.GetItems(QuickAccess.RecentFiles, It.IsAny<TimeSpan>()))
                .Returns(new List<string> { @"C:\NativeCaseSensitive.txt" });
            _manager.Dispose();
            _manager = CreateManager(nativeQuery.Object);

            Assert.IsTrue(_manager.ContainsItem("Sensitive", QuickAccess.RecentFiles));
            Assert.IsFalse(_manager.ContainsItem("sensitive", QuickAccess.RecentFiles));
        }

        [TestMethod]
        public void ContainsItemExact_UsesWindowsPathSemantics()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string> { @"C:\Folder\" });

            Assert.IsTrue(_manager.ContainsItemExact(@"c:/folder", QuickAccess.FrequentFolders));
        }

        [TestMethod]
        public void ContainsItemExact_NativeResult_UsesWindowsPathSemantics()
        {
            var nativeQuery = new Mock<IQuickAccessNativeQuery>(MockBehavior.Strict);
            nativeQuery.Setup(n => n.GetItems(QuickAccess.FrequentFolders, It.IsAny<TimeSpan>()))
                .Returns(new List<string> { @"C:\Folder\" });
            _manager.Dispose();
            _manager = CreateManager(nativeQuery.Object);

            Assert.IsTrue(_manager.ContainsItemExact(@"c:/folder", QuickAccess.FrequentFolders));
        }

        [TestMethod]
        public void ValidatePath_ProtectedPath_IsAllowedWhenItExists()
        {
            _fileSystem.Setup(f => f.FileExists(@"C:\Windows\System32\drivers\etc\hosts")).Returns(true);

            QuickAccessManager.ValidatePath(@"C:\Windows\System32\drivers\etc\hosts", PathType.File, _fileSystem.Object);
        }

        [TestMethod]
        [ExpectedException(typeof(QuickAccessItemAlreadyExistsException))]
        public void AddItem_ExistingItem_ThrowsSpecificException()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { @"C:\test.txt" });

            _manager.AddItem(@"C:\test.txt", QuickAccess.RecentFiles);
        }

        [TestMethod]
        public void AddItem_MissingPath_ValidatesPathBeforeBusinessState()
        {
            _fileSystem.Setup(f => f.FileExists(@"C:\missing.txt")).Returns(false);

            Assert.ThrowsException<FileNotFoundException>(
                () => _manager.AddItem(@"C:\missing.txt", QuickAccess.RecentFiles));

            _executor.Verify(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, It.IsAny<int>()), Times.Never);
        }

        [TestMethod]
        public void AddItem_RecentFile_UsesNativeRecentDocs()
        {
            ApartmentState? recentDocsApartment = null;
            _nativeMethods.Setup(n => n.SHAddToRecentDocs(It.IsAny<uint>(), It.IsAny<IntPtr>()))
                .Callback<uint, IntPtr>((flags, path) => recentDocsApartment = Thread.CurrentThread.GetApartmentState());
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string>());

            _manager.AddItem(@"C:\test.txt", QuickAccess.RecentFiles);

            Assert.AreEqual(ApartmentState.STA, recentDocsApartment.Value);
            _nativeMethods.Verify(n => n.CoInitializeEx(
                IntPtr.Zero,
                (uint)(NativeMethods.COINIT_APARTMENTTHREADED | NativeMethods.COINIT_DISABLE_OLE1DDE)),
                Times.Once);
            _nativeMethods.Verify(n => n.SHAddToRecentDocs(It.IsAny<uint>(), It.IsAny<IntPtr>()), Times.Once);
            _executor.Verify(e => e.ClearCache(), Times.Once);
        }

        [TestMethod]
        public void AddItem_RecentFileWithRefresh_RemovesRecentBackingFile()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string>());

            _manager.AddItem(@"C:\test.txt", QuickAccess.RecentFiles, new AddOptions { RefreshRecentFiles = true });

            _dataFiles.Verify(d => d.RemoveRecentFile(), Times.Once);
        }

        [TestMethod]
        public void AddItem_WithRefreshExplorer_RefreshesAfterSuccessfulAdd()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string>());
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)));

            _manager.AddItem(
                @"C:\test.txt",
                QuickAccess.RecentFiles,
                new AddOptions { RefreshExplorer = true });

            _explorerRefresher.Verify(r => r.Refresh(TimeSpan.FromSeconds(10)), Times.Once);
            _executor.Verify(e => e.ClearCache(), Times.Once);
        }

        [TestMethod]
        public void AddItem_RefreshExplorerFailure_ThrowsPostMutationExceptionAndClearsCache()
        {
            var refreshError = new InvalidOperationException("refresh failed");
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string>());
            _nativeMutation.Setup(m => m.PinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)));
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)))
                .Throws(refreshError);
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, 10))
                .Throws(CreatePowerShellException(PowerShellOperation.RefreshExplorer, PowerShellErrorKind.ProcessFailed, "refresh failed"));

            var ex = Assert.ThrowsException<QuickAccessPostMutationException>(
                () => _manager.AddItem(
                    @"C:\Folder",
                    QuickAccess.FrequentFolders,
                    new AddOptions { RefreshExplorer = true }));

            Assert.AreEqual(@"C:\Folder", ex.Path);
            Assert.AreEqual(QuickAccess.FrequentFolders, ex.Target);
            Assert.AreEqual(QuickAccessPostMutationStep.RefreshExplorer, ex.Step);
            Assert.IsInstanceOfType(ex.InnerException, typeof(PowerShellExecutionException));
            _executor.Verify(e => e.ClearCache(), Times.Once);
        }

        [TestMethod]
        public void AddItem_FrequentFolder_UsesNativePin()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string>());
            _nativeMutation.Setup(m => m.PinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)));

            _manager.AddItem(@"C:\Folder", QuickAccess.FrequentFolders);

            _nativeMutation.Verify(m => m.PinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)), Times.Once);
            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.PinToFrequentFolder, @"C:\Folder", 10), Times.Never);
        }

        [TestMethod]
        public void AddItem_FrequentFolder_NativeFailureFallsBackToPowerShell()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string>());
            _nativeMutation.Setup(m => m.PinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)))
                .Throws(new COMException("Shell verb failed.", unchecked((int)0x80004005)));

            _manager.AddItem(@"C:\Folder", QuickAccess.FrequentFolders);

            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.PinToFrequentFolder, @"C:\Folder", 10), Times.Once);
        }

        [TestMethod]
        public void AddItem_FrequentFolder_NonRecoverableNativeFailureDoesNotFallBack()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string>());
            _nativeMutation.Setup(m => m.PinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)))
                .Throws(new InvalidOperationException("native failed"));

            var ex = Assert.ThrowsException<InvalidOperationException>(
                () => _manager.AddItem(@"C:\Folder", QuickAccess.FrequentFolders));

            Assert.AreEqual("native failed", ex.Message);
            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.PinToFrequentFolder, @"C:\Folder", 10), Times.Never);
            _executor.Verify(e => e.ClearCache(), Times.Never);
        }

        [TestMethod]
        public void AddItem_FrequentFolder_ShellNamespaceFailureFallsBackToPowerShell()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string>());
            _nativeMutation.Setup(m => m.PinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)))
                .Throws(new InvalidOperationException("Failed to open shell namespace: shell:::broken"));

            _manager.AddItem(@"C:\Folder", QuickAccess.FrequentFolders);

            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.PinToFrequentFolder, @"C:\Folder", 10), Times.Once);
        }

        [TestMethod]
        public void AddItem_FrequentFolder_PowerShellAlreadyExistsSentinel_ThrowsAlreadyExists()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string>());
            _nativeMutation.Setup(m => m.PinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)))
                .Throws(new COMException("Shell verb failed.", unchecked((int)0x80004005)));
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.PinToFrequentFolder, @"C:\Folder", 10))
                .Throws(new QuickAccessItemAlreadyExistsException(@"C:\Folder", QuickAccess.FrequentFolders));

            var ex = Assert.ThrowsException<QuickAccessItemAlreadyExistsException>(
                () => _manager.AddItem(@"C:\Folder", QuickAccess.FrequentFolders));

            Assert.AreEqual(@"C:\Folder", ex.Path);
            Assert.AreEqual(QuickAccess.FrequentFolders, ex.Target);
            _executor.Verify(e => e.ClearCache(), Times.Never);
        }

        [TestMethod]
        public void AddItem_FrequentFolder_NativeTimeoutDoesNotFallBackToPowerShell()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string>());
            _nativeMutation.Setup(m => m.PinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)))
                .Throws(new TimeoutException("native timed out"));

            Assert.ThrowsException<TimeoutException>(() =>
                _manager.AddItem(@"C:\Folder", QuickAccess.FrequentFolders));

            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.PinToFrequentFolder, @"C:\Folder", 10), Times.Never);
            _executor.Verify(e => e.ClearCache(), Times.Never);
        }

        [TestMethod]
        [ExpectedException(typeof(UnsupportedQuickAccessOperationException))]
        public void AddItem_All_ThrowsUnsupportedOperation()
        {
            _manager.AddItem(@"C:\test.txt", QuickAccess.All);
        }

        [TestMethod]
        [ExpectedException(typeof(QuickAccessItemNotFoundException))]
        public void RemoveItem_MissingItem_ThrowsSpecificException()
        {
            _fileSystem.Setup(f => f.FileExists(@"C:\test.txt")).Returns(true);
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string>());

            _manager.RemoveItem(@"C:\test.txt", QuickAccess.RecentFiles);
        }

        [TestMethod]
        public void RemoveItem_EmptyPath_ValidatesPathBeforeBusinessState()
        {
            Assert.ThrowsException<ArgumentException>(
                () => _manager.RemoveItem(string.Empty, QuickAccess.RecentFiles));

            _executor.Verify(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, It.IsAny<int>()), Times.Never);
        }

        [TestMethod]
        public void RemoveItem_RecentFile_UsesNativeRemove()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { @"C:\test.txt" });
            _nativeMutation.Setup(m => m.RemoveRecentFile(@"C:\test.txt", TimeSpan.FromSeconds(10)));

            _manager.RemoveItem(@"C:\test.txt", QuickAccess.RecentFiles);

            _nativeMutation.Verify(m => m.RemoveRecentFile(@"C:\test.txt", TimeSpan.FromSeconds(10)), Times.Once);
            _recentLinksCleaner.Verify(c => c.DeleteForTarget(It.IsAny<string>(), It.IsAny<TimeSpan>()), Times.Never);
            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.RemoveRecentFile, @"C:\test.txt", 10), Times.Never);
        }

        [TestMethod]
        public void RemoveItem_WithRefreshExplorer_RefreshesAfterSuccessfulRemove()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { @"C:\test.txt" });
            _nativeMutation.Setup(m => m.RemoveRecentFile(@"C:\test.txt", TimeSpan.FromSeconds(10)));
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)));

            _manager.RemoveItem(
                @"C:\test.txt",
                QuickAccess.RecentFiles,
                new RemoveOptions { RefreshExplorer = true });

            _nativeMutation.Verify(m => m.RemoveRecentFile(@"C:\test.txt", TimeSpan.FromSeconds(10)), Times.Once);
            _explorerRefresher.Verify(r => r.Refresh(TimeSpan.FromSeconds(10)), Times.Once);
            _executor.Verify(e => e.ClearCache(), Times.Once);
        }

        [TestMethod]
        public void RemoveItem_RefreshExplorerFailure_ThrowsPostMutationExceptionBeforeDeepClean()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { @"C:\test.txt" });
            _nativeMutation.Setup(m => m.RemoveRecentFile(@"C:\test.txt", TimeSpan.FromSeconds(10)));
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)))
                .Throws(new InvalidOperationException("native refresh failed"));
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, 10))
                .Throws(CreatePowerShellException(PowerShellOperation.RefreshExplorer, PowerShellErrorKind.ProcessFailed, "refresh failed"));

            var ex = Assert.ThrowsException<QuickAccessPostMutationException>(
                () => _manager.RemoveItem(
                    @"C:\test.txt",
                    QuickAccess.RecentFiles,
                    new RemoveOptions { RefreshExplorer = true, DeepCleanRecentLinks = true }));

            Assert.AreEqual(@"C:\test.txt", ex.Path);
            Assert.AreEqual(QuickAccess.RecentFiles, ex.Target);
            Assert.AreEqual(QuickAccessPostMutationStep.RefreshExplorer, ex.Step);
            _recentLinksCleaner.Verify(c => c.DeleteForTarget(It.IsAny<string>(), It.IsAny<TimeSpan>()), Times.Never);
            _executor.Verify(e => e.ClearCache(), Times.Once);
        }

        [TestMethod]
        public void RemoveItem_RecentFile_NativeFailureFallsBackToPowerShell()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { @"C:\test.txt" });
            _nativeMutation.Setup(m => m.RemoveRecentFile(@"C:\test.txt", TimeSpan.FromSeconds(10)))
                .Throws(new COMException("Shell verb failed.", unchecked((int)0x80004005)));

            _manager.RemoveItem(@"C:\test.txt", QuickAccess.RecentFiles);

            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.RemoveRecentFile, @"C:\test.txt", 10), Times.Once);
        }

        [TestMethod]
        public void RemoveItem_RecentFile_NonRecoverableNativeFailureDoesNotFallBack()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { @"C:\test.txt" });
            _nativeMutation.Setup(m => m.RemoveRecentFile(@"C:\test.txt", TimeSpan.FromSeconds(10)))
                .Throws(new InvalidOperationException("native failed"));

            var ex = Assert.ThrowsException<InvalidOperationException>(
                () => _manager.RemoveItem(@"C:\test.txt", QuickAccess.RecentFiles));

            Assert.AreEqual("native failed", ex.Message);
            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.RemoveRecentFile, @"C:\test.txt", 10), Times.Never);
            _executor.Verify(e => e.ClearCache(), Times.Never);
        }

        [TestMethod]
        public void RemoveItem_RecentFile_NativeNotFoundDoesNotFallBack()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { @"C:\test.txt" });
            _nativeMutation.Setup(m => m.RemoveRecentFile(@"C:\test.txt", TimeSpan.FromSeconds(10)))
                .Throws(new QuickAccessItemNotFoundException(@"C:\test.txt", QuickAccess.RecentFiles));

            Assert.ThrowsException<QuickAccessItemNotFoundException>(
                () => _manager.RemoveItem(@"C:\test.txt", QuickAccess.RecentFiles));

            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.RemoveRecentFile, @"C:\test.txt", 10), Times.Never);
        }

        [TestMethod]
        public void RemoveItem_DeepCleanRecentLinks_CleansRecentLinks()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { @"C:\test.txt" });
            _nativeMutation.Setup(m => m.RemoveRecentFile(@"C:\test.txt", TimeSpan.FromSeconds(10)));
            _recentLinksCleaner.Setup(c => c.DeleteForTarget(@"C:\test.txt", TimeSpan.FromSeconds(10)))
                .Returns(new List<string> { @"C:\Users\Test\Recent\test.lnk" });

            _manager.RemoveItem(
                @"C:\test.txt",
                QuickAccess.RecentFiles,
                new RemoveOptions { DeepCleanRecentLinks = true });

            _nativeMutation.Verify(m => m.RemoveRecentFile(@"C:\test.txt", TimeSpan.FromSeconds(10)), Times.Once);
            _recentLinksCleaner.Verify(c => c.DeleteForTarget(@"C:\test.txt", TimeSpan.FromSeconds(10)), Times.Once);
            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.RemoveRecentFile, @"C:\test.txt", 10), Times.Never);
            _dataFiles.Verify(d => d.RemoveRecentFile(), Times.Never);
        }

        [TestMethod]
        public void RemoveItem_DeepCleanRecentLinks_CleanupFailurePropagatesAndClearsCache()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { @"C:\test.txt" });
            _nativeMutation.Setup(m => m.RemoveRecentFile(@"C:\test.txt", TimeSpan.FromSeconds(10)));
            _recentLinksCleaner.Setup(c => c.DeleteForTarget(@"C:\test.txt", TimeSpan.FromSeconds(10)))
                .Throws(new IOException("delete failed"));

            Assert.ThrowsException<IOException>(
                () => _manager.RemoveItem(
                    @"C:\test.txt",
                    QuickAccess.RecentFiles,
                    new RemoveOptions { DeepCleanRecentLinks = true }));

            _executor.Verify(e => e.ClearCache(), Times.Once);
        }

        [TestMethod]
        [ExpectedException(typeof(PowerShellExecutionException))]
        public void RemoveItem_PowerShellFailure_ThrowsPowerShellExecutionException()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string> { @"C:\Folder" });
            _nativeMutation.Setup(m => m.UnpinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)))
                .Throws(new COMException("Shell verb failed.", unchecked((int)0x80004005)));
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.UnpinFromFrequentFolder, @"C:\Folder", 10))
                .Throws(CreatePowerShellException(PowerShellOperation.UnpinFrequentFolder, PowerShellErrorKind.ProcessFailed, "err"));

            _manager.RemoveItem(@"C:\Folder", QuickAccess.FrequentFolders);
        }

        [TestMethod]
        public void RemoveItem_FrequentFolder_PowerShellNotFoundSentinel_ThrowsNotFound()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string> { @"C:\Folder" });
            _nativeMutation.Setup(m => m.UnpinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)))
                .Throws(new COMException("Shell verb failed.", unchecked((int)0x80004005)));
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.UnpinFromFrequentFolder, @"C:\Folder", 10))
                .Throws(new QuickAccessItemNotFoundException(@"C:\Folder", QuickAccess.FrequentFolders));

            var ex = Assert.ThrowsException<QuickAccessItemNotFoundException>(
                () => _manager.RemoveItem(@"C:\Folder", QuickAccess.FrequentFolders));

            Assert.AreEqual(@"C:\Folder", ex.Path);
            Assert.AreEqual(QuickAccess.FrequentFolders, ex.Target);
            _executor.Verify(e => e.ClearCache(), Times.Never);
        }

        [TestMethod]
        public void RemoveItem_FrequentFolder_UsesNativeUnpin()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string> { @"C:\Folder" });
            _nativeMutation.Setup(m => m.UnpinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)));

            _manager.RemoveItem(@"C:\Folder", QuickAccess.FrequentFolders);

            _nativeMutation.Verify(m => m.UnpinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)), Times.Once);
            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.UnpinFromFrequentFolder, @"C:\Folder", 10), Times.Never);
        }

        [TestMethod]
        public void RemoveItem_FrequentFolderWithDeepClean_CleansRecentLinks()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string> { @"C:\Folder" });
            _nativeMutation.Setup(m => m.UnpinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)));
            _recentLinksCleaner.Setup(c => c.DeleteForTarget(@"C:\Folder", TimeSpan.FromSeconds(10)))
                .Returns(new List<string> { @"C:\Users\Test\Recent\Folder.lnk" });

            _manager.RemoveItem(
                @"C:\Folder",
                QuickAccess.FrequentFolders,
                new RemoveOptions { DeepCleanRecentLinks = true });

            _recentLinksCleaner.Verify(c => c.DeleteForTarget(@"C:\Folder", TimeSpan.FromSeconds(10)), Times.Once);
        }

        [TestMethod]
        public void RemoveItem_FrequentFolder_NativeNotFoundDoesNotFallBack()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string> { @"C:\Folder" });
            _nativeMutation.Setup(m => m.UnpinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)))
                .Throws(new QuickAccessItemNotFoundException(@"C:\Folder", QuickAccess.FrequentFolders));

            Assert.ThrowsException<QuickAccessItemNotFoundException>(
                () => _manager.RemoveItem(@"C:\Folder", QuickAccess.FrequentFolders));

            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.UnpinFromFrequentFolder, @"C:\Folder", 10), Times.Never);
        }

        [TestMethod]
        public void RemoveItem_TransientPowerShellFailure_RetriesAndThenSucceeds()
        {
            _manager.Dispose();
            _manager = new QuickAccessManager(
                _executor.Object,
                TimeSpan.FromSeconds(10),
                _fileSystem.Object,
                _nativeMethods.Object,
                _dataFiles.Object,
                new RetryPolicy(1, TimeSpan.FromMilliseconds(1), TimeSpan.FromMilliseconds(1), 1.0, false),
                new PowerShellFallbackNativeQuery(),
                _nativeMutation.Object);

            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string> { @"C:\Folder" });
            _nativeMutation.Setup(m => m.UnpinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)))
                .Throws(new COMException("Shell verb failed.", unchecked((int)0x80004005)));
            _executor.SetupSequence(e => e.ExecutePSScriptWithTimeout(PSScript.UnpinFromFrequentFolder, @"C:\Folder", 10))
                .Throws(CreatePowerShellException(PowerShellOperation.UnpinFrequentFolder, PowerShellErrorKind.Timeout, "operation timed out"))
                .ReturnsAsync(new ScriptResult(0, string.Empty, string.Empty));

            _manager.RemoveItem(@"C:\Folder", QuickAccess.FrequentFolders);

            _nativeMutation.Verify(m => m.UnpinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)), Times.Exactly(2));
            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.UnpinFromFrequentFolder, @"C:\Folder", 10), Times.Exactly(2));
        }

        [TestMethod]
        public void RemoveItem_PermanentPowerShellFailure_DoesNotRetry()
        {
            _manager.Dispose();
            _manager = new QuickAccessManager(
                _executor.Object,
                TimeSpan.FromSeconds(10),
                _fileSystem.Object,
                _nativeMethods.Object,
                _dataFiles.Object,
                new RetryPolicy(1, TimeSpan.FromMilliseconds(1), TimeSpan.FromMilliseconds(1), 1.0, false));

            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string> { @"C:\Folder" });
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.UnpinFromFrequentFolder, @"C:\Folder", 10))
                .Throws(CreatePowerShellException(PowerShellOperation.UnpinFrequentFolder, PowerShellErrorKind.AccessDenied, "Access is denied"));

            Assert.ThrowsException<PowerShellExecutionException>(
                () => _manager.RemoveItem(@"C:\Folder", QuickAccess.FrequentFolders));

            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.UnpinFromFrequentFolder, @"C:\Folder", 10), Times.Once);
        }

        [TestMethod]
        public void RemoveItem_RetryPolicyNone_DoesNotRetryTransientFailure()
        {
            _manager.Dispose();
            _manager = new QuickAccessManager(
                _executor.Object,
                TimeSpan.FromSeconds(10),
                _fileSystem.Object,
                _nativeMethods.Object,
                _dataFiles.Object,
                RetryPolicy.None);

            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string> { @"C:\Folder" });
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.UnpinFromFrequentFolder, @"C:\Folder", 10))
                .Throws(CreatePowerShellException(PowerShellOperation.UnpinFrequentFolder, PowerShellErrorKind.Timeout, "operation timed out"));

            Assert.ThrowsException<PowerShellExecutionException>(
                () => _manager.RemoveItem(@"C:\Folder", QuickAccess.FrequentFolders));

            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.UnpinFromFrequentFolder, @"C:\Folder", 10), Times.Once);
        }

        [TestMethod]
        public void ClearItems_RecentFiles_UsesNativeClear()
        {
            ApartmentState? recentDocsApartment = null;
            _nativeMethods.Setup(n => n.SHAddToRecentDocs(It.IsAny<uint>(), IntPtr.Zero))
                .Callback<uint, IntPtr>((flags, path) => recentDocsApartment = Thread.CurrentThread.GetApartmentState());

            _manager.ClearItems(QuickAccess.RecentFiles);

            Assert.AreEqual(ApartmentState.STA, recentDocsApartment.Value);
            _nativeMethods.Verify(n => n.CoInitializeEx(
                IntPtr.Zero,
                (uint)(NativeMethods.COINIT_APARTMENTTHREADED | NativeMethods.COINIT_DISABLE_OLE1DDE)),
                Times.Once);
            _nativeMethods.Verify(n => n.SHAddToRecentDocs(It.IsAny<uint>(), IntPtr.Zero), Times.Once);
            _executor.Verify(e => e.ClearCache(), Times.Once);
        }

        [TestMethod]
        public void ClearItems_InvalidTarget_ThrowsBeforeDoingWork()
        {
            Assert.ThrowsException<ArgumentOutOfRangeException>(
                () => _manager.ClearItems((QuickAccess)999));

            _nativeMethods.Verify(n => n.SHAddToRecentDocs(It.IsAny<uint>(), It.IsAny<IntPtr>()), Times.Never);
            _executor.Verify(e => e.ClearCache(), Times.Never);
        }

        [TestMethod]
        [ExpectedException(typeof(ComApartmentMismatchException))]
        public void ClearItems_ComInitializationFailure_ThrowsComApartmentMismatchException()
        {
            _nativeMethods.Setup(n => n.CoInitializeEx(It.IsAny<IntPtr>(), It.IsAny<uint>()))
                .Returns(unchecked((int)0x80010106));

            _manager.ClearItems(QuickAccess.RecentFiles);
        }

        [TestMethod]
        public void ClearItems_PositiveComInformationCode_DoesNotThrowComException()
        {
            // Any non-negative HRESULT represents success or informational status for this guard.
            _nativeMethods.Setup(n => n.CoInitializeEx(It.IsAny<IntPtr>(), It.IsAny<uint>()))
                .Returns(2);

            _manager.ClearItems(QuickAccess.RecentFiles);

            _nativeMethods.Verify(n => n.SHAddToRecentDocs(It.IsAny<uint>(), IntPtr.Zero), Times.Once);
        }

        [TestMethod]
        public void ClearItems_FrequentFolders_DeletesBackingFile()
        {
            _manager.ClearItems(QuickAccess.FrequentFolders);

            _fileSystem.Verify(f => f.DeleteFile(@"C:\frequent.automaticDestinations-ms"), Times.Once);
        }

        [TestMethod]
        public void ClearItems_FrequentFoldersPinnedCleanup_UsesPinnedFoldersTimeout()
        {
            var nativeQuery = new Mock<IQuickAccessNativeQuery>(MockBehavior.Strict);
            nativeQuery.Setup(q => q.GetItems(QuickAccess.FrequentFolders, TimeSpan.FromSeconds(3)))
                .Returns(new List<string> { @"C:\Pinned" });
            _nativeMutation.Setup(m => m.UnpinFrequentFolder(@"C:\Pinned", TimeSpan.FromSeconds(3)));
            _manager.Dispose();
            _manager = CreateManager(nativeQuery.Object, _nativeMutation.Object);

            _manager.ClearItems(
                QuickAccess.FrequentFolders,
                new ClearOptions
                {
                    RemovePinnedFolders = true,
                    PinnedFoldersTimeout = TimeSpan.FromSeconds(3)
                });

            nativeQuery.Verify(q => q.GetItems(QuickAccess.FrequentFolders, TimeSpan.FromSeconds(3)), Times.Once);
            _nativeMutation.Verify(m => m.UnpinFrequentFolder(@"C:\Pinned", TimeSpan.FromSeconds(3)), Times.Once);
            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.EmptyPinnedFolders, null, It.IsAny<int>()), Times.Never);
        }

        [TestMethod]
        public void ClearItems_FrequentFoldersPinnedCleanup_DefaultsToManagerTimeout()
        {
            var nativeQuery = new Mock<IQuickAccessNativeQuery>(MockBehavior.Strict);
            nativeQuery.Setup(q => q.GetItems(QuickAccess.FrequentFolders, TimeSpan.FromSeconds(10)))
                .Returns(new List<string> { @"C:\Pinned" });
            _nativeMutation.Setup(m => m.UnpinFrequentFolder(@"C:\Pinned", TimeSpan.FromSeconds(10)));
            _manager.Dispose();
            _manager = CreateManager(nativeQuery.Object, _nativeMutation.Object);

            _manager.ClearItems(
                QuickAccess.FrequentFolders,
                new ClearOptions { RemovePinnedFolders = true });

            nativeQuery.Verify(q => q.GetItems(QuickAccess.FrequentFolders, TimeSpan.FromSeconds(10)), Times.Once);
            _nativeMutation.Verify(m => m.UnpinFrequentFolder(@"C:\Pinned", TimeSpan.FromSeconds(10)), Times.Once);
        }

        [TestMethod]
        public void ClearItems_PinnedFoldersTimeoutZeroWithPinnedCleanup_ThrowsBeforeDoingWork()
        {
            Assert.ThrowsException<ArgumentOutOfRangeException>(
                () => _manager.ClearItems(
                    QuickAccess.FrequentFolders,
                    new ClearOptions
                    {
                        RemovePinnedFolders = true,
                        PinnedFoldersTimeout = TimeSpan.Zero
                    }));

            _fileSystem.Verify(f => f.DeleteFile(It.IsAny<string>()), Times.Never);
            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.EmptyPinnedFolders, null, It.IsAny<int>()), Times.Never);
        }

        [TestMethod]
        public void ClearItems_PinnedFoldersTimeoutZeroWithoutPinnedCleanup_DoesNotThrow()
        {
            _manager.ClearItems(
                QuickAccess.FrequentFolders,
                new ClearOptions { PinnedFoldersTimeout = TimeSpan.Zero });

            _fileSystem.Verify(f => f.DeleteFile(@"C:\frequent.automaticDestinations-ms"), Times.Once);
            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.EmptyPinnedFolders, null, It.IsAny<int>()), Times.Never);
        }

        [TestMethod]
        public void ClearItems_WithRefresh_UsesNativeExplorerRefresh()
        {
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)));

            _manager.ClearItems(QuickAccess.RecentFiles, new ClearOptions { RefreshExplorer = true });

            _explorerRefresher.Verify(r => r.Refresh(TimeSpan.FromSeconds(10)), Times.Once);
            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, 10), Times.Never);
        }

        [TestMethod]
        public void ClearItems_WithRefresh_NativeRefreshFailureFallsBackToPowerShell()
        {
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)))
                .Throws(new InvalidOperationException("native refresh failed"));

            _manager.ClearItems(QuickAccess.RecentFiles, new ClearOptions { RefreshExplorer = true });

            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, 10), Times.Once);
        }

        [TestMethod]
        public void ClearItems_WithRefresh_AllRefreshPathsFail_ThrowsPowerShellException()
        {
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)))
                .Throws(new InvalidOperationException("native refresh failed"));
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, 10))
                .Throws(CreatePowerShellException(PowerShellOperation.RefreshExplorer, PowerShellErrorKind.ProcessFailed, "refresh failed"));

            Assert.ThrowsException<PowerShellExecutionException>(
                () => _manager.ClearItems(QuickAccess.RecentFiles, new ClearOptions { RefreshExplorer = true }));
        }

        [TestMethod]
        public void ClearItems_AllPartialFailure_ThrowsPartialClearException()
        {
            var source = new Win32Exception(1);
            _nativeMethods.Setup(n => n.CoInitializeEx(It.IsAny<IntPtr>(), It.IsAny<uint>()))
                .Throws(source);

            var ex = Assert.ThrowsException<PartialClearException>(() =>
                _manager.ClearItems(QuickAccess.All));

            Assert.IsFalse(ex.RecentFilesCleared);
            Assert.IsTrue(ex.FrequentFoldersCleared);
            Assert.IsTrue(ex.HasPartialProgress);
            Assert.IsFalse(ex.IsCompleteFailure);
            Assert.AreSame(ex.InnerException, ex.SourceException);
            Assert.AreSame(source, ex.SourceException);
            StringAssert.Contains(ex.Message, "recent_files_cleared: false");
            StringAssert.Contains(ex.Message, "frequent_folders_cleared: true");
        }

        [TestMethod]
        public void ClearItems_AllCompleteFailure_ReportsNoProgress()
        {
            var source = new Win32Exception(1);
            _nativeMethods.Setup(n => n.CoInitializeEx(It.IsAny<IntPtr>(), It.IsAny<uint>()))
                .Throws(source);
            _fileSystem.Setup(f => f.FileExists(@"C:\frequent.automaticDestinations-ms")).Returns(true);
            _fileSystem.Setup(f => f.DeleteFile(@"C:\frequent.automaticDestinations-ms"))
                .Throws(new IOException("delete failed"));

            var ex = Assert.ThrowsException<PartialClearException>(() =>
                _manager.ClearItems(QuickAccess.All));

            Assert.IsFalse(ex.RecentFilesCleared);
            Assert.IsFalse(ex.FrequentFoldersCleared);
            Assert.IsFalse(ex.HasPartialProgress);
            Assert.IsTrue(ex.IsCompleteFailure);
            Assert.AreSame(source, ex.SourceException);
        }

        [TestMethod]
        public void ClearItems_FrequentFoldersPinnedCleanupFailure_ReportsFrequentProgress()
        {
            var nativeQuery = new Mock<IQuickAccessNativeQuery>(MockBehavior.Strict);
            nativeQuery.Setup(q => q.GetItems(QuickAccess.FrequentFolders, TimeSpan.FromSeconds(10)))
                .Returns(new List<string> { @"C:\Pinned" });
            _nativeMutation.Setup(m => m.UnpinFrequentFolder(@"C:\Pinned", TimeSpan.FromSeconds(10)))
                .Throws(new InvalidOperationException("unpin failed"));
            _manager.Dispose();
            _manager = CreateManager(nativeQuery.Object, _nativeMutation.Object);

            var ex = Assert.ThrowsException<PartialClearException>(() =>
                _manager.ClearItems(QuickAccess.FrequentFolders, new ClearOptions { RemovePinnedFolders = true }));

            Assert.IsFalse(ex.RecentFilesCleared);
            Assert.IsTrue(ex.FrequentFoldersCleared);
            Assert.IsTrue(ex.HasPartialProgress);
            Assert.IsFalse(ex.IsCompleteFailure);
            Assert.AreSame(ex.InnerException, ex.SourceException);
            Assert.IsInstanceOfType(ex.SourceException, typeof(InvalidOperationException));
        }

        [TestMethod]
        public void ClearItems_AllFrequentPartialFailure_MergesProgressFlags()
        {
            var nativeQuery = new Mock<IQuickAccessNativeQuery>(MockBehavior.Strict);
            nativeQuery.Setup(q => q.GetItems(QuickAccess.FrequentFolders, TimeSpan.FromSeconds(10)))
                .Returns(new List<string> { @"C:\Pinned" });
            _nativeMutation.Setup(m => m.UnpinFrequentFolder(@"C:\Pinned", TimeSpan.FromSeconds(10)))
                .Throws(new InvalidOperationException("unpin failed"));
            _manager.Dispose();
            _manager = CreateManager(nativeQuery.Object, _nativeMutation.Object);

            var ex = Assert.ThrowsException<PartialClearException>(() =>
                _manager.ClearItems(QuickAccess.All, new ClearOptions { RemovePinnedFolders = true }));

            Assert.IsTrue(ex.RecentFilesCleared);
            Assert.IsTrue(ex.FrequentFoldersCleared);
            Assert.IsTrue(ex.HasPartialProgress);
            Assert.IsFalse(ex.IsCompleteFailure);
            Assert.AreSame(ex.InnerException, ex.SourceException);
        }

        [TestMethod]
        public void ClearItems_FrequentFoldersPinnedCleanup_IgnoresMissingPinnedFoldersAndContinues()
        {
            var nativeQuery = new Mock<IQuickAccessNativeQuery>(MockBehavior.Strict);
            nativeQuery.Setup(q => q.GetItems(QuickAccess.FrequentFolders, TimeSpan.FromSeconds(10)))
                .Returns(new List<string> { @"C:\Missing", @"C:\Pinned" });
            _nativeMutation.Setup(m => m.UnpinFrequentFolder(@"C:\Missing", TimeSpan.FromSeconds(10)))
                .Throws(new QuickAccessItemNotFoundException(@"C:\Missing", QuickAccess.FrequentFolders));
            _nativeMutation.Setup(m => m.UnpinFrequentFolder(@"C:\Pinned", TimeSpan.FromSeconds(10)));
            _manager.Dispose();
            _manager = CreateManager(nativeQuery.Object, _nativeMutation.Object);

            _manager.ClearItems(QuickAccess.FrequentFolders, new ClearOptions { RemovePinnedFolders = true });

            _nativeMutation.Verify(m => m.UnpinFrequentFolder(@"C:\Missing", TimeSpan.FromSeconds(10)), Times.Once);
            _nativeMutation.Verify(m => m.UnpinFrequentFolder(@"C:\Pinned", TimeSpan.FromSeconds(10)), Times.Once);
        }

        [TestMethod]
        public void ClearItems_PartialFailureWithRefresh_PreservesPartialClearException()
        {
            _nativeMethods.Setup(n => n.CoInitializeEx(It.IsAny<IntPtr>(), It.IsAny<uint>()))
                .Throws(new Win32Exception(1));
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)))
                .Throws(new InvalidOperationException("native refresh failed"));
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, 10))
                .Throws(CreatePowerShellException(PowerShellOperation.RefreshExplorer, PowerShellErrorKind.ProcessFailed, "refresh failed"));

            var ex = Assert.ThrowsException<PartialClearException>(() =>
                _manager.ClearItems(QuickAccess.All, new ClearOptions { RefreshExplorer = true }));

            Assert.IsFalse(ex.RecentFilesCleared);
            Assert.IsTrue(ex.FrequentFoldersCleared);
            Assert.IsTrue(ex.HasPartialProgress);
            Assert.IsFalse(ex.IsCompleteFailure);
            Assert.AreSame(ex.InnerException, ex.SourceException);
        }

        [TestMethod]
        public void AddItems_RecordsPerItemFailures()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { @"C:\exists.txt" });
            _nativeMutation.Setup(m => m.PinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)));

            var result = _manager.AddItems(new[]
            {
                QuickAccessItem.RecentFile(@"C:\exists.txt"),
                QuickAccessItem.FrequentFolder(@"C:\Folder")
            });

            Assert.AreEqual(2, result.Total);
            Assert.AreEqual(1, result.Succeeded.Count);
            Assert.AreEqual(1, result.Failed.Count);
            Assert.IsTrue(result.HasPartialSuccess);
            _nativeMutation.Verify(m => m.PinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)), Times.Once);
        }

        [TestMethod]
        public void AddItems_CoalescesRecentRefreshOnceAfterSuccessfulRecentAdds()
        {
            var result = _manager.AddItems(
                new[]
                {
                    QuickAccessItem.RecentFile(@"C:\one.txt"),
                    QuickAccessItem.RecentFile(@"C:\two.txt")
                },
                new BatchOptions { RefreshRecentFiles = true });

            Assert.AreEqual(2, result.Total);
            Assert.AreEqual(2, result.Succeeded.Count);
            Assert.AreEqual(0, result.Failed.Count);
            _dataFiles.Verify(d => d.RemoveRecentFile(), Times.Once);
        }

        [TestMethod]
        public void AddItems_DoesNotRefreshWhenAllRecentAddsFail()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { @"C:\one.txt", @"C:\two.txt" });

            var result = _manager.AddItems(
                new[]
                {
                    QuickAccessItem.RecentFile(@"C:\one.txt"),
                    QuickAccessItem.RecentFile(@"C:\two.txt")
                },
                new BatchOptions { RefreshRecentFiles = true });

            Assert.AreEqual(2, result.Total);
            Assert.AreEqual(0, result.Succeeded.Count);
            Assert.AreEqual(2, result.Failed.Count);
            _dataFiles.Verify(d => d.RemoveRecentFile(), Times.Never);
        }

        [TestMethod]
        public void AddItems_DoesNotRefreshWhenOnlyFrequentFoldersSucceed()
        {
            _nativeMutation.Setup(m => m.PinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)));

            var result = _manager.AddItems(
                new[] { QuickAccessItem.FrequentFolder(@"C:\Folder") },
                new BatchOptions { RefreshRecentFiles = true });

            Assert.AreEqual(1, result.Total);
            Assert.AreEqual(1, result.Succeeded.Count);
            Assert.AreEqual(0, result.Failed.Count);
            _dataFiles.Verify(d => d.RemoveRecentFile(), Times.Never);
        }

        [TestMethod]
        public void AddItems_RecentBackingRefreshFailureRecordsPostMutationFailure()
        {
            var recent = QuickAccessItem.RecentFile(@"C:\recent.txt");
            var folder = QuickAccessItem.FrequentFolder(@"C:\Folder");
            var refreshError = new IOException("refresh failed");
            _nativeMutation.Setup(m => m.PinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)));
            _dataFiles.Setup(d => d.RemoveRecentFile()).Throws(refreshError);

            var result = _manager.AddItems(
                new[] { recent, folder },
                new BatchOptions { RefreshRecentFiles = true });

            Assert.AreEqual(2, result.Total);
            Assert.AreEqual(1, result.Succeeded.Count);
            Assert.AreSame(folder, result.Succeeded[0]);
            Assert.AreEqual(1, result.Failed.Count);
            Assert.AreSame(recent, result.Failed[0].Item);
            var error = (QuickAccessPostMutationException)result.Failed[0].Error;
            Assert.AreEqual(recent.Path, error.Path);
            Assert.AreEqual(QuickAccess.RecentFiles, error.Target);
            Assert.AreEqual(QuickAccessPostMutationStep.DeleteRecentFilesBackingData, error.Step);
            Assert.AreSame(refreshError, error.InnerException);
        }

        [TestMethod]
        public void AddItems_WithRefreshExplorer_RefreshesOnceAfterSuccessfulAdds()
        {
            var recent = QuickAccessItem.RecentFile(@"C:\recent.txt");
            var folder = QuickAccessItem.FrequentFolder(@"C:\Folder");
            _nativeMutation.Setup(m => m.PinFrequentFolder(@"C:\Folder", TimeSpan.FromSeconds(10)));
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)));

            var result = _manager.AddItems(
                new[] { folder, recent },
                new BatchOptions { RefreshExplorer = true });

            Assert.AreEqual(2, result.Total);
            Assert.AreEqual(2, result.Succeeded.Count);
            Assert.AreEqual(0, result.Failed.Count);
            _explorerRefresher.Verify(r => r.Refresh(TimeSpan.FromSeconds(10)), Times.Once);
        }

        [TestMethod]
        public void AddItems_RefreshExplorerFailureRecordsPostMutationFailure()
        {
            var folderA = QuickAccessItem.FrequentFolder(@"C:\FolderA");
            var folderB = QuickAccessItem.FrequentFolder(@"C:\FolderB");
            _nativeMutation.Setup(m => m.PinFrequentFolder(folderA.Path, TimeSpan.FromSeconds(10)));
            _nativeMutation.Setup(m => m.PinFrequentFolder(folderB.Path, TimeSpan.FromSeconds(10)));
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)))
                .Throws(new InvalidOperationException("native refresh failed"));
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, 10))
                .Throws(CreatePowerShellException(PowerShellOperation.RefreshExplorer, PowerShellErrorKind.ProcessFailed, "refresh failed"));

            var result = _manager.AddItems(
                new[] { folderA, folderB },
                new BatchOptions { RefreshExplorer = true });

            Assert.AreEqual(2, result.Total);
            Assert.AreEqual(1, result.Succeeded.Count);
            Assert.AreSame(folderB, result.Succeeded[0]);
            Assert.AreEqual(1, result.Failed.Count);
            Assert.AreSame(folderA, result.Failed[0].Item);
            var error = (QuickAccessPostMutationException)result.Failed[0].Error;
            Assert.AreEqual(folderA.Path, error.Path);
            Assert.AreEqual(QuickAccess.FrequentFolders, error.Target);
            Assert.AreEqual(QuickAccessPostMutationStep.RefreshExplorer, error.Step);
        }

        [TestMethod]
        public void RemoveItems_DeepCleanFailureRecordsPerItemFailureAndContinues()
        {
            var failing = QuickAccessItem.RecentFile(@"C:\failing.txt");
            var succeeding = QuickAccessItem.RecentFile(@"C:\succeeding.txt");
            var cleanupError = new IOException("delete failed");
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { failing.Path, succeeding.Path });
            _nativeMutation.Setup(m => m.RemoveRecentFile(failing.Path, TimeSpan.FromSeconds(10)));
            _nativeMutation.Setup(m => m.RemoveRecentFile(succeeding.Path, TimeSpan.FromSeconds(10)));
            _recentLinksCleaner.Setup(c => c.DeleteForTarget(failing.Path, TimeSpan.FromSeconds(10)))
                .Throws(cleanupError);
            _recentLinksCleaner.Setup(c => c.DeleteForTarget(succeeding.Path, TimeSpan.FromSeconds(10)))
                .Returns(new List<string>());

            var result = _manager.RemoveItems(
                new[] { failing, succeeding },
                new RemoveOptions { DeepCleanRecentLinks = true });

            Assert.AreEqual(2, result.Total);
            Assert.AreEqual(1, result.Succeeded.Count);
            Assert.AreSame(succeeding, result.Succeeded[0]);
            Assert.AreEqual(1, result.Failed.Count);
            Assert.AreSame(failing, result.Failed[0].Item);
            Assert.AreSame(cleanupError, result.Failed[0].Error);
            _nativeMutation.Verify(m => m.RemoveRecentFile(succeeding.Path, TimeSpan.FromSeconds(10)), Times.Once);
            _recentLinksCleaner.Verify(c => c.DeleteForTarget(succeeding.Path, TimeSpan.FromSeconds(10)), Times.Once);
        }

        [TestMethod]
        public void RemoveItems_DeepCleanFailureAfterMutation_StillRefreshesExplorer()
        {
            var item = QuickAccessItem.RecentFile(@"C:\failing.txt");
            var cleanupError = new IOException("delete failed");
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { item.Path });
            _nativeMutation.Setup(m => m.RemoveRecentFile(item.Path, TimeSpan.FromSeconds(10)));
            _recentLinksCleaner.Setup(c => c.DeleteForTarget(item.Path, TimeSpan.FromSeconds(10)))
                .Throws(cleanupError);

            var result = _manager.RemoveItems(
                new[] { item },
                new BatchOptions { RefreshExplorer = true },
                new RemoveOptions { DeepCleanRecentLinks = true });

            Assert.AreEqual(1, result.Total);
            Assert.AreEqual(0, result.Succeeded.Count);
            Assert.AreEqual(1, result.Failed.Count);
            Assert.AreSame(item, result.Failed[0].Item);
            Assert.AreSame(cleanupError, result.Failed[0].Error);
            _explorerRefresher.Verify(r => r.Refresh(TimeSpan.FromSeconds(10)), Times.Once);
            _executor.Verify(e => e.ClearCache(), Times.Once);
        }

        [TestMethod]
        public void RemoveItems_DeepCleanFailureAndRefreshFailure_DoesNotRecordDuplicateFailure()
        {
            var item = QuickAccessItem.RecentFile(@"C:\failing.txt");
            var cleanupError = new IOException("delete failed");
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { item.Path });
            _nativeMutation.Setup(m => m.RemoveRecentFile(item.Path, TimeSpan.FromSeconds(10)));
            _recentLinksCleaner.Setup(c => c.DeleteForTarget(item.Path, TimeSpan.FromSeconds(10)))
                .Throws(cleanupError);
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)))
                .Throws(new InvalidOperationException("native refresh failed"));
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, 10))
                .Throws(CreatePowerShellException(PowerShellOperation.RefreshExplorer, PowerShellErrorKind.ProcessFailed, "refresh failed"));

            var result = _manager.RemoveItems(
                new[] { item },
                new BatchOptions { RefreshExplorer = true },
                new RemoveOptions { DeepCleanRecentLinks = true });

            Assert.AreEqual(1, result.Total);
            Assert.AreEqual(0, result.Succeeded.Count);
            Assert.AreEqual(1, result.Failed.Count);
            Assert.AreSame(item, result.Failed[0].Item);
            Assert.AreSame(cleanupError, result.Failed[0].Error);
            _explorerRefresher.Verify(r => r.Refresh(TimeSpan.FromSeconds(10)), Times.Once);
            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, 10), Times.Once);
        }

        [TestMethod]
        public void RemoveItems_DeepCleanDisabledDoesNotCallCleaner()
        {
            var item = QuickAccessItem.RecentFile(@"C:\test.txt");
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { item.Path });
            _nativeMutation.Setup(m => m.RemoveRecentFile(item.Path, TimeSpan.FromSeconds(10)));

            var result = _manager.RemoveItems(new[] { item });

            Assert.AreEqual(1, result.Total);
            Assert.AreEqual(1, result.Succeeded.Count);
            Assert.AreEqual(0, result.Failed.Count);
            _recentLinksCleaner.Verify(c => c.DeleteForTarget(It.IsAny<string>(), It.IsAny<TimeSpan>()), Times.Never);
        }

        [TestMethod]
        public void RemoveItems_FrequentFoldersWithDeepClean_CleansRecentLinks()
        {
            var item = QuickAccessItem.FrequentFolder(@"C:\Folder");
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string> { item.Path });
            _nativeMutation.Setup(m => m.UnpinFrequentFolder(item.Path, TimeSpan.FromSeconds(10)));
            _recentLinksCleaner.Setup(c => c.DeleteForTarget(item.Path, TimeSpan.FromSeconds(10)))
                .Returns(new List<string> { @"C:\Users\Test\Recent\Folder.lnk" });

            var result = _manager.RemoveItems(
                new[] { item },
                new RemoveOptions { DeepCleanRecentLinks = true });

            Assert.AreEqual(1, result.Total);
            Assert.AreEqual(1, result.Succeeded.Count);
            Assert.AreEqual(0, result.Failed.Count);
            _recentLinksCleaner.Verify(c => c.DeleteForTarget(item.Path, TimeSpan.FromSeconds(10)), Times.Once);
        }

        [TestMethod]
        public void RemoveItems_AllTargetRecordsFailureAndContinues()
        {
            var unsupported = new QuickAccessItem(@"C:\unsupported.txt", QuickAccess.All);
            var valid = QuickAccessItem.RecentFile(@"C:\valid.txt");
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { valid.Path });
            _nativeMutation.Setup(m => m.RemoveRecentFile(valid.Path, TimeSpan.FromSeconds(10)));

            var result = _manager.RemoveItems(new[] { unsupported, valid });

            Assert.AreEqual(2, result.Total);
            Assert.AreEqual(1, result.Succeeded.Count);
            Assert.AreSame(valid, result.Succeeded[0]);
            Assert.AreEqual(1, result.Failed.Count);
            Assert.AreSame(unsupported, result.Failed[0].Item);
            Assert.IsInstanceOfType(result.Failed[0].Error, typeof(UnsupportedQuickAccessOperationException));
        }

        [TestMethod]
        public void RemoveItems_WithBatchRefreshExplorer_RefreshesOnceAndIgnoresRemoveOptionRefresh()
        {
            var item = QuickAccessItem.RecentFile(@"C:\test.txt");
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { item.Path });
            _nativeMutation.Setup(m => m.RemoveRecentFile(item.Path, TimeSpan.FromSeconds(10)));
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)));

            var result = _manager.RemoveItems(
                new[] { item },
                new BatchOptions { RefreshExplorer = true },
                new RemoveOptions { RefreshExplorer = true });

            Assert.AreEqual(1, result.Total);
            Assert.AreEqual(1, result.Succeeded.Count);
            Assert.AreEqual(0, result.Failed.Count);
            _explorerRefresher.Verify(r => r.Refresh(TimeSpan.FromSeconds(10)), Times.Once);
        }

        [TestMethod]
        public void RemoveItems_RefreshExplorerFailureRecordsPostMutationFailure()
        {
            var one = QuickAccessItem.RecentFile(@"C:\one.txt");
            var two = QuickAccessItem.RecentFile(@"C:\two.txt");
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { one.Path, two.Path });
            _nativeMutation.Setup(m => m.RemoveRecentFile(one.Path, TimeSpan.FromSeconds(10)));
            _nativeMutation.Setup(m => m.RemoveRecentFile(two.Path, TimeSpan.FromSeconds(10)));
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)))
                .Throws(new InvalidOperationException("native refresh failed"));
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, 10))
                .Throws(CreatePowerShellException(PowerShellOperation.RefreshExplorer, PowerShellErrorKind.ProcessFailed, "refresh failed"));

            var result = _manager.RemoveItems(
                new[] { one, two },
                new BatchOptions { RefreshExplorer = true },
                new RemoveOptions());

            Assert.AreEqual(2, result.Total);
            Assert.AreEqual(1, result.Succeeded.Count);
            Assert.AreSame(one, result.Succeeded[0]);
            Assert.AreEqual(1, result.Failed.Count);
            Assert.AreSame(two, result.Failed[0].Item);
            var error = (QuickAccessPostMutationException)result.Failed[0].Error;
            Assert.AreEqual(two.Path, error.Path);
            Assert.AreEqual(QuickAccess.RecentFiles, error.Target);
            Assert.AreEqual(QuickAccessPostMutationStep.RefreshExplorer, error.Step);
        }

        [TestMethod]
        public void EmptyBatch_IsCompleteSuccess()
        {
            var result = _manager.RemoveItems(Array.Empty<QuickAccessItem>());

            Assert.AreEqual(0, result.Total);
            Assert.IsTrue(result.IsCompleteSuccess);
            Assert.IsFalse(result.HasPartialSuccess);
            Assert.AreEqual(1.0, result.SuccessRate);
        }

        [TestMethod]
        public void BatchOperations_NullArgumentsThrow()
        {
            Assert.ThrowsException<ArgumentNullException>(() => _manager.AddItems(null));
            Assert.ThrowsException<ArgumentNullException>(() => _manager.AddItems(Array.Empty<QuickAccessItem>(), null));
            Assert.ThrowsException<ArgumentNullException>(() => _manager.RemoveItems(null));
            Assert.ThrowsException<ArgumentNullException>(() => _manager.RemoveItems(Array.Empty<QuickAccessItem>(), null));
        }

        [TestMethod]
        public void BatchOperations_NullItemThrowsArgumentException()
        {
            Assert.ThrowsException<ArgumentException>(
                () => _manager.AddItems(new QuickAccessItem[] { null }));
            Assert.ThrowsException<ArgumentException>(
                () => _manager.RemoveItems(new QuickAccessItem[] { null }));
        }

        [TestMethod]
        public void RetryPolicy_None_RejectsGetDelay()
        {
            Assert.ThrowsException<InvalidOperationException>(() => RetryPolicy.None.GetDelay(0));
        }

        [TestMethod]
        public void LockQuickAccess_UsesAllLockTarget()
        {
            var lockFactory = new Mock<IQuickAccessLockFactory>(MockBehavior.Strict);
            var expected = CreateEmptyLock(QuickAccessLockTarget.All);
            lockFactory.Setup(f => f.Lock(QuickAccessLockTarget.All)).Returns(expected);
            _manager.Dispose();
            _manager = CreateManager(lockFactory.Object);

            var result = _manager.LockQuickAccess();

            Assert.AreSame(expected, result);
            lockFactory.Verify(f => f.Lock(QuickAccessLockTarget.All), Times.Once);
        }

        [TestMethod]
        public void LockRecentFiles_UsesRecentFilesLockTarget()
        {
            var lockFactory = new Mock<IQuickAccessLockFactory>(MockBehavior.Strict);
            var expected = CreateEmptyLock(QuickAccessLockTarget.RecentFiles);
            lockFactory.Setup(f => f.Lock(QuickAccessLockTarget.RecentFiles)).Returns(expected);
            _manager.Dispose();
            _manager = CreateManager(lockFactory.Object);

            var result = _manager.LockRecentFiles();

            Assert.AreSame(expected, result);
            lockFactory.Verify(f => f.Lock(QuickAccessLockTarget.RecentFiles), Times.Once);
        }

        [TestMethod]
        public void LockFrequentFolders_UsesFrequentFoldersLockTarget()
        {
            var lockFactory = new Mock<IQuickAccessLockFactory>(MockBehavior.Strict);
            var expected = CreateEmptyLock(QuickAccessLockTarget.FrequentFolders);
            lockFactory.Setup(f => f.Lock(QuickAccessLockTarget.FrequentFolders)).Returns(expected);
            _manager.Dispose();
            _manager = CreateManager(lockFactory.Object);

            var result = _manager.LockFrequentFolders();

            Assert.AreSame(expected, result);
            lockFactory.Verify(f => f.Lock(QuickAccessLockTarget.FrequentFolders), Times.Once);
        }

        [TestMethod]
        public void IsVisible_DelegatesToVisibilityService()
        {
            var visibility = new Mock<IQuickAccessVisibility>(MockBehavior.Strict);
            visibility.Setup(v => v.IsVisible(QuickAccess.RecentFiles)).Returns(false);
            _manager.Dispose();
            _manager = CreateManager(visibility.Object);

            var result = _manager.IsVisible(QuickAccess.RecentFiles);

            Assert.IsFalse(result);
            visibility.Verify(v => v.IsVisible(QuickAccess.RecentFiles), Times.Once);
        }

        [TestMethod]
        public void SetVisible_DelegatesToVisibilityService()
        {
            var visibility = new Mock<IQuickAccessVisibility>(MockBehavior.Strict);
            visibility.Setup(v => v.SetVisible(QuickAccess.FrequentFolders, false));
            _manager.Dispose();
            _manager = CreateManager(visibility.Object);

            _manager.SetVisible(QuickAccess.FrequentFolders, false);

            visibility.Verify(v => v.SetVisible(QuickAccess.FrequentFolders, false), Times.Once);
        }

        [TestMethod]
        public void SetVisible_VisibilityFailure_PropagatesVisibilityExceptionWithoutRefresh()
        {
            var error = new QuickAccessVisibilityException(
                "WriteVisibility",
                QuickAccess.RecentFiles,
                "ShowRecent",
                new UnauthorizedAccessException("write denied"));
            var visibility = new Mock<IQuickAccessVisibility>(MockBehavior.Strict);
            visibility.Setup(v => v.SetVisible(QuickAccess.RecentFiles, false)).Throws(error);
            _manager.Dispose();
            _manager = CreateManager(visibility.Object);

            var ex = Assert.ThrowsException<QuickAccessVisibilityException>(() =>
                _manager.SetVisible(
                    QuickAccess.RecentFiles,
                    false,
                    new VisibilityOptions { RefreshExplorer = true }));

            Assert.AreSame(error, ex);
            _explorerRefresher.Verify(r => r.Refresh(It.IsAny<TimeSpan>()), Times.Never);
            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, It.IsAny<int>()), Times.Never);
        }

        [TestMethod]
        public void SetVisible_DefaultOptions_DoesNotRefreshExplorer()
        {
            var visibility = new Mock<IQuickAccessVisibility>(MockBehavior.Strict);
            visibility.Setup(v => v.SetVisible(QuickAccess.FrequentFolders, false));
            _manager.Dispose();
            _manager = CreateManager(visibility.Object);

            _manager.SetVisible(QuickAccess.FrequentFolders, false);

            visibility.Verify(v => v.SetVisible(QuickAccess.FrequentFolders, false), Times.Once);
            _explorerRefresher.Verify(r => r.Refresh(It.IsAny<TimeSpan>()), Times.Never);
        }

        [TestMethod]
        public void SetVisible_WithRefresh_RefreshesExplorer()
        {
            var visibility = new Mock<IQuickAccessVisibility>(MockBehavior.Strict);
            visibility.Setup(v => v.SetVisible(QuickAccess.RecentFiles, true));
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)));
            _manager.Dispose();
            _manager = CreateManager(visibility.Object);

            _manager.SetVisible(
                QuickAccess.RecentFiles,
                true,
                new VisibilityOptions { RefreshExplorer = true });

            visibility.Verify(v => v.SetVisible(QuickAccess.RecentFiles, true), Times.Once);
            _explorerRefresher.Verify(r => r.Refresh(TimeSpan.FromSeconds(10)), Times.Once);
        }

        [TestMethod]
        public void SetVisible_WithRefresh_NativeRefreshFailureUsesPowerShellFallback()
        {
            var visibility = new Mock<IQuickAccessVisibility>(MockBehavior.Strict);
            visibility.Setup(v => v.SetVisible(QuickAccess.RecentFiles, false));
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)))
                .Throws(new InvalidOperationException("native refresh failed"));
            _manager.Dispose();
            _manager = CreateManager(visibility.Object);

            _manager.SetVisible(
                QuickAccess.RecentFiles,
                false,
                new VisibilityOptions { RefreshExplorer = true });

            visibility.Verify(v => v.SetVisible(QuickAccess.RecentFiles, false), Times.Once);
            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, 10), Times.Once);
        }

        [TestMethod]
        public void SetVisible_WithRefresh_AllRefreshPathsFail_ThrowsPowerShellException()
        {
            var visibility = new Mock<IQuickAccessVisibility>(MockBehavior.Strict);
            visibility.Setup(v => v.SetVisible(QuickAccess.RecentFiles, false));
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)))
                .Throws(new InvalidOperationException("native refresh failed"));
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, 10))
                .Throws(CreatePowerShellException(PowerShellOperation.RefreshExplorer, PowerShellErrorKind.ProcessFailed, "refresh failed"));
            _manager.Dispose();
            _manager = CreateManager(visibility.Object);

            Assert.ThrowsException<PowerShellExecutionException>(() =>
                _manager.SetVisible(
                    QuickAccess.RecentFiles,
                    false,
                    new VisibilityOptions { RefreshExplorer = true }));

            visibility.Verify(v => v.SetVisible(QuickAccess.RecentFiles, false), Times.Once);
        }

        [TestMethod]
        public void SetVisible_NullOptions_ThrowsBeforeWrite()
        {
            var visibility = new Mock<IQuickAccessVisibility>(MockBehavior.Strict);
            _manager.Dispose();
            _manager = CreateManager(visibility.Object);

            Assert.ThrowsException<ArgumentNullException>(() =>
                _manager.SetVisible(QuickAccess.RecentFiles, true, null));
        }

        [TestMethod]
        public void ShowSection_SetsVisibleTrue()
        {
            var visibility = new Mock<IQuickAccessVisibility>(MockBehavior.Strict);
            visibility.Setup(v => v.SetVisible(QuickAccess.All, true));
            _manager.Dispose();
            _manager = CreateManager(visibility.Object);

            _manager.ShowSection(QuickAccess.All);

            visibility.Verify(v => v.SetVisible(QuickAccess.All, true), Times.Once);
        }

        [TestMethod]
        public void ShowSection_WithOptions_SetsVisibleTrueAndRefreshes()
        {
            var visibility = new Mock<IQuickAccessVisibility>(MockBehavior.Strict);
            visibility.Setup(v => v.SetVisible(QuickAccess.All, true));
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)));
            _manager.Dispose();
            _manager = CreateManager(visibility.Object);

            _manager.ShowSection(QuickAccess.All, new VisibilityOptions { RefreshExplorer = true });

            visibility.Verify(v => v.SetVisible(QuickAccess.All, true), Times.Once);
            _explorerRefresher.Verify(r => r.Refresh(TimeSpan.FromSeconds(10)), Times.Once);
        }

        [TestMethod]
        public void HideSection_SetsVisibleFalse()
        {
            var visibility = new Mock<IQuickAccessVisibility>(MockBehavior.Strict);
            visibility.Setup(v => v.SetVisible(QuickAccess.All, false));
            _manager.Dispose();
            _manager = CreateManager(visibility.Object);

            _manager.HideSection(QuickAccess.All);

            visibility.Verify(v => v.SetVisible(QuickAccess.All, false), Times.Once);
        }

        [TestMethod]
        public void HideSection_WithOptions_SetsVisibleFalseAndRefreshes()
        {
            var visibility = new Mock<IQuickAccessVisibility>(MockBehavior.Strict);
            visibility.Setup(v => v.SetVisible(QuickAccess.All, false));
            _explorerRefresher.Setup(r => r.Refresh(TimeSpan.FromSeconds(10)));
            _manager.Dispose();
            _manager = CreateManager(visibility.Object);

            _manager.HideSection(QuickAccess.All, new VisibilityOptions { RefreshExplorer = true });

            visibility.Verify(v => v.SetVisible(QuickAccess.All, false), Times.Once);
            _explorerRefresher.Verify(r => r.Refresh(TimeSpan.FromSeconds(10)), Times.Once);
        }

        [TestMethod]
        public void GetRecentFilesMetadata_ParsesRecentBackingFile()
        {
            var reader = new Mock<IDestListMetadataReader>(MockBehavior.Strict);
            var entry = new DestListEntry { Path = @"C:\recent.txt" };
            reader.Setup(r => r.ParseFile(@"C:\recent.automaticDestinations-ms"))
                .Returns(CreateDestinations(entry));
            _manager.Dispose();
            _manager = CreateManager(reader.Object);

            var result = _manager.GetRecentFilesMetadata();

            Assert.AreSame(entry, result[0]);
            reader.Verify(r => r.ParseFile(@"C:\recent.automaticDestinations-ms"), Times.Once);
        }

        [TestMethod]
        public void GetFrequentFoldersMetadata_ParsesFrequentBackingFile()
        {
            var reader = new Mock<IDestListMetadataReader>(MockBehavior.Strict);
            var entry = new DestListEntry { Path = @"C:\Folder" };
            reader.Setup(r => r.ParseFile(@"C:\frequent.automaticDestinations-ms"))
                .Returns(CreateDestinations(entry));
            _manager.Dispose();
            _manager = CreateManager(reader.Object);

            var result = _manager.GetFrequentFoldersMetadata();

            Assert.AreSame(entry, result[0]);
            reader.Verify(r => r.ParseFile(@"C:\frequent.automaticDestinations-ms"), Times.Once);
        }

        [TestMethod]
        public void PublicApi_DoesNotExposeRemovedPhase0Surface()
        {
            var assembly = typeof(QuickAccessManager).Assembly;
            var managerMethods = typeof(QuickAccessManager).GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);

            Assert.IsFalse(managerMethods.Any(m => m.Name.EndsWith("Async", StringComparison.Ordinal)));
            Assert.IsFalse(managerMethods.Any(m => m.Name == "ClearCache"));
            Assert.IsNull(assembly.GetType("Wincent.IQuickAccessManager"));
            Assert.IsNull(assembly.GetType("Wincent.ExecutionFeasibilityStatus"));
        }

        [TestMethod]
        public void PublicApi_ExportsOnlySupportedSurface()
        {
            var expected = new List<string>
            {
                "Wincent.AddOptions",
                "Wincent.AutomaticDestinations",
                "Wincent.AutomaticDestinationsKind",
                "Wincent.BatchFailure",
                "Wincent.BatchOptions",
                "Wincent.BatchResult",
                "Wincent.CfbDirectoryEntry",
                "Wincent.CfbInfo",
                "Wincent.CfbObjectType",
                "Wincent.ClearOptions",
                "Wincent.ComApartmentMismatchException",
                "Wincent.DestList",
                "Wincent.DestListEntries",
                "Wincent.DestListEntry",
                "Wincent.DestListParseException",
                "Wincent.DestListUnsupportedVersionException",
                "Wincent.Diagnostic",
                "Wincent.DiagnosticSeverity",
                "Wincent.ExperimentalDestListRemoval",
                "Wincent.ExperimentalRemoveOptions",
                "Wincent.ExperimentalRemoveReport",
                "Wincent.FrequentRawPathRemoveReport",
                "Wincent.FrequentRestoreReport",
                "Wincent.PartialClearException",
                "Wincent.PathSource",
                "Wincent.PowerShellErrorKind",
                "Wincent.PowerShellExecutionException",
                "Wincent.PowerShellOperation",
                "Wincent.QuickAccess",
                "Wincent.QuickAccessItem",
                "Wincent.QuickAccessItemAlreadyExistsException",
                "Wincent.QuickAccessItemNotFoundException",
                "Wincent.QuickAccessLock",
                "Wincent.QuickAccessLockTarget",
                "Wincent.QuickAccessManager",
                "Wincent.QuickAccessManagerOptions",
                "Wincent.QuickAccessOperationException",
                "Wincent.QuickAccessPath",
                "Wincent.QuickAccessPostMutationException",
                "Wincent.QuickAccessPostMutationStep",
                "Wincent.QuickAccessUnlockFailure",
                "Wincent.QuickAccessUnlockOptions",
                "Wincent.QuickAccessUnlockReport",
                "Wincent.QuickAccessVisibilityException",
                "Wincent.RecentRestoreReport",
                "Wincent.RemoveOptions",
                "Wincent.RestoreDefaultsOptions",
                "Wincent.RestoreDefaultsReport",
                "Wincent.RetryPolicy",
                "Wincent.UnsupportedQuickAccessOperationException",
                "Wincent.VisibilityOptions",
                "Wincent.WincentException"
            };

            var exported = typeof(QuickAccessManager).Assembly
                .GetExportedTypes()
                .Select(t => t.FullName)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToList();

            expected.Sort(StringComparer.Ordinal);
            CollectionAssert.AreEqual(expected, exported);
        }

        private static PowerShellExecutionException CreatePowerShellException(
            PowerShellOperation operation,
            PowerShellErrorKind kind,
            string error)
        {
            return new PowerShellExecutionException(
                operation,
                kind,
                null,
                string.Empty,
                error,
                "test.ps1",
                null,
                TimeSpan.FromMilliseconds(1),
                null);
        }

        private QuickAccessManager CreateManager(IQuickAccessNativeQuery nativeQuery)
        {
            return CreateManager(nativeQuery, RetryPolicy.Standard);
        }

        private QuickAccessManager CreateManager(IQuickAccessNativeQuery nativeQuery, RetryPolicy retryPolicy)
        {
            return new QuickAccessManager(
                _executor.Object,
                TimeSpan.FromSeconds(10),
                _fileSystem.Object,
                _nativeMethods.Object,
                _dataFiles.Object,
                retryPolicy,
                nativeQuery);
        }

        private QuickAccessManager CreateManager(IQuickAccessNativeQuery nativeQuery, IQuickAccessNativeMutation nativeMutation)
        {
            return new QuickAccessManager(
                _executor.Object,
                TimeSpan.FromSeconds(10),
                _fileSystem.Object,
                _nativeMethods.Object,
                _dataFiles.Object,
                RetryPolicy.Standard,
                nativeQuery,
                nativeMutation);
        }

        private QuickAccessManager CreateManager(IQuickAccessLockFactory lockFactory)
        {
            return new QuickAccessManager(
                _executor.Object,
                TimeSpan.FromSeconds(10),
                _fileSystem.Object,
                _nativeMethods.Object,
                _dataFiles.Object,
                RetryPolicy.Standard,
                new PowerShellFallbackNativeQuery(),
                _nativeMutation.Object,
                _explorerRefresher.Object,
                _recentLinksCleaner.Object,
                lockFactory);
        }

        private QuickAccessManager CreateManager(IQuickAccessVisibility visibility)
        {
            return new QuickAccessManager(
                _executor.Object,
                TimeSpan.FromSeconds(10),
                _fileSystem.Object,
                _nativeMethods.Object,
                _dataFiles.Object,
                RetryPolicy.Standard,
                new PowerShellFallbackNativeQuery(),
                _nativeMutation.Object,
                _explorerRefresher.Object,
                _recentLinksCleaner.Object,
                new NoOpQuickAccessLockFactory(),
                visibility);
        }

        private QuickAccessManager CreateManager(IDestListMetadataReader destListReader)
        {
            return new QuickAccessManager(
                _executor.Object,
                TimeSpan.FromSeconds(10),
                _fileSystem.Object,
                _nativeMethods.Object,
                _dataFiles.Object,
                RetryPolicy.Standard,
                new PowerShellFallbackNativeQuery(),
                _nativeMutation.Object,
                _explorerRefresher.Object,
                _recentLinksCleaner.Object,
                new NoOpQuickAccessLockFactory(),
                new NoOpQuickAccessVisibility(),
                destListReader);
        }

        private static AutomaticDestinations CreateDestinations(params DestListEntry[] entries)
        {
            return new AutomaticDestinations(
                new CfbInfo(512, 64, 4096, Array.Empty<CfbDirectoryEntry>()),
                new DestList { Entries = entries });
        }

        private static QuickAccessLock CreateEmptyLock(QuickAccessLockTarget target)
        {
            return new QuickAccessLock(
                target,
                string.Empty,
                Array.Empty<string>(),
                Array.Empty<IQuickAccessBackingFileHandle>(),
                new DefaultRecentLinkFileSystem());
        }
    }
}
