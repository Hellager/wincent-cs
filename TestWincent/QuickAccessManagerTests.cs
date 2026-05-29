using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
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
        public void GetItems_All_UsesQuickAccessQuery()
        {
            var expected = new List<string> { @"C:\a.txt", @"C:\folder" };
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null, 10))
                .ReturnsAsync(expected);

            var result = _manager.GetItems(QuickAccess.All);

            CollectionAssert.AreEqual(expected, result.ToList());
            _executor.Verify(e => e.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null, 10), Times.Once);
        }

        [TestMethod]
        public void GetItems_NativeSuccess_DoesNotUsePowerShellFallback()
        {
            var nativeQuery = new Mock<IQuickAccessNativeQuery>(MockBehavior.Strict);
            nativeQuery.Setup(n => n.GetItems(QuickAccess.RecentFiles))
                .Returns(new List<string> { @"C:\native.txt" });
            _manager.Dispose();
            _manager = CreateManager(nativeQuery.Object);

            var result = _manager.GetItems(QuickAccess.RecentFiles);

            CollectionAssert.AreEqual(new[] { @"C:\native.txt" }, result.ToList());
            _executor.Verify(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, It.IsAny<int>()), Times.Never);
        }

        [TestMethod]
        public void GetItems_NativeFailure_FallsBackToPowerShell()
        {
            var nativeQuery = new Mock<IQuickAccessNativeQuery>(MockBehavior.Strict);
            nativeQuery.Setup(n => n.GetItems(QuickAccess.FrequentFolders))
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
        public void GetItems_NativeAndPowerShellFailure_ThrowsPowerShellExecutionException()
        {
            var nativeQuery = new Mock<IQuickAccessNativeQuery>(MockBehavior.Strict);
            nativeQuery.Setup(n => n.GetItems(QuickAccess.All))
                .Throws(new InvalidOperationException("native failed"));
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null, 10))
                .Throws(CreatePowerShellException(PowerShellOperation.QueryQuickAccess, PowerShellErrorKind.ProcessFailed, "fallback failed"));
            _manager.Dispose();
            _manager = CreateManager(nativeQuery.Object);

            var exception = Assert.ThrowsException<PowerShellExecutionException>(
                () => _manager.GetItems(QuickAccess.All));

            Assert.AreEqual(PowerShellOperation.QueryQuickAccess, exception.Operation);
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
            nativeQuery.Setup(n => n.GetItems(QuickAccess.RecentFiles))
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
            nativeQuery.Setup(n => n.GetItems(QuickAccess.FrequentFolders))
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
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string>());

            _manager.AddItem(@"C:\test.txt", QuickAccess.RecentFiles);

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
                .Throws(new InvalidOperationException("native failed"));

            _manager.AddItem(@"C:\Folder", QuickAccess.FrequentFolders);

            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.PinToFrequentFolder, @"C:\Folder", 10), Times.Once);
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
        public void RemoveItem_RecentFile_NativeFailureFallsBackToPowerShell()
        {
            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { @"C:\test.txt" });
            _nativeMutation.Setup(m => m.RemoveRecentFile(@"C:\test.txt", TimeSpan.FromSeconds(10)))
                .Throws(new InvalidOperationException("native failed"));

            _manager.RemoveItem(@"C:\test.txt", QuickAccess.RecentFiles);

            _executor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.RemoveRecentFile, @"C:\test.txt", 10), Times.Once);
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
                .Throws(new InvalidOperationException("native failed"));
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.UnpinFromFrequentFolder, @"C:\Folder", 10))
                .Throws(CreatePowerShellException(PowerShellOperation.UnpinFrequentFolder, PowerShellErrorKind.ProcessFailed, "err"));

            _manager.RemoveItem(@"C:\Folder", QuickAccess.FrequentFolders);
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
                new RetryPolicy(1, TimeSpan.FromMilliseconds(1), TimeSpan.FromMilliseconds(1), 1.0, false));

            _executor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string> { @"C:\Folder" });
            _executor.SetupSequence(e => e.ExecutePSScriptWithTimeout(PSScript.UnpinFromFrequentFolder, @"C:\Folder", 10))
                .Throws(CreatePowerShellException(PowerShellOperation.UnpinFrequentFolder, PowerShellErrorKind.Timeout, "operation timed out"))
                .ReturnsAsync(new ScriptResult(0, string.Empty, string.Empty));

            _manager.RemoveItem(@"C:\Folder", QuickAccess.FrequentFolders);

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
            _manager.ClearItems(QuickAccess.RecentFiles);

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
            _nativeMethods.Setup(n => n.CoInitializeEx(It.IsAny<IntPtr>(), It.IsAny<uint>()))
                .Throws(new Win32Exception(1));

            try
            {
                _manager.ClearItems(QuickAccess.All);
                Assert.Fail("Expected PartialClearException.");
            }
            catch (PartialClearException ex)
            {
                Assert.IsFalse(ex.RecentFilesCleared);
                Assert.IsTrue(ex.FrequentFoldersCleared);
            }
        }

        [TestMethod]
        public void ClearItems_FrequentFoldersPinnedCleanupFailure_ReportsFrequentProgress()
        {
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.EmptyPinnedFolders, null, 10))
                .Throws(CreatePowerShellException(PowerShellOperation.ClearPinnedFolders, PowerShellErrorKind.ProcessFailed, "unpin failed"));

            try
            {
                _manager.ClearItems(QuickAccess.FrequentFolders, new ClearOptions { RemovePinnedFolders = true });
                Assert.Fail("Expected PartialClearException.");
            }
            catch (PartialClearException ex)
            {
                Assert.IsFalse(ex.RecentFilesCleared);
                Assert.IsTrue(ex.FrequentFoldersCleared);
            }
        }

        [TestMethod]
        public void ClearItems_AllFrequentPartialFailure_MergesProgressFlags()
        {
            _executor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.EmptyPinnedFolders, null, 10))
                .Throws(CreatePowerShellException(PowerShellOperation.ClearPinnedFolders, PowerShellErrorKind.ProcessFailed, "unpin failed"));

            try
            {
                _manager.ClearItems(QuickAccess.All, new ClearOptions { RemovePinnedFolders = true });
                Assert.Fail("Expected PartialClearException.");
            }
            catch (PartialClearException ex)
            {
                Assert.IsTrue(ex.RecentFilesCleared);
                Assert.IsTrue(ex.FrequentFoldersCleared);
            }
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

            try
            {
                _manager.ClearItems(QuickAccess.All, new ClearOptions { RefreshExplorer = true });
                Assert.Fail("Expected PartialClearException.");
            }
            catch (PartialClearException ex)
            {
                Assert.IsFalse(ex.RecentFilesCleared);
                Assert.IsTrue(ex.FrequentFoldersCleared);
            }
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
        public void AddItems_RefreshFailureRecordsPerItemFailure()
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
            Assert.AreSame(refreshError, result.Failed[0].Error);
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
        public void PublicApi_DoesNotExposeRemovedPhase0Surface()
        {
            var assembly = typeof(QuickAccessManager).Assembly;
            var managerMethods = typeof(QuickAccessManager).GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);

            Assert.IsFalse(managerMethods.Any(m => m.Name.EndsWith("Async", StringComparison.Ordinal)));
            Assert.IsFalse(managerMethods.Any(m => m.Name == "ClearCache"));
            Assert.IsFalse(managerMethods.Any(m => m.Name.StartsWith("Lock", StringComparison.Ordinal)));
            Assert.IsFalse(managerMethods.Any(m => m.Name.Contains("Visible")));
            Assert.IsFalse(managerMethods.Any(m => m.Name.EndsWith("Metadata", StringComparison.Ordinal)));
            Assert.IsNull(assembly.GetType("Wincent.IQuickAccessManager"));
            Assert.IsNull(assembly.GetType("Wincent.ExecutionFeasibilityStatus"));
        }

        [TestMethod]
        public void PublicApi_ExportsOnlyPhase0Surface()
        {
            var expected = new List<string>
            {
                "Wincent.AddOptions",
                "Wincent.AutomaticDestinations",
                "Wincent.BatchFailure",
                "Wincent.BatchOptions",
                "Wincent.BatchResult",
                "Wincent.CfbDirectoryEntry",
                "Wincent.CfbInfo",
                "Wincent.CfbObjectType",
                "Wincent.ClearOptions",
                "Wincent.ComApartmentMismatchException",
                "Wincent.DestList",
                "Wincent.DestListEntry",
                "Wincent.DestListParseException",
                "Wincent.DestListUnsupportedVersionException",
                "Wincent.PartialClearException",
                "Wincent.PowerShellErrorKind",
                "Wincent.PowerShellExecutionException",
                "Wincent.PowerShellOperation",
                "Wincent.QuickAccess",
                "Wincent.QuickAccessItem",
                "Wincent.QuickAccessItemAlreadyExistsException",
                "Wincent.QuickAccessItemNotFoundException",
                "Wincent.QuickAccessLockTarget",
                "Wincent.QuickAccessManager",
                "Wincent.QuickAccessManagerOptions",
                "Wincent.QuickAccessOperationException",
                "Wincent.QuickAccessUnlockFailure",
                "Wincent.QuickAccessUnlockOptions",
                "Wincent.QuickAccessUnlockReport",
                "Wincent.RemoveOptions",
                "Wincent.RetryPolicy",
                "Wincent.UnsupportedQuickAccessOperationException",
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
            return new QuickAccessManager(
                _executor.Object,
                TimeSpan.FromSeconds(10),
                _fileSystem.Object,
                _nativeMethods.Object,
                _dataFiles.Object,
                RetryPolicy.Standard,
                nativeQuery);
        }
    }
}
