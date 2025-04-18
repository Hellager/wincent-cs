using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Wincent;
using System.IO;
using System.Linq;
using System.Security;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Reflection;

namespace TestWincent
{
    [TestClass]
    public class QuickAccessManagerTests
    {
        private Mock<IScriptExecutor> _mockExecutor;
        private Mock<IFileSystemOperations> _mockFileSystem;
        private Mock<INativeMethods> _mockNativeMethods;
        private Mock<IQuickAccessDataFiles> _mockDataFiles;
        private IQuickAccessManager _manager;
        private readonly TimeSpan _timeout = TimeSpan.FromSeconds(5);

        [TestInitialize]
        public void Setup()
        {
            // 创建测试双对象
            _mockExecutor = new Mock<IScriptExecutor>(MockBehavior.Strict);
            _mockFileSystem = new Mock<IFileSystemOperations>(MockBehavior.Strict);
            _mockNativeMethods = new Mock<INativeMethods>(MockBehavior.Strict);
            _mockDataFiles = new Mock<IQuickAccessDataFiles>(MockBehavior.Strict);

            // 设置默认行为
            _mockExecutor.Setup(e => e.ExecutePSScriptWithCache(It.IsAny<PSScript>(), It.IsAny<string>(), It.IsAny<int>()))
                .ReturnsAsync(new List<string>());

            _mockExecutor.Setup(e => e.ExecutePSScript(It.IsAny<PSScript>(), It.IsAny<string>()))
                .ReturnsAsync(new ScriptResult(0, "", ""));

            _mockExecutor.Setup(e => e.ExecutePSScriptWithTimeout(It.IsAny<PSScript>(), It.IsAny<string>(), It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(0, "", ""));

            _mockExecutor.Setup(e => e.ClearCache()).Verifiable();

            _mockExecutor.Setup(e => e.Dispose()).Verifiable();

            // 文件系统模拟
            _mockFileSystem.Setup(fs => fs.FileExists(It.IsAny<string>())).Returns(true);
            _mockFileSystem.Setup(fs => fs.DirectoryExists(It.IsAny<string>())).Returns(true);
            _mockFileSystem.Setup(fs => fs.DeleteFile(It.IsAny<string>())).Verifiable();

            // 原生方法模拟
            _mockNativeMethods.Setup(n => n.CoInitializeEx(It.IsAny<IntPtr>(), It.IsAny<uint>())).Returns(0);
            _mockNativeMethods.Setup(n => n.SHAddToRecentDocs(It.IsAny<uint>(), It.IsAny<IntPtr>())).Verifiable();
            _mockNativeMethods.Setup(n => n.CoUninitialize()).Verifiable();

            // 修复 SHGetKnownFolderPath 的模拟
            IntPtr dummyPath = Marshal.StringToHGlobalUni(@"C:\Users\Test\AppData\Recent");
            _mockNativeMethods.Setup(n => n.SHGetKnownFolderPath(
                It.IsAny<Guid>(),
                It.IsAny<uint>(),
                It.IsAny<IntPtr>(),
                out It.Ref<IntPtr>.IsAny))
                .Callback(new SHGetKnownFolderPathCallback((Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr ppszPath) =>
                {
                    ppszPath = dummyPath;
                }))
                .Returns(0);

            _mockNativeMethods.Setup(n => n.CoTaskMemFree(It.IsAny<IntPtr>())).Verifiable();

            // 数据文件模拟
            _mockDataFiles.Setup(d => d.RemoveRecentFile()).Verifiable();
            _mockDataFiles.Setup(d => d.GetModifiedTimeForScript(It.IsAny<PSScript>())).Returns(DateTime.Now);

            // 设置数据文件路径属性
            _mockDataFiles.Setup(d => d.RecentFilesPath).Returns(@"C:\Users\Test\AppData\Recent\TestRecent.dat");
            _mockDataFiles.Setup(d => d.FrequentFoldersPath).Returns(@"C:\Users\Test\AppData\Recent\TestFrequent.dat");

            // 创建QuickAccessManager实例
            _manager = new QuickAccessManager(
                _mockExecutor.Object,
                _timeout,
                _mockFileSystem.Object,
                _mockNativeMethods.Object,
                _mockDataFiles.Object);
        }

        // 为了模拟带有 out 参数的方法而定义的委托类型
        private delegate void SHGetKnownFolderPathCallback(Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr ppszPath);

        [TestCleanup]
        public void Cleanup()
        {
            if (_manager != null)
            {
                _manager.Dispose();
                _manager = null;
            }
        }

        #region 可行性检查测试

        [TestMethod]
        public async Task CheckFeasibleAsync_ReturnsExpectedValues()
        {
            // 设置模拟执行器返回成功状态（退出码为0表示成功）
            _mockExecutor.Setup(e => e.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible,
                null,
                It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(0, "查询成功", ""));

            _mockExecutor.Setup(e => e.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible,
                null,
                It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(0, "操作成功", ""));

            // 执行
            var result = await _manager.CheckFeasibleAsync();

            // 验证
            Assert.IsTrue(result.QueryFeasible);
            Assert.IsTrue(result.HandleFeasible);

            // 验证调用了正确的方法
            _mockExecutor.Verify(e => e.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible,
                null,
                It.IsAny<int>()), Times.Once);

            _mockExecutor.Verify(e => e.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible,
                null,
                It.IsAny<int>()), Times.Once);
        }

        [TestMethod]
        public async Task CheckFeasibleAsync_OnlyQueryFeasible_ReturnsExpectedValues()
        {
            // 模拟执行可行性检查
            _mockExecutor.Setup(e => e.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible,
                null,
                It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(0, "success", ""));

            _mockExecutor.Setup(e => e.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible,
                null,
                It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(1, "", "error"));

            // 执行 - 重新创建管理器以刷新延迟加载
            _manager.Dispose();
            _manager = new QuickAccessManager(
                _mockExecutor.Object,
                _timeout,
                _mockFileSystem.Object,
                _mockNativeMethods.Object,
                _mockDataFiles.Object);

            var result = await _manager.CheckFeasibleAsync();

            // 验证
            Assert.IsTrue(result.QueryFeasible);
            Assert.IsFalse(result.HandleFeasible);
        }

        #endregion

        #region 路径验证测试

        [TestMethod]
        public void ValidatePath_ValidFilePath_NoException()
        {
            // 设置
            var mockFileSystem = new Mock<IFileSystemOperations>();
            mockFileSystem.Setup(fs => fs.FileExists(@"C:\test.txt")).Returns(true);

            // 执行 - 不应抛出异常
            QuickAccessManager.ValidatePath(@"C:\test.txt", PathType.File, mockFileSystem.Object);

            // 验证
            mockFileSystem.Verify(fs => fs.FileExists(@"C:\test.txt"), Times.Once);
        }

        [TestMethod]
        public void ValidatePath_ValidDirectoryPath_NoException()
        {
            // 设置
            var mockFileSystem = new Mock<IFileSystemOperations>();
            mockFileSystem.Setup(fs => fs.DirectoryExists(@"C:\TestFolder")).Returns(true);

            // 执行 - 不应抛出异常
            QuickAccessManager.ValidatePath(@"C:\TestFolder", PathType.Directory, mockFileSystem.Object);

            // 验证
            mockFileSystem.Verify(fs => fs.DirectoryExists(@"C:\TestFolder"), Times.Once);
        }

        [TestMethod]
        public void ValidatePath_AnyPathType_ChecksBoth()
        {
            // 设置
            var mockFileSystem = new Mock<IFileSystemOperations>();
            mockFileSystem.Setup(fs => fs.FileExists(@"C:\test.txt")).Returns(false);
            mockFileSystem.Setup(fs => fs.DirectoryExists(@"C:\test.txt")).Returns(true);

            // 执行 - 不应抛出异常
            QuickAccessManager.ValidatePath(@"C:\test.txt", PathType.Any, mockFileSystem.Object);

            // 验证
            mockFileSystem.Verify(fs => fs.FileExists(@"C:\test.txt"), Times.Once);
            mockFileSystem.Verify(fs => fs.DirectoryExists(@"C:\test.txt"), Times.Once);
        }

        [TestMethod]
        [ExpectedException(typeof(SecurityException))]
        public void ValidatePath_SystemPathSystem32_ThrowsSecurityException()
        {
            // 执行
            QuickAccessManager.ValidatePath(@"C:\Windows\System32\test.dll", PathType.File, _mockFileSystem.Object);
        }

        [TestMethod]
        [ExpectedException(typeof(SecurityException))]
        public void ValidatePath_ProgramFilesPath_ThrowsSecurityException()
        {
            // 执行
            QuickAccessManager.ValidatePath(@"C:\Program Files\test\app.exe", PathType.File, _mockFileSystem.Object);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ValidatePath_EmptyPath_ThrowsArgumentException()
        {
            // 执行
            QuickAccessManager.ValidatePath("", PathType.File, _mockFileSystem.Object);
        }

        [TestMethod]
        [ExpectedException(typeof(FileNotFoundException))]
        public void ValidatePath_NonExistentFile_ThrowsFileNotFoundException()
        {
            // 设置
            _mockFileSystem.Setup(fs => fs.FileExists(@"C:\nonexistent.txt")).Returns(false);

            // 执行
            QuickAccessManager.ValidatePath(@"C:\nonexistent.txt", PathType.File, _mockFileSystem.Object);
        }

        #endregion

        #region 项目操作测试

        [TestMethod]
        public async Task GetItemsAsync_QueryQuickAccess_CallsCorrectScript()
        {
            // 设置
            var expectedItems = new List<string> { @"C:\test1.txt", @"C:\test2.txt" };
            _mockExecutor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null, 10))
                .ReturnsAsync(expectedItems);

            // 执行
            var result = await _manager.GetItemsAsync(QuickAccess.All);

            // 验证
            Assert.AreEqual(expectedItems.Count, result.Count);
            Assert.AreEqual(expectedItems[0], result[0]);
            Assert.AreEqual(expectedItems[1], result[1]);
            _mockExecutor.Verify(e => e.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null, 10), Times.Once);
        }

        [TestMethod]
        public async Task GetItemsAsync_QueryRecentFiles_CallsCorrectScript()
        {
            // 设置
            var expectedItems = new List<string> { @"C:\test1.txt", @"C:\test2.txt" };
            _mockExecutor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(expectedItems);

            // 执行
            var result = await _manager.GetItemsAsync(QuickAccess.RecentFiles);

            // 验证
            _mockExecutor.Verify(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10), Times.Once);
        }

        [TestMethod]
        public async Task GetItemsAsync_QueryFrequentFolders_CallsCorrectScript()
        {
            // 设置
            var expectedItems = new List<string> { @"C:\Folder1", @"C:\Folder2" };
            _mockExecutor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(expectedItems);

            // 执行
            var result = await _manager.GetItemsAsync(QuickAccess.FrequentFolders);

            // 验证
            _mockExecutor.Verify(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10), Times.Once);
        }

        [TestMethod]
        public async Task CheckItemAsync_ItemExists_ReturnsTrue()
        {
            // 设置
            var items = new List<string> { @"C:\test.txt", @"C:\another.txt" };
            _mockExecutor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(items);

            // 执行
            var result = await _manager.CheckItemAsync(@"C:\test.txt", QuickAccess.RecentFiles);

            // 验证
            Assert.IsTrue(result);
        }

        [TestMethod]
        public async Task CheckItemAsync_ItemDoesNotExist_ReturnsFalse()
        {
            // 设置
            var items = new List<string> { @"C:\test.txt", @"C:\another.txt" };
            _mockExecutor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(items);

            // 执行
            var result = await _manager.CheckItemAsync(@"C:\not-exists.txt", QuickAccess.RecentFiles);

            // 验证
            Assert.IsFalse(result);
        }

        [TestMethod]
        public async Task AddItemAsync_RecentFile_ShouldCallNativeMethod()
        {
            // 设置
            _mockExecutor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string>());

            // 执行
            var result = await _manager.AddItemAsync(@"C:\test.txt", QuickAccess.RecentFiles);

            // 验证
            Assert.IsTrue(result);
            _mockNativeMethods.Verify(n => n.SHAddToRecentDocs(It.IsAny<uint>(), It.IsAny<IntPtr>()), Times.Once);
        }

        [TestMethod]
        public async Task AddItemAsync_RecentFileWithForceUpdate_ShouldCallRemoveRecentFile()
        {
            // 设置
            _mockExecutor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string>());

            // 执行
            var result = await _manager.AddItemAsync(@"C:\test.txt", QuickAccess.RecentFiles, true);

            // 验证
            Assert.IsTrue(result);
            _mockDataFiles.Verify(d => d.RemoveRecentFile(), Times.Once);
        }

        [TestMethod]
        public async Task AddItemAsync_FrequentFolder_ShouldCallCorrectScript()
        {
            // 设置
            _mockExecutor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string>());

            // 执行
            var result = await _manager.AddItemAsync(@"C:\TestFolder", QuickAccess.FrequentFolders);

            // 验证
            Assert.IsTrue(result);
            _mockExecutor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.PinToFrequentFolder, @"C:\TestFolder", 10), Times.Once);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public async Task AddItemAsync_ItemAlreadyExists_ThrowsInvalidOperationException()
        {
            // 设置
            _mockExecutor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { @"C:\test.txt" });

            // 执行 - 应该抛出异常
            await _manager.AddItemAsync(@"C:\test.txt", QuickAccess.RecentFiles);
        }

        [TestMethod]
        public async Task RemoveItemAsync_RecentFile_ShouldCallCorrectScript()
        {
            // 设置
            _mockExecutor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string> { @"C:\test.txt" });

            // 执行
            var result = await _manager.RemoveItemAsync(@"C:\test.txt", QuickAccess.RecentFiles);

            // 验证
            Assert.IsTrue(result);
            _mockExecutor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.RemoveRecentFile, @"C:\test.txt", 10), Times.Once);
            _mockExecutor.Verify(e => e.ClearCache(), Times.Once);
        }

        [TestMethod]
        public async Task RemoveItemAsync_FrequentFolder_ShouldCallCorrectScript()
        {
            // 设置
            _mockExecutor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string> { @"C:\TestFolder" });

            // 执行
            var result = await _manager.RemoveItemAsync(@"C:\TestFolder", QuickAccess.FrequentFolders);

            // 验证
            Assert.IsTrue(result);
            _mockExecutor.Verify(e => e.ExecutePSScriptWithTimeout(PSScript.UnpinFromFrequentFolder, @"C:\TestFolder", 10), Times.Once);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public async Task RemoveItemAsync_ItemDoesNotExist_ThrowsInvalidOperationException()
        {
            // 设置
            _mockExecutor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryRecentFile, null, 10))
                .ReturnsAsync(new List<string>());

            // 执行 - 应该抛出异常
            await _manager.RemoveItemAsync(@"C:\test.txt", QuickAccess.RecentFiles);
        }

        #endregion

        #region 批量操作测试

        [TestMethod]
        public async Task EmptyItemsAsync_RecentFiles_ShouldCallEmptyRecentFiles()
        {
            // 执行
            var result = await _manager.EmptyItemsAsync(QuickAccess.RecentFiles);

            // 验证
            Assert.IsTrue(result);
            _mockNativeMethods.Verify(n => n.SHAddToRecentDocs(It.IsAny<uint>(), IntPtr.Zero), Times.Once);
            _mockExecutor.Verify(e => e.ClearCache(), Times.Once);
        }

        [TestMethod]
        public async Task EmptyItemsAsync_FrequentFolders_ShouldDeleteJumplistFile()
        {
            // 设置
            _mockFileSystem.Setup(fs => fs.FileExists(It.IsAny<string>())).Returns(true);

            // 执行
            var result = await _manager.EmptyItemsAsync(QuickAccess.FrequentFolders);

            // 验证
            Assert.IsTrue(result);
            _mockFileSystem.Verify(fs => fs.DeleteFile(It.IsAny<string>()), Times.Once);
            _mockExecutor.Verify(e => e.ClearCache(), Times.Once);
        }

        [TestMethod]
        public async Task EmptyItemsAsync_All_ShouldCallBothEmptyMethods()
        {
            // 设置
            _mockFileSystem.Setup(fs => fs.FileExists(It.IsAny<string>())).Returns(true);

            // 执行
            var result = await _manager.EmptyItemsAsync(QuickAccess.All);

            // 验证
            Assert.IsTrue(result);
            _mockNativeMethods.Verify(n => n.SHAddToRecentDocs(It.IsAny<uint>(), IntPtr.Zero), Times.Once);
            _mockFileSystem.Verify(fs => fs.DeleteFile(It.IsAny<string>()), Times.Once);
            _mockExecutor.Verify(e => e.ClearCache(), Times.AtLeastOnce);
        }

        [TestMethod]
        public async Task EmptyItemsAsync_WithForceRefresh_ShouldCallRefreshExplorer()
        {
            // 执行
            var result = await _manager.EmptyItemsAsync(QuickAccess.RecentFiles, true);

            // 验证
            Assert.IsTrue(result);
            _mockExecutor.Verify(e => e.ExecutePSScript(PSScript.RefreshExplorer, null), Times.Once);
        }

        [TestMethod]
        public async Task EmptyItemsAsync_FrequentFoldersWithSystemDefault_ShouldCallEmptyPinnedFolders()
        {
            // 执行
            var result = await _manager.EmptyItemsAsync(QuickAccess.FrequentFolders, false, true);

            // 验证
            Assert.IsTrue(result);
            _mockExecutor.Verify(e => e.ExecutePSScriptWithCache(PSScript.EmptyPinnedFolders, null, 0), Times.Once);
        }

        #endregion

        #region 缓存控制测试

        [TestMethod]
        public void ClearCache_CallsExecutorClearCache()
        {
            // 执行
            _manager.ClearCache();

            // 验证
            _mockExecutor.Verify(e => e.ClearCache(), Times.Once);
        }

        #endregion

        #region 异常处理测试

        [TestMethod]
        public async Task HandleOperationAsync_ExecutorReturnsNonZeroExitCode_ReturnsFalse()
        {
            // 设置
            _mockExecutor.Setup(e => e.ExecutePSScriptWithCache(PSScript.QueryFrequentFolder, null, 10))
                .ReturnsAsync(new List<string>());

            _mockExecutor.Setup(e => e.ExecutePSScriptWithTimeout(PSScript.PinToFrequentFolder, @"C:\TestFolder", 10))
                .ReturnsAsync(new ScriptResult(1, "", "Error occurred"));

            // 执行
            var result = await _manager.AddItemAsync(@"C:\TestFolder", QuickAccess.FrequentFolders);

            // 验证
            Assert.IsFalse(result);
        }

        [TestMethod]
        public async Task EmptyItemsAsync_ExceptionInOperation_ReturnsFalse()
        {
            // 设置
            _mockNativeMethods.Setup(n => n.CoInitializeEx(It.IsAny<IntPtr>(), It.IsAny<uint>()))
                .Throws(new Win32Exception(1, "COM initialization failed"));

            // 执行
            var result = await _manager.EmptyItemsAsync(QuickAccess.RecentFiles);

            // 验证
            Assert.IsFalse(result);
        }

        #endregion
    }

    /// <summary>
    /// 用于替代文件系统静态方法的辅助类
    /// </summary>
    public class FileSystemHelper : IDisposable
    {
        // 保存原始的文件系统检查方法
        private readonly Func<string, bool> _originalFileExists;
        private readonly Func<string, bool> _originalDirectoryExists;

        // 用于测试的替代字典
        private readonly Dictionary<string, bool> _fileExistsMap = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, bool> _directoryExistsMap = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

        public FileSystemHelper()
        {
            // 保存原始方法引用（在实际测试中，您可能需要使用反射或其他方式替换静态方法）
            _originalFileExists = File.Exists;
            _originalDirectoryExists = Directory.Exists;

            // 这里应使用反射或其他方式替换静态方法，此处仅为示例
            // ReplacementHelper.ReplaceMethod(typeof(File), "Exists", new Func<string, bool>(TestFileExists));
            // ReplacementHelper.ReplaceMethod(typeof(Directory), "Exists", new Func<string, bool>(TestDirectoryExists));
        }

        public void SetupFileExists(string path, bool exists)
        {
            _fileExistsMap[path] = exists;
        }

        public void SetupDirectoryExists(string path, bool exists)
        {
            _directoryExistsMap[path] = exists;
        }

        private bool TestFileExists(string path)
        {
            return _fileExistsMap.TryGetValue(path, out bool exists) ? exists : _originalFileExists(path);
        }

        private bool TestDirectoryExists(string path)
        {
            return _directoryExistsMap.TryGetValue(path, out bool exists) ? exists : _originalDirectoryExists(path);
        }

        public void Dispose()
        {
            // 恢复原始方法（在实际测试中，您可能需要使用反射或其他方式恢复静态方法）
            // ReplacementHelper.RestoreMethod(typeof(File), "Exists");
            // ReplacementHelper.RestoreMethod(typeof(Directory), "Exists");
        }
    }
}
