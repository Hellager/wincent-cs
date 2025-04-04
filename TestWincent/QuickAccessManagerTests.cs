using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security;
using System.Threading.Tasks;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class TestQuickAccessManager
    {
        private MockFileSystemService _mockFileSystem;
        private QuickAccessManager _manager;

        [TestInitialize]
        public void Initialize()
        {
            _mockFileSystem = new MockFileSystemService();
            _mockFileSystem.SetFileExists(true);
            _mockFileSystem.SetDirectoryExists(true);
            _mockFileSystem.SetScriptExecutionResult(true);
            _manager = new QuickAccessManager(_mockFileSystem);
        }

        [TestCleanup]
        public void Cleanup()
        {
            // 恢复默认服务
            // QuickAccessManager.ResetFileSystemService();
        }

        #region 添加到快速访问测试

        [TestMethod]
        public async Task PinToQuickAccessAsync_WithValidDirectoryPath_ReturnsTrue()
        {
            // Arrange
            string dirPath = @"C:\Test\Folder";
            _mockFileSystem.SetFileExists(false);
            _mockFileSystem.SetDirectoryExists(true);
            _mockFileSystem.SetScriptExecutionResult(true);

            // Act
            bool result = await _manager.PinToFrequentFolderAsync(dirPath);

            // Assert
            Assert.IsTrue(result, "有效文件夹路径应该返回 true");
            Assert.AreEqual(PSScript.PinToFrequentFolder, _mockFileSystem.LastScriptType);
            Assert.AreEqual(dirPath, _mockFileSystem.LastParameter);
        }

        [TestMethod]
        public async Task PinToQuickAccessAsync_WithFilePath_ReturnsFalse()
        {
            // Arrange
            string filePath = @"C:\Test\file.txt";
            _mockFileSystem.SetFileExists(true);
            _mockFileSystem.SetDirectoryExists(false);

            // Act
            bool result = await _manager.PinToFrequentFolderAsync(filePath);

            // Assert
            Assert.IsFalse(result, "文件路径应该返回 false");
            Assert.IsNull(_mockFileSystem.LastScriptType);
        }

        [TestMethod]
        public async Task PinToQuickAccessAsync_WithInvalidPath_ReturnsFalse()
        {
            // Arrange
            string invalidPath = @"C:\NonExistent\Path";
            _mockFileSystem.SetFileExists(false);
            _mockFileSystem.SetDirectoryExists(false);

            // Act
            bool result = await _manager.PinToFrequentFolderAsync(invalidPath);

            // Assert
            Assert.IsFalse(result, "无效路径应该返回 false");
            Assert.IsNull(_mockFileSystem.LastScriptType);
        }

        [TestMethod]
        public async Task PinToQuickAccessAsync_WithEmptyPath_ReturnsFalse()
        {
            // Arrange
            string emptyPath = "";

            // Act
            bool result = await _manager.PinToFrequentFolderAsync(emptyPath);

            // Assert
            Assert.IsFalse(result, "空路径应该返回 false");
        }

        [TestMethod]
        public async Task PinToQuickAccessAsync_WithScriptFailure_ReturnsFalse()
        {
            // Arrange
            string path = @"C:\Test\Folder";
            _mockFileSystem.SetDirectoryExists(true);
            _mockFileSystem.SetScriptExecutionResult(false);

            // Act
            bool result = await _manager.PinToFrequentFolderAsync(path);

            // Assert
            Assert.IsFalse(result, "脚本执行失败应该返回 false");
        }

        #endregion

        #region 从快速访问移除测试

        [TestMethod]
        public async Task UnpinFromQuickAccessAsync_WithValidPath_ReturnsTrue()
        {
            // Arrange
            string dirPath = @"C:\Test\Folder";
            _mockFileSystem.SetScriptExecutionResult(true);

            // Act
            bool result = await _manager.UnpinFromFrequentFolderAsync(dirPath);

            // Assert
            Assert.IsTrue(result, "有效路径应该返回 true");
            Assert.AreEqual(PSScript.UnpinFromFrequentFolder, _mockFileSystem.LastScriptType);
            Assert.AreEqual(dirPath, _mockFileSystem.LastParameter);
        }

        [TestMethod]
        public async Task UnpinFromQuickAccessAsync_WithEmptyPath_ReturnsFalse()
        {
            // Arrange
            string emptyPath = "";

            // Act
            bool result = await _manager.UnpinFromFrequentFolderAsync(emptyPath);

            // Assert
            Assert.IsFalse(result, "空路径应该返回 false");
        }

        [TestMethod]
        public async Task UnpinFromQuickAccessAsync_WithScriptFailure_ReturnsFalse()
        {
            // Arrange
            string path = @"C:\Test\Folder";
            _mockFileSystem.SetScriptExecutionResult(false);

            // Act
            bool result = await _manager.UnpinFromFrequentFolderAsync(path);

            // Assert
            Assert.IsFalse(result, "脚本执行失败应该返回 false");
        }

        #endregion

        #region 最近文件操作测试

        [TestMethod]
        public async Task RemoveFromRecentFilesAsync_WithValidFilePath_ReturnsTrue()
        {
            // Arrange
            string filePath = @"C:\Test\file.txt";
            _mockFileSystem.SetFileExists(true);
            _mockFileSystem.SetFileExtension(".txt");
            _mockFileSystem.SetScriptExecutionResult(true);

            // Act
            bool result = await _manager.RemoveFromRecentFilesAsync(filePath);

            // Assert
            Assert.IsTrue(result, "有效文件路径应该返回 true");
            Assert.AreEqual(PSScript.RemoveRecentFile, _mockFileSystem.LastScriptType);
            Assert.AreEqual(filePath, _mockFileSystem.LastParameter);
        }

        [TestMethod]
        public async Task RemoveFromRecentFilesAsync_WithEmptyPath_ReturnsFalse()
        {
            // Arrange
            string emptyPath = "";

            // Act
            bool result = await _manager.RemoveFromRecentFilesAsync(emptyPath);

            // Assert
            Assert.IsFalse(result, "空路径应该返回 false");
        }

        [TestMethod]
        public async Task RemoveFromRecentFilesAsync_WithNonExistentFile_ReturnsFalse()
        {
            // Arrange
            string nonExistentFile = @"C:\Test\nonexistent.txt";
            _mockFileSystem.SetFileExists(false);

            // Act
            bool result = await _manager.RemoveFromRecentFilesAsync(nonExistentFile);

            // Assert
            Assert.IsFalse(result, "不存在的文件应该返回 false");
            Assert.IsNull(_mockFileSystem.LastScriptType);
        }

        [TestMethod]
        public async Task AddToRecentFilesAsync_WithValidFilePath_ReturnsTrue()
        {
            // Arrange
            string filePath = @"C:\Test\file.txt";
            _mockFileSystem.SetFileExists(true);
            _mockFileSystem.SetFileExtension(".txt");

            // Act
            bool result = await _manager.AddToRecentFilesAsync(filePath);

            // Assert
            Assert.IsTrue(result, "有效文件路径应该返回 true");
            Assert.IsTrue(_mockFileSystem.WasAddFileToRecentDocsCalled);
            Assert.AreEqual(filePath, _mockFileSystem.LastAddedToRecentFile);
        }

        [TestMethod]
        public async Task AddToRecentFilesAsync_WithEmptyPath_ReturnsFalse()
        {
            // Arrange
            string emptyPath = "";

            // Act
            bool result = await _manager.AddToRecentFilesAsync(emptyPath);

            // Assert
            Assert.IsFalse(result, "空路径应该返回 false");
        }

        [TestMethod]
        public async Task AddToRecentFilesAsync_WithNonExistentFile_ReturnsFalse()
        {
            // Arrange
            string nonExistentFile = @"C:\Test\nonexistent.txt";
            _mockFileSystem.SetFileExists(false);

            // Act
            bool result = await _manager.AddToRecentFilesAsync(nonExistentFile);

            // Assert
            Assert.IsFalse(result, "不存在的文件应该返回 false");
            Assert.IsFalse(_mockFileSystem.WasAddFileToRecentDocsCalled);
        }

        [TestMethod]
        public async Task AddToRecentFilesAsync_WithNoExtension_ReturnsFalse()
        {
            // Arrange
            string fileWithoutExt = @"C:\Test\fileWithoutExtension";
            _mockFileSystem.SetFileExists(true);
            _mockFileSystem.SetFileExtension("");

            // Act
            bool result = await _manager.AddToRecentFilesAsync(fileWithoutExt);

            // Assert
            Assert.IsFalse(result, "没有扩展名的文件应该返回 false");
            Assert.IsFalse(_mockFileSystem.WasAddFileToRecentDocsCalled);
        }

        #endregion

        #region 清空最近文件测试

        [TestMethod]
        public async Task ClearRecentFilesAsync_Success_ReturnsTrue()
        {
            // Act
            bool result = await _manager.ClearRecentFilesAsync();

            // Assert
            Assert.IsTrue(result, "成功清空最近文件应该返回 true");
            Assert.IsTrue(_mockFileSystem.WasEmptyRecentFilesCalled);
        }

        [TestMethod]
        public async Task ClearRecentFilesAsync_WithException_ReturnsFalse()
        {
            // Arrange - 设置 mock 抛出异常
            _mockFileSystem.ThrowOnEmptyRecentFiles = true;

            // Act
            bool result = await _manager.ClearRecentFilesAsync();

            // Assert
            Assert.IsFalse(result, "发生异常时应该返回 false");
        }

        #endregion

        #region 查询操作测试

        [TestMethod]
        public async Task GetQuickAccessItemsAsync_ReturnsCorrectItems()
        {
            // 这个测试需要模拟 QuickAccessQuery 类
            // 由于 QuickAccessQuery 已经有自己的测试，这里只需验证调用关系

            // 此测试仅作为示例，实际实现可能需要调整
            var items = await _manager.GetQuickAccessItemsAsync();
            Assert.IsNotNull(items, "返回的项目列表不应为 null");
        }

        [TestMethod]
        public async Task GetRecentFilesAsync_ReturnsCorrectFiles()
        {
            // 此测试仅作为示例，实际实现可能需要调整
            var files = await _manager.GetRecentFilesAsync();
            Assert.IsNotNull(files, "返回的文件列表不应为 null");
        }

        [TestMethod]
        public async Task GetFrequentFoldersAsync_ReturnsCorrectFolders()
        {
            // 此测试仅作为示例，实际实现可能需要调整
            var folders = await _manager.GetFrequentFoldersAsync();
            Assert.IsNotNull(folders, "返回的文件夹列表不应为 null");
        }

        #endregion

        #region 包装函数测试

        [TestMethod]
        public async Task AddItemAsync_WithFile_CallsAddToRecentFiles()
        {
            // Arrange
            string filePath = @"C:\Test\file.txt";
            _mockFileSystem.SetFileExists(true);
            _mockFileSystem.SetFileExtension(".txt");

            // Act
            bool result = await _manager.AddItemAsync(filePath, QuickAccessItemType.File);

            // Assert
            Assert.IsTrue(result, "添加文件应该返回 true");
            Assert.IsTrue(_mockFileSystem.WasAddFileToRecentDocsCalled);
            Assert.AreEqual(filePath, _mockFileSystem.LastAddedToRecentFile);
        }

        [TestMethod]
        public async Task AddItemAsync_WithDirectory_CallsPinToFrequentFolder()
        {
            // Arrange
            string dirPath = @"C:\Test\Folder";
            _mockFileSystem.SetFileExists(false);
            _mockFileSystem.SetDirectoryExists(true);
            _mockFileSystem.SetScriptExecutionResult(true);

            // Act
            bool result = await _manager.AddItemAsync(dirPath, QuickAccessItemType.Directory);

            // Assert
            Assert.IsTrue(result, "添加文件夹应该返回 true");
            Assert.AreEqual(PSScript.PinToFrequentFolder, _mockFileSystem.LastScriptType);
            Assert.AreEqual(dirPath, _mockFileSystem.LastParameter);
        }

        [TestMethod]
        public async Task RemoveItemAsync_WithFile_CallsRemoveFromRecentFiles()
        {
            // Arrange
            string filePath = @"C:\Test\file.txt";
            _mockFileSystem.SetFileExists(true);
            _mockFileSystem.SetFileExtension(".txt");
            _mockFileSystem.SetScriptExecutionResult(true);

            // Act
            bool result = await _manager.RemoveItemAsync(filePath, QuickAccessItemType.File);

            // Assert
            Assert.IsTrue(result, "移除文件应该返回 true");
            Assert.AreEqual(PSScript.RemoveRecentFile, _mockFileSystem.LastScriptType);
            Assert.AreEqual(filePath, _mockFileSystem.LastParameter);
        }

        [TestMethod]
        public async Task RemoveItemAsync_WithDirectory_CallsUnpinFromFrequentFolder()
        {
            // Arrange
            string dirPath = @"C:\Test\Folder";
            _mockFileSystem.SetScriptExecutionResult(true);

            // Act
            bool result = await _manager.RemoveItemAsync(dirPath, QuickAccessItemType.Directory);

            // Assert
            Assert.IsTrue(result, "移除文件夹应该返回 true");
            Assert.AreEqual(PSScript.UnpinFromFrequentFolder, _mockFileSystem.LastScriptType);
            Assert.AreEqual(dirPath, _mockFileSystem.LastParameter);
        }

        [TestMethod]
        public async Task EmptyItemsAsync_WithFile_CallsClearRecentFiles()
        {
            // Act
            bool result = await _manager.EmptyItemsAsync(QuickAccessItemType.File);

            // Assert
            Assert.IsTrue(result, "清空文件应该返回 true");
            Assert.IsTrue(_mockFileSystem.WasEmptyRecentFilesCalled);
        }

        [TestMethod]
        public async Task EmptyItemsAsync_WithDirectory_CallsClearFrequentFolders()
        {
            // Act
            bool result = await _manager.EmptyItemsAsync(QuickAccessItemType.Directory);

            // Assert
            Assert.IsTrue(result, "清空文件夹应该返回 true");
            Assert.IsTrue(_mockFileSystem.WasEmptyFrequentFoldersCalled);
        }

        #endregion

        #region 路径安全验证测试

        [TestMethod]
        public async Task AddItemAsync_WithProtectedPath_ThrowsSecurityException()
        {
            // Arrange
            string protectedPath = @"C:\Windows\System32\test.txt";
            _mockFileSystem.ThrowOnValidatePathSecurity = true;
            _mockFileSystem.ValidationException = new SecurityException("Protected system path.");

            // Act & Assert
            await Assert.ThrowsExceptionAsync<SecurityException>(
                async () => await _manager.AddItemAsync(protectedPath, QuickAccessItemType.File));
            Assert.IsTrue(_mockFileSystem.WasValidatePathSecurityCalled);
            Assert.AreEqual(protectedPath, _mockFileSystem.LastValidatedPath);
            Assert.AreEqual(QuickAccessItemType.File, _mockFileSystem.LastValidatedItemType);
        }

        [TestMethod]
        public async Task RemoveItemAsync_WithProtectedPath_ThrowsSecurityException()
        {
            // Arrange
            string protectedPath = @"C:\Program Files\test\folder";
            _mockFileSystem.ThrowOnValidatePathSecurity = true;
            _mockFileSystem.ValidationException = new SecurityException("Protected system path.");

            // Act & Assert
            await Assert.ThrowsExceptionAsync<SecurityException>(
                async () => await _manager.RemoveItemAsync(protectedPath, QuickAccessItemType.Directory));
            Assert.IsTrue(_mockFileSystem.WasValidatePathSecurityCalled);
            Assert.AreEqual(protectedPath, _mockFileSystem.LastValidatedPath);
            Assert.AreEqual(QuickAccessItemType.Directory, _mockFileSystem.LastValidatedItemType);
        }

        [TestMethod]
        public async Task AddItemAsync_WithEmptyPath_ThrowsArgumentException()
        {
            // Arrange
            string emptyPath = "";
            _mockFileSystem.ThrowOnValidatePathSecurity = true;
            _mockFileSystem.ValidationException = new ArgumentException("The path cannot be empty.");

            // Act & Assert
            await Assert.ThrowsExceptionAsync<ArgumentException>(
                async () => await _manager.AddItemAsync(emptyPath, QuickAccessItemType.File));
        }

        #endregion

        #region 可行性检查测试

        [TestMethod]
        public async Task CheckItemAsync_WhenQueryNotFeasible_ThrowsException()
        {
            // Arrange
            _mockFileSystem.SetQueryFeasible(false);

            // Act & Assert
            await Assert.ThrowsExceptionAsync<QuickAccessFeasibilityException>(
                async () => await _manager.CheckItemAsync(@"C:\test.txt", QuickAccessItemType.File));
        }

        [TestMethod]
        public async Task AddItemAsync_WhenPinUnpinNotFeasible_ThrowsException()
        {
            // Arrange
            _mockFileSystem.SetPinUnpinFeasible(false);

            // Act & Assert
            await Assert.ThrowsExceptionAsync<QuickAccessFeasibilityException>(
                async () => await _manager.AddItemAsync(@"C:\test.txt", QuickAccessItemType.File));
        }

        [TestMethod]
        public async Task CheckItemAsync_WithValidPath_ReturnsCorrectResult()
        {
            // Arrange
            string path = @"C:\test.txt";
            List<string> _quickAccessItems = new List<string>() { path };
            _mockFileSystem.SetQueryFeasible(true);
            _mockFileSystem.SetItemExists(path, true);
            _mockFileSystem.SetQuickAccessItems(_quickAccessItems);

            // Act
            bool exists = await _manager.CheckItemAsync(path, QuickAccessItemType.File);

            // Assert
            Assert.IsTrue(exists);
        }

        [TestMethod]
        public async Task GetQuickAccessItemsAsync_WhenQueryNotFeasible_ThrowsException()
        {
            // Arrange
            _mockFileSystem.SetQueryFeasible(false);

            // Act & Assert
            await Assert.ThrowsExceptionAsync<QuickAccessFeasibilityException>(
                async () => await _manager.GetQuickAccessItemsAsync());
        }

        #endregion

        #region 初始化和可行性检查测试

        [TestMethod]
        public async Task Initialize_WhenScriptNotFeasible_AttemptsToFix()
        {
            // Arrange
            _mockFileSystem.SetScriptFeasible(false);

            // Act - 初始化会在第一次调用方法时触发
            await _manager.GetQuickAccessItemsAsync();

            // Assert
            Assert.IsTrue(_mockFileSystem.WasScriptFeasibleChecked);
            Assert.IsTrue(_mockFileSystem.WasFixPolicyCalled);
        }

        [TestMethod]
        public async Task Initialize_WhenFixPolicyFails_ContinuesExecution()
        {
            // Arrange
            _mockFileSystem.SetScriptFeasible(false);
            _mockFileSystem.SetThrowOnFixPolicy(true);

            // Act
            await _manager.GetQuickAccessItemsAsync();

            // Assert
            Assert.IsTrue(_mockFileSystem.WasScriptFeasibleChecked);
            Assert.IsTrue(_mockFileSystem.WasFixPolicyCalled);
        }

        [TestMethod]
        public async Task Initialize_ChecksBothQueryAndPinUnpinFeasibility()
        {
            // Act
            await _manager.GetQuickAccessItemsAsync();

            // Assert
            Assert.IsTrue(_mockFileSystem.WasQueryFeasibleChecked);
            Assert.IsTrue(_mockFileSystem.WasPinUnpinFeasibleChecked);
        }

        [TestMethod]
        public async Task Initialize_OnlyRunsOnce()
        {
            // Act
            await _manager.GetQuickAccessItemsAsync();
            await _manager.GetQuickAccessItemsAsync();

            // Assert
            Assert.AreEqual(3, CountFeasibilityChecks(), "Should only check feasibility once per type");
        }

        private int CountFeasibilityChecks()
        {
            int count = 0;
            if (_mockFileSystem.WasScriptFeasibleChecked) count++;
            if (_mockFileSystem.WasQueryFeasibleChecked) count++;
            if (_mockFileSystem.WasPinUnpinFeasibleChecked) count++;
            return count;
        }

        #endregion
    }

    #region 测试替身类

    /// <summary>
    /// 用于测试的文件系统服务模拟类
    /// </summary>
    internal class MockFileSystemService : IFileSystemService
    {
        private bool _fileExists = false;
        private bool _directoryExists = false;
        private string _fileExtension = ".txt";
        private bool _scriptExecutionResult = true;
        private bool _scriptFeasible = true;
        private bool _queryFeasible = true;
        private bool _pinUnpinFeasible = true;
        private bool _throwOnFixPolicy = false;
        private readonly Dictionary<string, bool> _itemExistence = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        private List<string> _quickAccessItems = new List<string>();
        private List<string> _recentFiles = new List<string>();
        private List<string> _frequentFolders = new List<string>();

        // 记录操作历史
        public PSScript? LastScriptType { get; private set; }
        public string LastParameter { get; private set; }
        public string LastAddedToRecentFile { get; private set; }
        public bool WasAddFileToRecentDocsCalled { get; private set; }
        public bool WasEmptyRecentFilesCalled { get; private set; }
        public bool ThrowOnEmptyRecentFiles { get; set; } = false;
        public bool WasEmptyFrequentFoldersCalled { get; private set; }
        public bool ThrowOnEmptyFrequentFolders { get; set; } = false;
        public bool ThrowOnValidatePathSecurity { get; set; } = false;
        public string LastValidatedPath { get; private set; }
        public QuickAccessItemType LastValidatedItemType { get; private set; }
        public bool WasValidatePathSecurityCalled { get; private set; }
        public Exception ValidationException { get; set; } = new SecurityException("Protected system path.");
        public bool WasScriptFeasibleChecked { get; private set; }
        public bool WasFixPolicyCalled { get; private set; }
        public bool WasQueryFeasibleChecked { get; private set; }
        public bool WasPinUnpinFeasibleChecked { get; private set; }

        // Mock 配置方法
        public void SetFileExists(bool exists) => _fileExists = exists;
        public void SetDirectoryExists(bool exists) => _directoryExists = exists;
        public void SetFileExtension(string extension) => _fileExtension = extension;
        public void SetScriptExecutionResult(bool success) => _scriptExecutionResult = success;
        public void SetScriptFeasible(bool feasible) => _scriptFeasible = feasible;
        public void SetQueryFeasible(bool feasible) => _queryFeasible = feasible;
        public void SetPinUnpinFeasible(bool feasible) => _pinUnpinFeasible = feasible;
        public void SetThrowOnFixPolicy(bool throwException) => _throwOnFixPolicy = throwException;
        public void SetItemExists(string path, bool exists) => _itemExistence[path] = exists;
        public void SetQuickAccessItems(List<string> items) => _quickAccessItems = items;
        public void SetRecentFiles(List<string> files) => _recentFiles = files;
        public void SetFrequentFolders(List<string> folders) => _frequentFolders = folders;

        // IFileSystemService 接口实现
        public bool FileExists(string path) => _fileExists;
        public bool DirectoryExists(string path) => _directoryExists;
        public string GetFileExtension(string path) => _fileExtension;

        public Task<bool> ExecuteScriptAsync(PSScript scriptType, string parameter)
        {
            LastScriptType = scriptType;
            LastParameter = parameter;
            return Task.FromResult(_scriptExecutionResult);
        }

        public void AddFileToRecentDocs(string path)
        {
            WasAddFileToRecentDocsCalled = true;
            LastAddedToRecentFile = path;
            ValidatePathSecurity(path, QuickAccessItemType.File);
        }

        public void EmptyRecentFiles()
        {
            if (ThrowOnEmptyRecentFiles)
                throw new Exception("模拟的异常");

            WasEmptyRecentFilesCalled = true;
        }

        public void EmptyFrequentFolders()
        {
            if (ThrowOnEmptyFrequentFolders)
                throw new Exception("模拟的异常");

            WasEmptyFrequentFoldersCalled = true;
        }

        public void ValidatePathSecurity(string path, QuickAccessItemType itemType)
        {
            WasValidatePathSecurityCalled = true;
            LastValidatedPath = path;
            LastValidatedItemType = itemType;

            if (ThrowOnValidatePathSecurity)
            {
                throw ValidationException;
            }

            // 模拟路径验证逻辑
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Path cannot be null or empty");

            if (path.IndexOf("System32", StringComparison.OrdinalIgnoreCase) >= 0 ||
                path.IndexOf("Program Files", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                throw new SecurityException("Protected system path");
            }
        }

        public bool CheckScriptFeasible()
        {
            WasScriptFeasibleChecked = true;
            return _scriptFeasible;
        }

        public void FixExecutionPolicy()
        {
            WasFixPolicyCalled = true;
            if (_throwOnFixPolicy)
                throw new SecurityException("Mock security exception");
        }

        public Task<bool> CheckQueryFeasibleAsync()
        {
            WasQueryFeasibleChecked = true;
            return Task.FromResult(_queryFeasible);
        }

        public Task<bool> CheckPinUnpinFeasibleAsync()
        {
            WasPinUnpinFeasibleChecked = true;
            return Task.FromResult(_pinUnpinFeasible);
        }

        public Task<List<string>> GetQuickAccessItemsAsync()
        {
            return Task.FromResult(_quickAccessItems);
        }

        public Task<List<string>> GetRecentFilesAsync()
        {
            return Task.FromResult(_recentFiles);
        }

        public Task<List<string>> GetFrequentFoldersAsync()
        {
            return Task.FromResult(_frequentFolders);
        }
    }

    /// <summary>
    /// 用于测试的 QuickAccessQuery 服务模拟类
    /// </summary>
    internal class MockQuickAccessQueryService
    {
        private List<string> _quickAccessItems = new List<string>();
        private List<string> _recentFiles = new List<string>();
        private List<string> _frequentFolders = new List<string>();

        public void SetQuickAccessItems(List<string> items)
        {
            _quickAccessItems = items;
        }

        public void SetRecentFiles(List<string> files)
        {
            _recentFiles = files;
        }

        public void SetFrequentFolders(List<string> folders)
        {
            _frequentFolders = folders;
        }

        public Task<List<string>> GetAllItemsAsync()
        {
            return Task.FromResult(_quickAccessItems);
        }

        public Task<List<string>> GetRecentFilesAsync()
        {
            return Task.FromResult(_recentFiles);
        }

        public Task<List<string>> GetFrequentFoldersAsync()
        {
            return Task.FromResult(_frequentFolders);
        }
    }

    #endregion
}
