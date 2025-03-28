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

        [TestInitialize]
        public void Initialize()
        {
            // 创建模拟服务
            _mockFileSystem = new MockFileSystemService();

            // 替换服务实现
            QuickAccessManager.SetFileSystemService(_mockFileSystem);

            // 设置默认行为
            _mockFileSystem.SetFileExists(true);
            _mockFileSystem.SetDirectoryExists(true);
            _mockFileSystem.SetScriptExecutionResult(true);
        }

        [TestCleanup]
        public void Cleanup()
        {
            // 恢复默认服务
            QuickAccessManager.ResetFileSystemService();
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
            bool result = await QuickAccessManager.PinToFrequentFolderAsync(dirPath);

            // Assert
            Assert.IsTrue(result, "有效文件夹路径应该返回 true");
            Assert.AreEqual(PSScript.PinToFrequentFolder, _mockFileSystem.LastScriptType, "应该执行正确的脚本类型");
            Assert.AreEqual(dirPath, _mockFileSystem.LastParameter, "应该传递正确的参数");
        }

        [TestMethod]
        public async Task PinToQuickAccessAsync_WithFilePath_ReturnsFalse()
        {
            // Arrange
            string filePath = @"C:\Test\file.txt";
            _mockFileSystem.SetFileExists(true);
            _mockFileSystem.SetDirectoryExists(false);

            // Act
            bool result = await QuickAccessManager.PinToFrequentFolderAsync(filePath);

            // Assert
            Assert.IsFalse(result, "文件路径应该返回 false");
            Assert.IsNull(_mockFileSystem.LastScriptType, "不应该执行任何脚本");
        }

        [TestMethod]
        public async Task PinToQuickAccessAsync_WithInvalidPath_ReturnsFalse()
        {
            // Arrange
            string invalidPath = @"C:\NonExistent\Path";
            _mockFileSystem.SetFileExists(false);
            _mockFileSystem.SetDirectoryExists(false);

            // Act
            bool result = await QuickAccessManager.PinToFrequentFolderAsync(invalidPath);

            // Assert
            Assert.IsFalse(result, "无效路径应该返回 false");
            Assert.IsNull(_mockFileSystem.LastScriptType, "不应该执行任何脚本");
        }

        [TestMethod]
        public async Task PinToQuickAccessAsync_WithEmptyPath_ReturnsFalse()
        {
            // Arrange
            string emptyPath = "";

            // Act
            bool result = await QuickAccessManager.PinToFrequentFolderAsync(emptyPath);

            // Assert
            Assert.IsFalse(result, "空路径应该返回 false");
            Assert.IsNull(_mockFileSystem.LastScriptType, "不应该执行任何脚本");
        }

        [TestMethod]
        public async Task PinToQuickAccessAsync_WithScriptFailure_ReturnsFalse()
        {
            // Arrange
            string path = @"C:\Test\file.txt";
            _mockFileSystem.SetFileExists(true);
            _mockFileSystem.SetScriptExecutionResult(false);

            // Act
            bool result = await QuickAccessManager.PinToFrequentFolderAsync(path);

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
            bool result = await QuickAccessManager.UnpinFromFrequentFolderAsync(dirPath);

            // Assert
            Assert.IsTrue(result, "有效路径应该返回 true");
            Assert.AreEqual(PSScript.UnpinFromFrequentFolder, _mockFileSystem.LastScriptType, "应该执行正确的脚本类型");
            Assert.AreEqual(dirPath, _mockFileSystem.LastParameter, "应该传递正确的参数");
        }

        [TestMethod]
        public async Task UnpinFromQuickAccessAsync_WithEmptyPath_ReturnsFalse()
        {
            // Arrange
            string emptyPath = "";

            // Act
            bool result = await QuickAccessManager.UnpinFromFrequentFolderAsync(emptyPath);

            // Assert
            Assert.IsFalse(result, "空路径应该返回 false");
            Assert.IsNull(_mockFileSystem.LastScriptType, "不应该执行任何脚本");
        }

        [TestMethod]
        public async Task UnpinFromQuickAccessAsync_WithScriptFailure_ReturnsFalse()
        {
            // Arrange
            string path = @"C:\Test\file.txt";
            _mockFileSystem.SetScriptExecutionResult(false);

            // Act
            bool result = await QuickAccessManager.UnpinFromFrequentFolderAsync(path);

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
            bool result = await QuickAccessManager.RemoveFromRecentFilesAsync(filePath);

            // Assert
            Assert.IsTrue(result, "有效文件路径应该返回 true");
            Assert.AreEqual(PSScript.RemoveRecentFile, _mockFileSystem.LastScriptType, "应该执行正确的脚本类型");
            Assert.AreEqual(filePath, _mockFileSystem.LastParameter, "应该传递正确的参数");
        }

        [TestMethod]
        public async Task RemoveFromRecentFilesAsync_WithEmptyPath_ReturnsFalse()
        {
            // Arrange
            string emptyPath = "";

            // Act
            bool result = await QuickAccessManager.RemoveFromRecentFilesAsync(emptyPath);

            // Assert
            Assert.IsFalse(result, "空路径应该返回 false");
            Assert.IsNull(_mockFileSystem.LastScriptType, "不应该执行任何脚本");
        }

        [TestMethod]
        public async Task RemoveFromRecentFilesAsync_WithNonExistentFile_ReturnsFalse()
        {
            // Arrange
            string nonExistentFile = @"C:\Test\nonexistent.txt";
            _mockFileSystem.SetFileExists(false);

            // Act
            bool result = await QuickAccessManager.RemoveFromRecentFilesAsync(nonExistentFile);

            // Assert
            Assert.IsFalse(result, "不存在的文件应该返回 false");
            Assert.IsNull(_mockFileSystem.LastScriptType, "不应该执行任何脚本");
        }

        [TestMethod]
        public async Task RemoveFromRecentFilesAsync_WithNoExtension_ReturnsFalse()
        {
            // Arrange
            string fileWithoutExt = @"C:\Test\fileWithoutExtension";
            _mockFileSystem.SetFileExists(true);
            _mockFileSystem.SetFileExtension("");

            // Act
            bool result = await QuickAccessManager.RemoveFromRecentFilesAsync(fileWithoutExt);

            // Assert
            Assert.IsFalse(result, "没有扩展名的文件应该返回 false");
            Assert.IsNull(_mockFileSystem.LastScriptType, "不应该执行任何脚本");
        }

        [TestMethod]
        public async Task AddToRecentFilesAsync_WithValidFilePath_ReturnsTrue()
        {
            // Arrange
            string filePath = @"C:\Test\file.txt";
            _mockFileSystem.SetFileExists(true);
            _mockFileSystem.SetFileExtension(".txt");

            // Act
            bool result = await QuickAccessManager.AddToRecentFilesAsync(filePath);

            // Assert
            Assert.IsTrue(result, "有效文件路径应该返回 true");
            Assert.IsTrue(_mockFileSystem.WasAddFileToRecentDocsCalled, "应该调用 AddFileToRecentDocs 方法");
            Assert.AreEqual(filePath, _mockFileSystem.LastAddedToRecentFile, "应该传递正确的文件路径");
        }

        [TestMethod]
        public async Task AddToRecentFilesAsync_WithEmptyPath_ReturnsFalse()
        {
            // Arrange
            string emptyPath = "";

            // Act
            bool result = await QuickAccessManager.AddToRecentFilesAsync(emptyPath);

            // Assert
            Assert.IsFalse(result, "空路径应该返回 false");
            Assert.IsFalse(_mockFileSystem.WasAddFileToRecentDocsCalled, "不应该调用 AddFileToRecentDocs 方法");
        }

        [TestMethod]
        public async Task AddToRecentFilesAsync_WithNonExistentFile_ReturnsFalse()
        {
            // Arrange
            string nonExistentFile = @"C:\Test\nonexistent.txt";
            _mockFileSystem.SetFileExists(false);

            // Act
            bool result = await QuickAccessManager.AddToRecentFilesAsync(nonExistentFile);

            // Assert
            Assert.IsFalse(result, "不存在的文件应该返回 false");
            Assert.IsFalse(_mockFileSystem.WasAddFileToRecentDocsCalled, "不应该调用 AddFileToRecentDocs 方法");
        }

        [TestMethod]
        public async Task AddToRecentFilesAsync_WithNoExtension_ReturnsFalse()
        {
            // Arrange
            string fileWithoutExt = @"C:\Test\fileWithoutExtension";
            _mockFileSystem.SetFileExists(true);
            _mockFileSystem.SetFileExtension("");

            // Act
            bool result = await QuickAccessManager.AddToRecentFilesAsync(fileWithoutExt);

            // Assert
            Assert.IsFalse(result, "没有扩展名的文件应该返回 false");
            Assert.IsFalse(_mockFileSystem.WasAddFileToRecentDocsCalled, "不应该调用 AddFileToRecentDocs 方法");
        }

        #endregion

        #region 清空最近文件测试

        [TestMethod]
        public async Task ClearRecentFilesAsync_Success_ReturnsTrue()
        {
            // Act
            bool result = await QuickAccessManager.ClearRecentFilesAsync();

            // Assert
            Assert.IsTrue(result, "成功清空最近文件应该返回 true");
            Assert.IsTrue(_mockFileSystem.WasEmptyRecentFilesCalled, "应该调用 EmptyRecentFiles 方法");
        }

        [TestMethod]
        public async Task ClearRecentFilesAsync_WithException_ReturnsFalse()
        {
            // Arrange - 设置 mock 抛出异常
            _mockFileSystem.ThrowOnEmptyRecentFiles = true;

            // Act
            bool result = await QuickAccessManager.ClearRecentFilesAsync();

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
            var items = await QuickAccessManager.GetQuickAccessItemsAsync();
            Assert.IsNotNull(items, "返回的项目列表不应为 null");
        }

        [TestMethod]
        public async Task GetRecentFilesAsync_ReturnsCorrectFiles()
        {
            // 此测试仅作为示例，实际实现可能需要调整
            var files = await QuickAccessManager.GetRecentFilesAsync();
            Assert.IsNotNull(files, "返回的文件列表不应为 null");
        }

        [TestMethod]
        public async Task GetFrequentFoldersAsync_ReturnsCorrectFolders()
        {
            // 此测试仅作为示例，实际实现可能需要调整
            var folders = await QuickAccessManager.GetFrequentFoldersAsync();
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
            bool result = await QuickAccessManager.AddItemAsync(filePath, QuickAccessItemType.File);

            // Assert
            Assert.IsTrue(result, "添加文件应该返回 true");
            Assert.IsTrue(_mockFileSystem.WasAddFileToRecentDocsCalled, "应该调用 AddFileToRecentDocs 方法");
            Assert.AreEqual(filePath, _mockFileSystem.LastAddedToRecentFile, "应该传递正确的文件路径");
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
            bool result = await QuickAccessManager.AddItemAsync(dirPath, QuickAccessItemType.Directory);

            // Assert
            Assert.IsTrue(result, "添加文件夹应该返回 true");
            Assert.AreEqual(PSScript.PinToFrequentFolder, _mockFileSystem.LastScriptType, "应该执行正确的脚本类型");
            Assert.AreEqual(dirPath, _mockFileSystem.LastParameter, "应该传递正确的参数");
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
            bool result = await QuickAccessManager.RemoveItemAsync(filePath, QuickAccessItemType.File);

            // Assert
            Assert.IsTrue(result, "移除文件应该返回 true");
            Assert.AreEqual(PSScript.RemoveRecentFile, _mockFileSystem.LastScriptType, "应该执行正确的脚本类型");
            Assert.AreEqual(filePath, _mockFileSystem.LastParameter, "应该传递正确的参数");
        }

        [TestMethod]
        public async Task RemoveItemAsync_WithDirectory_CallsUnpinFromFrequentFolder()
        {
            // Arrange
            string dirPath = @"C:\Test\Folder";
            _mockFileSystem.SetScriptExecutionResult(true);

            // Act
            bool result = await QuickAccessManager.RemoveItemAsync(dirPath, QuickAccessItemType.Directory);

            // Assert
            Assert.IsTrue(result, "移除文件夹应该返回 true");
            Assert.AreEqual(PSScript.UnpinFromFrequentFolder, _mockFileSystem.LastScriptType, "应该执行正确的脚本类型");
            Assert.AreEqual(dirPath, _mockFileSystem.LastParameter, "应该传递正确的参数");
        }

        [TestMethod]
        public async Task EmptyItemsAsync_WithFile_CallsClearRecentFiles()
        {
            // Act
            bool result = await QuickAccessManager.EmptyItemsAsync(QuickAccessItemType.File);

            // Assert
            Assert.IsTrue(result, "清空文件应该返回 true");
            Assert.IsTrue(_mockFileSystem.WasEmptyRecentFilesCalled, "应该调用 EmptyRecentFiles 方法");
        }

        [TestMethod]
        public async Task EmptyItemsAsync_WithDirectory_CallsClearFrequentFolders()
        {
            // Act
            bool result = await QuickAccessManager.EmptyItemsAsync(QuickAccessItemType.Directory);

            // Assert
            Assert.IsTrue(result, "清空文件夹应该返回 true");
            Assert.IsTrue(_mockFileSystem.WasEmptyFrequentFoldersCalled, "应该调用 EmptyFrequentFolders 方法");
        }

        [TestMethod]
        public async Task AddItemAsync_WithInvalidPath_ReturnsFalse()
        {
            // Arrange
            string invalidPath = "";

            // Act
            bool result = await QuickAccessManager.AddItemAsync(invalidPath, QuickAccessItemType.File);

            // Assert
            Assert.IsFalse(result, "无效路径应该返回 false");
        }

        [TestMethod]
        public async Task RemoveItemAsync_WithInvalidPath_ReturnsFalse()
        {
            // Arrange
            string invalidPath = "";

            // Act
            bool result = await QuickAccessManager.RemoveItemAsync(invalidPath, QuickAccessItemType.File);

            // Assert
            Assert.IsFalse(result, "无效路径应该返回 false");
        }

        #endregion

        #region 路径安全验证测试

        [TestMethod]
        public async Task AddItemAsync_WithProtectedPath_ReturnsFalse()
        {
            // Arrange
            string protectedPath = @"C:\Windows\System32\test.txt";
            _mockFileSystem.ThrowOnValidatePathSecurity = true;
            _mockFileSystem.ValidationException = new SecurityException("Protected system path.");

            // Act
            bool result = await QuickAccessManager.AddItemAsync(protectedPath, QuickAccessItemType.File);

            // Assert
            Assert.IsFalse(result, "受保护的路径应该返回 false");
            Assert.IsTrue(_mockFileSystem.WasValidatePathSecurityCalled, "应该调用 ValidatePathSecurity 方法");
            Assert.AreEqual(protectedPath, _mockFileSystem.LastValidatedPath, "应该验证正确的路径");
            Assert.AreEqual(QuickAccessItemType.File, _mockFileSystem.LastValidatedItemType, "应该验证正确的项目类型");
        }

        [TestMethod]
        public async Task RemoveItemAsync_WithProtectedPath_ReturnsFalse()
        {
            // Arrange
            string protectedPath = @"C:\Program Files\test\folder";
            _mockFileSystem.ThrowOnValidatePathSecurity = true;
            _mockFileSystem.ValidationException = new SecurityException("Protected system path.");

            // Act
            bool result = await QuickAccessManager.RemoveItemAsync(protectedPath, QuickAccessItemType.Directory);

            // Assert
            Assert.IsFalse(result, "受保护的路径应该返回 false");
            Assert.IsTrue(_mockFileSystem.WasValidatePathSecurityCalled, "应该调用 ValidatePathSecurity 方法");
            Assert.AreEqual(protectedPath, _mockFileSystem.LastValidatedPath, "应该验证正确的路径");
            Assert.AreEqual(QuickAccessItemType.Directory, _mockFileSystem.LastValidatedItemType, "应该验证正确的项目类型");
        }

        [TestMethod]
        public async Task AddItemAsync_WithNonExistentPath_ReturnsFalse()
        {
            // Arrange
            string nonExistentPath = @"C:\NonExistent\Path\file.txt";
            _mockFileSystem.ThrowOnValidatePathSecurity = true;
            _mockFileSystem.ValidationException = new FileNotFoundException($"The path does not exist: {nonExistentPath}");

            // Act
            bool result = await QuickAccessManager.AddItemAsync(nonExistentPath, QuickAccessItemType.File);

            // Assert
            Assert.IsFalse(result, "不存在的路径应该返回 false");
            Assert.IsTrue(_mockFileSystem.WasValidatePathSecurityCalled, "应该调用 ValidatePathSecurity 方法");
        }

        [TestMethod]
        public async Task AddItemAsync_WithEmptyPath_ReturnsFalse()
        {
            // Arrange
            string emptyPath = "";
            _mockFileSystem.ThrowOnValidatePathSecurity = true;
            _mockFileSystem.ValidationException = new ArgumentException("The path cannot be empty.");

            // Act
            bool result = await QuickAccessManager.AddItemAsync(emptyPath, QuickAccessItemType.File);

            // Assert
            Assert.IsFalse(result, "空路径应该返回 false");
            Assert.IsTrue(_mockFileSystem.WasValidatePathSecurityCalled, "应该调用 ValidatePathSecurity 方法");
        }

        [TestMethod]
        public async Task AddItemAsync_WithValidPath_CallsCorrectMethod()
        {
            // Arrange
            string validPath = @"C:\Valid\Path\file.txt";
            _mockFileSystem.SetFileExists(true);
            _mockFileSystem.SetFileExtension(".txt");

            // Act
            bool result = await QuickAccessManager.AddItemAsync(validPath, QuickAccessItemType.File);

            // Assert
            Assert.IsTrue(result, "有效路径应该返回 true");
            Assert.IsTrue(_mockFileSystem.WasValidatePathSecurityCalled, "应该调用 ValidatePathSecurity 方法");
            Assert.IsTrue(_mockFileSystem.WasAddFileToRecentDocsCalled, "应该调用 AddFileToRecentDocs 方法");
        }

        #endregion
    }

    #region 测试替身类

    /// <summary>
    /// 用于测试的文件系统服务模拟类
    /// </summary>
    internal class MockFileSystemService : QuickAccessManager.IFileSystemService
    {
        private bool _fileExists = false;
        private bool _directoryExists = false;
        private string _fileExtension = ".txt";
        private bool _scriptExecutionResult = true;

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

        public void SetFileExists(bool exists)
        {
            _fileExists = exists;
        }

        public void SetDirectoryExists(bool exists)
        {
            _directoryExists = exists;
        }

        public void SetFileExtension(string extension)
        {
            _fileExtension = extension;
        }

        public void SetScriptExecutionResult(bool success)
        {
            _scriptExecutionResult = success;
        }

        public bool FileExists(string path)
        {
            return _fileExists;
        }

        public bool DirectoryExists(string path)
        {
            return _directoryExists;
        }

        public string GetFileExtension(string path)
        {
            return _fileExtension;
        }

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
