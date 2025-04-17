using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Wincent;

namespace TestWincent
{
    /// <summary>
    /// 用于测试的模拟文件系统实现
    /// </summary>
    public class MockFileSystem : IFileSystem
    {
        private readonly Dictionary<string, bool> _fileExistsResults = new Dictionary<string, bool>();
        private readonly Dictionary<string, DateTime> _fileTimestamps = new Dictionary<string, DateTime>();
        private readonly List<string> _deletedFiles = new List<string>();
        private Action<string> _onDeleteFile;

        public MockFileSystem()
        {
            // 默认文件不存在
            FileExistsDefault = false;
            // 默认时间戳
            LastWriteTimeDefault = DateTime.Now;
        }

        // 默认的文件存在状态
        public bool FileExistsDefault { get; set; }

        // 默认的文件时间戳
        public DateTime LastWriteTimeDefault { get; set; }

        // 模拟文件存在状态
        public void SetFileExists(string path, bool exists)
        {
            _fileExistsResults[path] = exists;
        }

        // 模拟文件时间戳
        public void SetLastWriteTime(string path, DateTime timestamp)
        {
            _fileTimestamps[path] = timestamp;
        }

        // 设置删除文件时的回调
        public void SetDeleteFileCallback(Action<string> callback)
        {
            _onDeleteFile = callback;
        }

        // 获取被删除的文件列表
        public IReadOnlyList<string> DeletedFiles => _deletedFiles;

        // 实现 IFileSystem 接口
        public bool FileExists(string path)
        {
            return _fileExistsResults.ContainsKey(path) ? _fileExistsResults[path] : FileExistsDefault;
        }

        public void DeleteFile(string path)
        {
            _deletedFiles.Add(path);
            _onDeleteFile?.Invoke(path);
        }

        public DateTime GetLastWriteTime(string path)
        {
            return _fileTimestamps.ContainsKey(path) ? _fileTimestamps[path] : LastWriteTimeDefault;
        }
    }

    [TestClass]
    public class QuickAccessDataFilesTests
    {
        [TestMethod]
        public void RemoveRecentFile_FileExists_DeletesFile()
        {
            // Arrange
            var mockFileSystem = new MockFileSystem();
            mockFileSystem.FileExistsDefault = true;

            var quickAccess = new QuickAccessDataFiles(mockFileSystem);
            string recentFilesPath = quickAccess.RecentFilesPath;

            bool fileDeleted = false;
            mockFileSystem.SetDeleteFileCallback(path => {
                if (path == recentFilesPath)
                    fileDeleted = true;
            });

            // Act
            quickAccess.RemoveRecentFile();

            // Assert
            Assert.IsTrue(fileDeleted, "文件应该被删除");
            Assert.IsTrue(mockFileSystem.DeletedFiles.Contains(recentFilesPath), "最近访问文件应该在被删除文件列表中");
        }

        [TestMethod]
        public void RemoveRecentFile_FileDoesNotExist_DoesNotDeleteFile()
        {
            // Arrange
            var mockFileSystem = new MockFileSystem();
            mockFileSystem.FileExistsDefault = false;

            var quickAccess = new QuickAccessDataFiles(mockFileSystem);

            // Act
            quickAccess.RemoveRecentFile();

            // Assert
            Assert.AreEqual(0, mockFileSystem.DeletedFiles.Count, "不存在的文件不应该被删除");
        }

        [TestMethod]
        [ExpectedException(typeof(IOException))]
        public void RemoveRecentFile_DeleteThrowsException_PropagatesException()
        {
            // Arrange
            var mockFileSystem = new MockFileSystem();
            mockFileSystem.FileExistsDefault = true;
            mockFileSystem.SetDeleteFileCallback(_ => throw new IOException("测试异常"));

            var quickAccess = new QuickAccessDataFiles(mockFileSystem);

            // Act - 应该抛出异常
            quickAccess.RemoveRecentFile();
        }

        [TestMethod]
        public void GetRecentFilesModifiedTime_FileExists_ReturnsLastWriteTime()
        {
            // Arrange
            var expectedTime = new DateTime(2023, 5, 15);
            var mockFileSystem = new MockFileSystem();
            mockFileSystem.FileExistsDefault = true;

            var quickAccess = new QuickAccessDataFiles(mockFileSystem);
            string recentFilesPath = quickAccess.RecentFilesPath;

            mockFileSystem.SetLastWriteTime(recentFilesPath, expectedTime);

            // Act
            var result = quickAccess.GetRecentFilesModifiedTime();

            // Assert
            Assert.AreEqual(expectedTime, result);
        }

        [TestMethod]
        public void GetFrequentFoldersModifiedTime_FileExists_ReturnsLastWriteTime()
        {
            // Arrange
            var expectedTime = new DateTime(2023, 6, 20);
            var mockFileSystem = new MockFileSystem();
            mockFileSystem.FileExistsDefault = true;

            var quickAccess = new QuickAccessDataFiles(mockFileSystem);
            string frequentFoldersPath = quickAccess.FrequentFoldersPath;

            mockFileSystem.SetLastWriteTime(frequentFoldersPath, expectedTime);

            // Act
            var result = quickAccess.GetFrequentFoldersModifiedTime();

            // Assert
            Assert.AreEqual(expectedTime, result);
        }

        [TestMethod]
        public void GetQuickAccessModifiedTime_ReturnsMostRecentTime()
        {
            // Arrange
            var recentTime = new DateTime(2023, 5, 15);
            var frequentTime = new DateTime(2023, 6, 20); // 更近的时间

            var mockFileSystem = new MockFileSystem();
            mockFileSystem.FileExistsDefault = true;

            var quickAccess = new QuickAccessDataFiles(mockFileSystem);

            mockFileSystem.SetLastWriteTime(quickAccess.RecentFilesPath, recentTime);
            mockFileSystem.SetLastWriteTime(quickAccess.FrequentFoldersPath, frequentTime);

            // Act
            var result = quickAccess.GetQuickAccessModifiedTime();

            // Assert - 应该返回最新的时间
            Assert.AreEqual(frequentTime, result);
        }

        [TestMethod]
        public void GetModifiedTimeForScript_QueryRecentFile_ReturnsRecentFileTime()
        {
            // Arrange
            var expectedTime = new DateTime(2023, 5, 15);
            var mockFileSystem = new MockFileSystem();
            mockFileSystem.FileExistsDefault = true;

            var quickAccess = new QuickAccessDataFiles(mockFileSystem);
            mockFileSystem.SetLastWriteTime(quickAccess.RecentFilesPath, expectedTime);

            // Act
            var result = quickAccess.GetModifiedTimeForScript(PSScript.QueryRecentFile);

            // Assert
            Assert.AreEqual(expectedTime, result);
        }

        [TestMethod]
        public void GetModifiedTimeForScript_QueryFrequentFolder_ReturnsFrequentFolderTime()
        {
            // Arrange
            var expectedTime = new DateTime(2023, 6, 20);
            var mockFileSystem = new MockFileSystem();
            mockFileSystem.FileExistsDefault = true;

            var quickAccess = new QuickAccessDataFiles(mockFileSystem);
            mockFileSystem.SetLastWriteTime(quickAccess.FrequentFoldersPath, expectedTime);

            // Act
            var result = quickAccess.GetModifiedTimeForScript(PSScript.QueryFrequentFolder);

            // Assert
            Assert.AreEqual(expectedTime, result);
        }

        [TestMethod]
        public void GetModifiedTimeForScript_NonQueryScript_ReturnsCurrentTime()
        {
            // Arrange
            var mockFileSystem = new MockFileSystem();
            var quickAccess = new QuickAccessDataFiles(mockFileSystem);

            // Act
            var result = quickAccess.GetModifiedTimeForScript(PSScript.RefreshExplorer);

            // Assert - 非查询脚本应返回接近当前时间的值
            var now = DateTime.Now;
            var timeDifference = (now - result).TotalSeconds;
            Assert.IsTrue(timeDifference < 5); // 允许5秒的误差
        }

        [TestMethod]
        public void GetRecentFilesModifiedTime_FileDoesNotExist_ReturnsCurrentTime()
        {
            // Arrange
            var mockFileSystem = new MockFileSystem();
            mockFileSystem.FileExistsDefault = false;

            var quickAccess = new QuickAccessDataFiles(mockFileSystem);

            // Act
            var result = quickAccess.GetRecentFilesModifiedTime();

            // Assert
            var now = DateTime.Now;
            var timeDifference = (now - result).TotalSeconds;
            Assert.IsTrue(timeDifference < 5); // 允许5秒的误差
        }

        [TestMethod]
        public void GetRecentFilesModifiedTime_ThrowsException_ReturnsCurrentTime()
        {
            // Arrange
            var mockFileSystem = new MockFileSystem();
            mockFileSystem.FileExistsDefault = true;

            // 设置获取时间时抛出异常
            var quickAccess = new QuickAccessDataFiles(new ThrowingFileSystem());

            // Act
            var result = quickAccess.GetRecentFilesModifiedTime();

            // Assert
            var now = DateTime.Now;
            var timeDifference = (now - result).TotalSeconds;
            Assert.IsTrue(timeDifference < 5); // 允许5秒的误差
        }
    }

    /// <summary>
    /// 总是抛出异常的文件系统实现
    /// </summary>
    public class ThrowingFileSystem : IFileSystem
    {
        public bool FileExists(string path)
        {
            return true; // 文件总是存在
        }

        public void DeleteFile(string path)
        {
            throw new IOException("测试删除异常");
        }

        public DateTime GetLastWriteTime(string path)
        {
            throw new IOException("测试获取时间异常");
        }
    }
}
