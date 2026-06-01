using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Wincent;

namespace TestWincent
{
    /// <summary>
    /// Mock file system implementation for testing.
    /// </summary>
    public class MockFileSystem : IFileSystem
    {
        private readonly Dictionary<string, bool> _fileExistsResults = new Dictionary<string, bool>();
        private readonly Dictionary<string, DateTime> _fileTimestamps = new Dictionary<string, DateTime>();
        private readonly List<string> _deletedFiles = new List<string>();
        private Action<string> _onDeleteFile;

        public MockFileSystem()
        {
            // Default: file does not exist
            FileExistsDefault = false;
            // Default timestamp
            LastWriteTimeDefault = DateTime.Now;
        }

        // Default file existence state
        public bool FileExistsDefault { get; set; }

        // Default file timestamp
        public DateTime LastWriteTimeDefault { get; set; }

        // Simulate file existence
        public void SetFileExists(string path, bool exists)
        {
            _fileExistsResults[path] = exists;
        }

        // Simulate file timestamp
        public void SetLastWriteTime(string path, DateTime timestamp)
        {
            _fileTimestamps[path] = timestamp;
        }

        // Set delete file callback
        public void SetDeleteFileCallback(Action<string> callback)
        {
            _onDeleteFile = callback;
        }

        // Get list of deleted files
        public IReadOnlyList<string> DeletedFiles => _deletedFiles;

        // IFileSystem implementation
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
            Assert.IsTrue(fileDeleted, "File should be deleted");
            Assert.IsTrue(mockFileSystem.DeletedFiles.Contains(recentFilesPath), "Recent files should be in deleted files list");
        }

        [TestMethod]
        public void Constructor_UsesInjectedRecentFolderProvider()
        {
            var mockFileSystem = new MockFileSystem();
            var quickAccess = new QuickAccessDataFiles(
                mockFileSystem,
                new StubRecentFolder(@"C:\InjectedRecent"));

            Assert.AreEqual(
                @"C:\InjectedRecent\AutomaticDestinations\5f7b5f1e01b83767.automaticDestinations-ms",
                quickAccess.RecentFilesPath);
            Assert.AreEqual(
                @"C:\InjectedRecent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms",
                quickAccess.FrequentFoldersPath);
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
            Assert.AreEqual(0, mockFileSystem.DeletedFiles.Count, "Non-existent files should not be deleted");
        }

        [TestMethod]
        [ExpectedException(typeof(IOException))]
        public void RemoveRecentFile_DeleteThrowsException_PropagatesException()
        {
            // Arrange
            var mockFileSystem = new MockFileSystem();
            mockFileSystem.FileExistsDefault = true;
            mockFileSystem.SetDeleteFileCallback(_ => throw new IOException("Test exception"));

            var quickAccess = new QuickAccessDataFiles(mockFileSystem);

            // Act
            quickAccess.RemoveRecentFile();

            // Assert - ExpectedException handles verification
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
            var frequentTime = new DateTime(2023, 6, 20); // more recent time

            var mockFileSystem = new MockFileSystem();
            mockFileSystem.FileExistsDefault = true;

            var quickAccess = new QuickAccessDataFiles(mockFileSystem);

            mockFileSystem.SetLastWriteTime(quickAccess.RecentFilesPath, recentTime);
            mockFileSystem.SetLastWriteTime(quickAccess.FrequentFoldersPath, frequentTime);

            // Act
            var result = quickAccess.GetQuickAccessModifiedTime();

            // Assert - should return the most recent time
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

            // Assert - non-query scripts should return a value close to current time
            var now = DateTime.Now;
            var timeDifference = (now - result).TotalSeconds;
            Assert.IsTrue(timeDifference < 5); // allow 5-second tolerance
        }

        [TestMethod]
        public void GetModifiedTimeForScript_QueryScript_FileMissing_ThrowsFileNotFoundException()
        {
            var mockFileSystem = new MockFileSystem();
            mockFileSystem.FileExistsDefault = false;
            var quickAccess = new QuickAccessDataFiles(mockFileSystem);

            Assert.ThrowsException<FileNotFoundException>(
                () => quickAccess.GetModifiedTimeForScript(PSScript.QueryRecentFile));
            Assert.ThrowsException<FileNotFoundException>(
                () => quickAccess.GetModifiedTimeForScript(PSScript.QueryFrequentFolder));
            Assert.ThrowsException<FileNotFoundException>(
                () => quickAccess.GetModifiedTimeForScript(PSScript.QueryQuickAccess));
        }

        [TestMethod]
        [ExpectedException(typeof(FileNotFoundException))]
        public void GetRecentFilesModifiedTime_FileDoesNotExist_ThrowsException()
        {
            // Arrange
            var mockFileSystem = new MockFileSystem();
            mockFileSystem.FileExistsDefault = false;

            var quickAccess = new QuickAccessDataFiles(mockFileSystem);

            // Act - should throw exception
            var result = quickAccess.GetRecentFilesModifiedTime();

            // Assert - verified via ExpectedException attribute
        }

        [TestMethod]
        [ExpectedException(typeof(IOException))]
        public void GetRecentFilesModifiedTime_GetLastWriteTimeThrows_ThrowsIOException()
        {
            // Arrange: use a file system that throws when getting the last write time
            var quickAccess = new QuickAccessDataFiles(new ThrowingFileSystem());

            // Act - should throw IOException wrapping the underlying error
            var result = quickAccess.GetRecentFilesModifiedTime();

            // Assert - verified via ExpectedException attribute
        }

        [TestMethod]
        [ExpectedException(typeof(FileNotFoundException))]
        public void GetFrequentFoldersModifiedTime_FileDoesNotExist_ThrowsException()
        {
            // Arrange
            var mockFileSystem = new MockFileSystem();
            mockFileSystem.FileExistsDefault = false;

            var quickAccess = new QuickAccessDataFiles(mockFileSystem);

            // Act - should throw exception
            var result = quickAccess.GetFrequentFoldersModifiedTime();

            // Assert - verified via ExpectedException attribute
        }

        [TestMethod]
        public void GetQuickAccessModifiedTime_OnlyRecentFileExists_ReturnsRecentFileTime()
        {
            // Arrange
            var expectedTime = new DateTime(2023, 5, 15);
            var mockFileSystem = new MockFileSystem();

            // Set recent file exists but frequent folder does not exist
            mockFileSystem.SetFileExists("any", false); // default to not exist
            var quickAccess = new QuickAccessDataFiles(mockFileSystem);
            mockFileSystem.SetFileExists(quickAccess.RecentFilesPath, true); // set recent file exists
            mockFileSystem.SetLastWriteTime(quickAccess.RecentFilesPath, expectedTime);

            // Act
            var result = quickAccess.GetQuickAccessModifiedTime();

            // Assert
            Assert.AreEqual(expectedTime, result);
        }

        [TestMethod]
        public void GetQuickAccessModifiedTime_OnlyFrequentFolderExists_ReturnsFrequentFolderTime()
        {
            // Arrange
            var expectedTime = new DateTime(2023, 6, 20);
            var mockFileSystem = new MockFileSystem();

            // Set frequent folder exists but recent file does not exist
            mockFileSystem.SetFileExists("any", false); // default to not exist
            var quickAccess = new QuickAccessDataFiles(mockFileSystem);
            mockFileSystem.SetFileExists(quickAccess.FrequentFoldersPath, true); // set frequent folder exists
            mockFileSystem.SetLastWriteTime(quickAccess.FrequentFoldersPath, expectedTime);

            // Act
            var result = quickAccess.GetQuickAccessModifiedTime();

            // Assert
            Assert.AreEqual(expectedTime, result);
        }

        [TestMethod]
        [ExpectedException(typeof(FileNotFoundException))]
        public void GetQuickAccessModifiedTime_NeitherExists_ThrowsException()
        {
            // Arrange
            var mockFileSystem = new MockFileSystem();
            mockFileSystem.FileExistsDefault = false; // set all files to not exist

            var quickAccess = new QuickAccessDataFiles(mockFileSystem);

            // Act - should throw exception
            var result = quickAccess.GetQuickAccessModifiedTime();

            // Assert - verified via ExpectedException attribute
        }
    }

    /// <summary>
    /// File system implementation that always throws exceptions.
    /// </summary>
    public class ThrowingFileSystem : IFileSystem
    {
        public bool FileExists(string path)
        {
            return true; // file always exists
        }

        public void DeleteFile(string path)
        {
            throw new IOException("Test delete exception");
        }

        public DateTime GetLastWriteTime(string path)
        {
            throw new IOException("Test get time exception");
        }
    }

    internal sealed class StubRecentFolder : IWindowsRecentFolder
    {
        private readonly string _path;

        public StubRecentFolder(string path)
        {
            _path = path;
        }

        public string GetPath()
        {
            return _path;
        }
    }
}
