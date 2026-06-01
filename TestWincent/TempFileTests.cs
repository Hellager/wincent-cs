using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class TestTempFile
    {
        private const string TestDirName = "WincentTempTest";

        [TestInitialize]
        public void Initialize()
        {
            DeleteDirectoryIfExists(Path.Combine(Path.GetTempPath(), TestDirName));
        }

        [TestCleanup]
        public void Cleanup()
        {
            DeleteDirectoryIfExists(Path.Combine(Path.GetTempPath(), TestDirName));
        }

        #region Basic Functionality Tests

        [TestMethod]
        public void Create_WithStringContent_CreatesFile()
        {
            // Arrange
            const string content = "Test content";

            // Act
            using (var tempFile = TempFile.Create(content, "txt", directoryName: TestDirName))
            {
                // Assert
                Assert.IsTrue(File.Exists(tempFile.FullPath), "File should have been created");
                Assert.AreEqual(content, File.ReadAllText(tempFile.FullPath), "File content should match");
                Assert.AreEqual(".txt", Path.GetExtension(tempFile.FullPath), "File extension should match");
            }
        }

        [TestMethod]
        public void Create_WithBinaryContent_CreatesFile()
        {
            // Arrange
            byte[] content = Encoding.UTF8.GetBytes("Binary test content");

            // Act
            using (var tempFile = TempFile.Create(content, "bin", directoryName: TestDirName))
            {
                // Assert
                Assert.IsTrue(File.Exists(tempFile.FullPath), "File should have been created");
                CollectionAssert.AreEqual(content, File.ReadAllBytes(tempFile.FullPath), "Binary content should match");
                Assert.AreEqual(".bin", Path.GetExtension(tempFile.FullPath), "File extension should match");
            }
        }

        [TestMethod]
        public void Create_WithCustomTextEncoding_CreatesFile()
        {
            // Arrange
            const string content = "测试内容";

            // Act
            using (var tempFile = TempFile.Create(content, "txt", Encoding.Unicode, TestDirName))
            {
                // Assert
                Assert.IsTrue(File.Exists(tempFile.FullPath), "File should have been created");
                Assert.AreEqual(content, File.ReadAllText(tempFile.FullPath, Encoding.Unicode), "Encoded content should match");
            }
        }

        [TestMethod]
        public void Dispose_DeletesFile()
        {
            // Arrange
            string filePath;

            // Act
            using (var tempFile = TempFile.Create("Test content", "txt", directoryName: TestDirName))
            {
                filePath = tempFile.FullPath;
                Assert.IsTrue(File.Exists(filePath), "File should exist while in use");
            }

            // Assert
            Assert.IsFalse(File.Exists(filePath), "File should be deleted after Dispose");
        }

        #endregion

        #region Property Tests

        [TestMethod]
        public void FullPath_ReturnsCorrectPath()
        {
            // Arrange & Act
            using (var tempFile = TempFile.Create("Test", "txt", directoryName: TestDirName))
            {
                // Assert
                Assert.IsTrue(tempFile.FullPath.StartsWith(Path.GetTempPath()), "Path should start with temp directory");
                Assert.IsTrue(tempFile.FullPath.Contains(TestDirName), "Path should contain custom directory name");
                Assert.IsTrue(tempFile.FullPath.EndsWith(".txt"), "Path should end with correct extension");
            }
        }

        [TestMethod]
        public void FileName_ReturnsCorrectName()
        {
            // Arrange & Act
            using (var tempFile = TempFile.Create("Test", "txt", directoryName: TestDirName))
            {
                // Assert
                string fileName = tempFile.FileName;
                Assert.IsTrue(fileName.EndsWith(".txt"), "File name should contain extension");
                Assert.AreEqual(Path.GetFileName(tempFile.FullPath), fileName, "FileName should return file name part");
            }
        }

        [TestMethod]
        public void Create_WithDirectoryName_AffectsOnlyCurrentFile()
        {
            // Arrange
            const string customDirName = "CustomTempDir";

            try
            {
                // Act
                using (var customTempFile = TempFile.Create("Test", "txt", directoryName: customDirName))
                using (var defaultTempFile = TempFile.Create("Test", "txt", directoryName: TestDirName))
                {
                    // Assert
                    Assert.IsTrue(customTempFile.FullPath.Contains(customDirName), "Custom directory name should be applied");
                    Assert.IsTrue(defaultTempFile.FullPath.Contains(TestDirName), "Subsequent call should use its own directory name");
                    Assert.IsFalse(defaultTempFile.FullPath.Contains(customDirName), "Custom directory name should not leak to subsequent calls");
                }
            }
            finally
            {
                DeleteDirectoryIfExists(Path.Combine(Path.GetTempPath(), customDirName));
            }
        }

        #endregion

        #region Exception Handling Tests

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Create_WithNullStringContent_ThrowsException()
        {
            // Act
            TempFile.Create((string)null, "txt", directoryName: TestDirName);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Create_WithNullBinaryContent_ThrowsException()
        {
            // Act
            TempFile.Create((byte[])null, "bin", directoryName: TestDirName);
        }

        [TestMethod]
        public void Create_WithNullEncoding_UsesDefaultEncoding()
        {
            // Arrange
            const string content = "Test";

            // Act
            using (var tempFile = TempFile.Create(content, "txt", null, TestDirName))
            {
                // Assert - DefaultTextEncoding is UTF-8 without BOM
                Assert.IsTrue(File.Exists(tempFile.FullPath), "File should have been created");
                byte[] bytes = File.ReadAllBytes(tempFile.FullPath);
                Assert.AreEqual(content, Encoding.UTF8.GetString(bytes), "Should use default encoding");
                Assert.IsFalse(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF,
                    "Should not contain BOM");
            }
        }

        [TestMethod]
        public void Create_WithNullExtension_UsesDefaultExtension()
        {
            // Act
            using (var tempFile = TempFile.Create("Test", null, directoryName: TestDirName))
            {
                // Assert
                Assert.AreEqual(".ps1", Path.GetExtension(tempFile.FullPath), "Should use default extension");
            }
        }

        [TestMethod]
        public void Create_WithEmptyExtension_UsesDefaultExtension()
        {
            // Act
            using (var tempFile = TempFile.Create("Test", string.Empty, directoryName: TestDirName))
            {
                // Assert
                Assert.AreEqual(".ps1", Path.GetExtension(tempFile.FullPath), "Should use default extension");
            }
        }

        [TestMethod]
        public void Create_WithExtensionWithoutDot_AddsDot()
        {
            // Act
            using (var tempFile = TempFile.Create("Test", "ext", directoryName: TestDirName))
            {
                // Assert
                Assert.AreEqual(".ext", Path.GetExtension(tempFile.FullPath), "Should add dot");
            }
        }

        [TestMethod]
        public void Create_WithExtensionWithDot_PreservesDot()
        {
            // Act
            using (var tempFile = TempFile.Create("Test", ".ext", directoryName: TestDirName))
            {
                // Assert
                Assert.AreEqual(".ext", Path.GetExtension(tempFile.FullPath), "Should preserve dot");
            }
        }

        [TestMethod]
        public void ReadAllText_AfterDispose_ThrowsObjectDisposedException()
        {
            // Arrange
            TempFile tempFile = TempFile.Create("Test", "txt", directoryName: TestDirName);
            tempFile.Dispose();

            // Act & Assert
            Assert.ThrowsException<ObjectDisposedException>(() => tempFile.ReadAllText());
        }

        [TestMethod]
        public void OpenRead_AfterDispose_ThrowsObjectDisposedException()
        {
            // Arrange
            TempFile tempFile = TempFile.Create("Test", "txt", directoryName: TestDirName);
            tempFile.Dispose();

            // Act & Assert
            Assert.ThrowsException<ObjectDisposedException>(() => tempFile.OpenRead());
        }

        [TestMethod]
        public void Dispose_CalledMultipleTimes_IsSafe()
        {
            // Arrange
            var tempFile = TempFile.Create("Test", "txt", directoryName: TestDirName);

            // Act
            tempFile.Dispose();
            tempFile.Dispose();

            // Assert
            Assert.IsFalse(File.Exists(tempFile.FullPath), "File should remain deleted after repeated Dispose");
        }

        [TestMethod]
        public void Dispose_WhenDeleteFails_DoesNotThrow()
        {
            // Arrange
            TempFile tempFile = TempFile.Create("Test", "txt", directoryName: TestDirName);
            string filePath = tempFile.FullPath;

            using (new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                // Act & Assert
                tempFile.Dispose();
            }

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }

        [TestMethod]
        public void Create_WhenCustomDirectoryCreationFails_FallsBackToSystemTemp()
        {
            // Arrange
            string blockingName = "WincentTempBlockingFile";
            string blockingPath = Path.Combine(Path.GetTempPath(), blockingName);

            File.WriteAllText(blockingPath, "blocks directory creation");

            try
            {
                // Act
                using (var tempFile = TempFile.Create("Test", "txt", directoryName: blockingName))
                {
                    // Assert
                    string expectedDirectory = NormalizeDirectory(Path.GetTempPath());
                    string actualDirectory = NormalizeDirectory(Path.GetDirectoryName(tempFile.FullPath));
                    Assert.AreEqual(expectedDirectory, actualDirectory, "Should fall back to system temp when directory creation fails");
                    Assert.IsTrue(File.Exists(tempFile.FullPath), "File should still be created after fallback");
                }
            }
            finally
            {
                if (File.Exists(blockingPath))
                {
                    File.Delete(blockingPath);
                }
            }
        }

        #endregion

        #region File Operation Tests

        [TestMethod]
        public void ReadAllText_ReturnsCorrectContent()
        {
            // Arrange
            const string content = "Test content for reading";

            // Act
            using (var tempFile = TempFile.Create(content, "txt", directoryName: TestDirName))
            {
                // Assert
                Assert.AreEqual(content, tempFile.ReadAllText(), "ReadAllText should return correct content");
            }
        }

        [TestMethod]
        public void ReadAllText_WithEncoding_ReturnsCorrectContent()
        {
            // Arrange
            const string content = "测试内容";

            // Act
            using (var tempFile = TempFile.Create(content, "txt", Encoding.Unicode, TestDirName))
            {
                // Assert
                Assert.AreEqual(content, tempFile.ReadAllText(Encoding.Unicode), "Should return correct content with specified encoding");
            }
        }

        [TestMethod]
        public void OpenRead_ReturnsValidStream()
        {
            // Arrange
            const string content = "Test content for stream";

            // Act
            using (var tempFile = TempFile.Create(content, "txt", directoryName: TestDirName))
            {
                using (var stream = tempFile.OpenRead())
                {
                    // Assert
                    Assert.IsTrue(stream.CanRead, "Stream should be readable");
                    Assert.IsFalse(stream.CanWrite, "Stream should not be writable");

                    byte[] buffer = new byte[100];
                    int bytesRead = stream.Read(buffer, 0, buffer.Length);
                    string readContent = Encoding.UTF8.GetString(buffer, 0, bytesRead);

                    Assert.AreEqual(content, readContent, "Stream should contain correct content");
                }
            }
        }

        #endregion

        #region Concurrency and Performance Tests

        [TestMethod]
        public void Create_MultipleTempFiles_CreatesUniqueFiles()
        {
            // Arrange
            const int fileCount = 10;
            var files = new TempFile[fileCount];
            var paths = new HashSet<string>();

            try
            {
                // Act
                for (int i = 0; i < fileCount; i++)
                {
                    files[i] = TempFile.Create($"Test content {i}", "txt", directoryName: TestDirName);

                    // Verify file exists
                    Assert.IsTrue(File.Exists(files[i].FullPath), $"File {i} should have been created");
                    Assert.IsTrue(paths.Add(files[i].FullPath), $"File {i} should have a unique path");
                }

                // Assert
                Assert.AreEqual(fileCount, paths.Count, "All paths should be unique");
            }
            finally
            {
                // Cleanup
                for (int i = 0; i < fileCount; i++)
                {
                    files[i]?.Dispose();
                }
            }
        }

        [TestMethod]
        public void Create_LargeContent_HandlesEfficiently()
        {
            // Arrange
            const int contentSize = 1024 * 1024; // 1MB
            byte[] largeContent = new byte[contentSize];
            new Random().NextBytes(largeContent); // Fill with random data

            // Act
            using (var tempFile = TempFile.Create(largeContent, "bin", directoryName: TestDirName))
            {
                // Assert
                Assert.IsTrue(File.Exists(tempFile.FullPath), "Large file should have been created");
                Assert.AreEqual(contentSize, new FileInfo(tempFile.FullPath).Length, "File size should be correct");
            }
        }

        #endregion

        #region Edge Case Tests

        [TestMethod]
        public void Create_WithWhitespaceExtension_UsesDefaultExtension()
        {
            // Act
            using (var tempFile = TempFile.Create("Test", "   ", directoryName: TestDirName))
            {
                // Assert
                Assert.AreEqual(".ps1", Path.GetExtension(tempFile.FullPath), "Should use default extension");
            }
        }

        [TestMethod]
        public void Create_EmptyStringContent_CreatesEmptyFile()
        {
            // Act
            using (var tempFile = TempFile.Create(string.Empty, "txt", directoryName: TestDirName))
            {
                // Assert
                Assert.IsTrue(File.Exists(tempFile.FullPath), "File should have been created");
                Assert.AreEqual(0, new FileInfo(tempFile.FullPath).Length, "File should be empty");
            }
        }

        [TestMethod]
        public void Create_EmptyByteArray_CreatesEmptyFile()
        {
            // Act
            using (var tempFile = TempFile.Create(new byte[0], "bin", directoryName: TestDirName))
            {
                // Assert
                Assert.IsTrue(File.Exists(tempFile.FullPath), "File should have been created");
                Assert.AreEqual(0, new FileInfo(tempFile.FullPath).Length, "File should be empty");
            }
        }

        #endregion

        private static void DeleteDirectoryIfExists(string path)
        {
            if (!Directory.Exists(path))
                return;

            try
            {
                Directory.Delete(path, true);
            }
            catch (IOException)
            {
                Thread.Sleep(100);
                Directory.Delete(path, true);
            }
        }

        private static string NormalizeDirectory(string path)
            => path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }
}
