using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Text;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class TestTempFile
    {
        private string? _tempFilePath;

        [TestCleanup]
        public void Cleanup()
        {
            // Ensure cleanup of residual files after all tests
            if (!string.IsNullOrEmpty(_tempFilePath) && File.Exists(_tempFilePath))
            {
                File.Delete(_tempFilePath);
            }
        }

        [TestMethod]
        public void CreateTextFile_ContentMatches()
        {
            // Arrange
            const string testContent = "Unit Test Content";

            // Act
            using var tempFile = TempFile.Create(testContent, "txt");
            _tempFilePath = tempFile.FullPath;

            // Assert
            Assert.IsTrue(File.Exists(tempFile.FullPath));
            Assert.AreEqual(testContent, File.ReadAllText(tempFile.FullPath));
            Assert.IsTrue(tempFile.FileName.EndsWith(".txt"));
        }

        [TestMethod]
        public void CreateBinaryFile_ContentMatches()
        {
            // Arrange
            byte[] testData = Encoding.UTF8.GetBytes("Binary Data");

            // Act
            using var tempFile = TempFile.Create(testData, "bin");
            _tempFilePath = tempFile.FullPath;

            // Assert
            Assert.IsTrue(File.Exists(tempFile.FullPath));
            CollectionAssert.AreEqual(testData, File.ReadAllBytes(tempFile.FullPath));
            Assert.IsTrue(tempFile.FileName.EndsWith(".bin"));
        }

        [TestMethod]
        public void ExtensionHandling_AddsDotWhenMissing()
        {
            // Act
            using var file1 = TempFile.Create("test", "csv");
            using var file2 = TempFile.Create("test", ".log");

            // Assert
            Assert.IsTrue(file1.FileName.EndsWith(".csv"));
            Assert.IsTrue(file2.FileName.EndsWith(".log"));
        }

        [TestMethod]
        public void FileDeletedAfterDispose()
        {
            // Arrange
            string filePath;
            using (var tempFile = TempFile.Create("test"))
            {
                filePath = tempFile.FullPath;
                Assert.IsTrue(File.Exists(filePath));
            }

            // Assert
            Assert.IsFalse(File.Exists(filePath));
        }

        [TestMethod]
        public void DoubleDisposeIsSafe()
        {
            // Arrange
            var tempFile = TempFile.Create("test");
            string filePath = tempFile.FullPath;

            // Act
            tempFile.Dispose();
            tempFile.Dispose(); // Second call

            // Assert
            Assert.IsFalse(File.Exists(filePath));
        }

        [TestMethod]
        public void InvalidExtensionDefaultsToTmp()
        {
            // Act
            using var file1 = TempFile.Create("test", " ");
            using var file2 = TempFile.Create("test", null);

            // Assert
            Assert.IsTrue(file1.FileName.EndsWith(".tmp"));
            Assert.IsTrue(file2.FileName.EndsWith(".tmp"));
        }

        [TestMethod]
        public void FilePathInSystemTempDirectory()
        {
            // Arrange
            string systemTemp = Path.GetTempPath();

            // Act
            using var tempFile = TempFile.Create("test");

            // Assert
            StringAssert.StartsWith(tempFile.FullPath, systemTemp);
        }

        [TestMethod]
        public void ReadMethods_ReturnCorrectContent()
        {
            // Arrange
            const string content = "Test Content";
            using var tempFile = TempFile.Create(content);

            // Act & Assert
            Assert.AreEqual(content, tempFile.ReadAllText());

            using var stream = tempFile.OpenRead();
            using var reader = new StreamReader(stream);
            Assert.AreEqual(content, reader.ReadToEnd());
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void CreateWithNullBytes_ThrowsException()
        {
            // Act
            TempFile.Create(null as byte[]);
        }

        [TestMethod]
        public void CreatePs1File_HasCorrectExtension()
        {
            // Test different extension name formats
            var testCases = new[]
            {
                ("ps1", ".ps1"),
                (".ps1", ".ps1"),
                ("PS1", ".PS1")
            };

            foreach (var (inputExt, expectedExt) in testCases)
            {
                using var tempFile = TempFile.Create("Get-Process", inputExt);
                Assert.IsTrue(tempFile.FileName.EndsWith(expectedExt, StringComparison.OrdinalIgnoreCase));
                Assert.AreEqual("Get-Process", tempFile.ReadAllText());
            }
        }

        [TestMethod]
        public void CreateTextFileWithBom_ContainsBomHeader()
        {
            // Arrange
            const string content = "Write-Host 'Hello'";
            var encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);

            // Act (Create file with BOM using bytes)
            byte[] bomBytes = encoding.GetPreamble();
            byte[] contentBytes = encoding.GetBytes(content);
            byte[] fullContent = [.. bomBytes, .. contentBytes];

            using var tempFile = TempFile.Create(fullContent, "ps1");

            // Assert
            // Verify BOM header
            byte[] fileBytes = File.ReadAllBytes(tempFile.FullPath);
            CollectionAssert.AreEqual(
                new byte[] { 0xEF, 0xBB, 0xBF },
                fileBytes.Take(3).ToArray(),
                "File is missing UTF-8 BOM header"
            );

            // Verify content decoding
            string fileContent = File.ReadAllText(tempFile.FullPath, encoding);
            Assert.AreEqual(content, fileContent);
        }

        [TestMethod]
        public void CreatePs1WithBom_UsingTextApi()
        {
            // Arrange
            const string content = "Get-Content .\\file.txt";
            var encoding = new UTF8Encoding(true);  // BOM encoding

            // Act (Create using text API)
            using var tempFile = TempFile.Create(content, "ps1");
            File.WriteAllText(tempFile.FullPath, content, encoding);  // Overwrite with BOM version

            // Assert
            using var stream = tempFile.OpenRead();
            byte[] bom = new byte[3];
            stream.Read(bom, 0, 3);

            Assert.AreEqual(0xEF, bom[0]);
            Assert.AreEqual(0xBB, bom[1]);
            Assert.AreEqual(0xBF, bom[2]);

            // Verify file content
            stream.Position = 0;  // Reset stream position
            using var reader = new StreamReader(stream, Encoding.UTF8, true);
            Assert.AreEqual(content, reader.ReadToEnd());
        }
    }
}
