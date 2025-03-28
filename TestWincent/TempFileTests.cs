using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Text;
using Wincent;
using System.Threading;

namespace TestWincent
{
    [TestClass]
    public class TestTempFile
    {
        private string _originalDirName;

        [TestInitialize]
        public void Initialize()
        {
            // 保存原始目录名称
            _originalDirName = TempFile.DirName;

            // 设置测试专用目录名称
            TempFile.DirName = "WincentTempTest";

            // 确保测试目录不存在（清理上次测试可能的残留）
            string testDir = Path.Combine(Path.GetTempPath(), TempFile.DirName);
            if (Directory.Exists(testDir))
            {
                try
                {
                    Directory.Delete(testDir, true);
                }
                catch (IOException)
                {
                    // 如果文件被锁定，等待一会再试
                    Thread.Sleep(100);
                    Directory.Delete(testDir, true);
                }
            }
        }

        [TestCleanup]
        public void Cleanup()
        {
            // 恢复原始目录名称
            TempFile.DirName = _originalDirName;

            // 清理测试目录
            string testDir = Path.Combine(Path.GetTempPath(), "WincentTempTest");
            if (Directory.Exists(testDir))
            {
                try
                {
                    Directory.Delete(testDir, true);
                }
                catch
                {
                    // 忽略清理错误
                }
            }
        }

        #region 基本功能测试

        [TestMethod]
        public void Create_WithStringContent_CreatesFile()
        {
            // Arrange
            const string content = "Test content";

            // Act
            using (var tempFile = TempFile.Create(content, "txt"))
            {
                // Assert
                Assert.IsTrue(File.Exists(tempFile.FullPath), "文件应该被创建");
                Assert.AreEqual(content, File.ReadAllText(tempFile.FullPath), "文件内容应该匹配");
                Assert.AreEqual(".txt", Path.GetExtension(tempFile.FullPath), "文件扩展名应该匹配");
            }
        }

        [TestMethod]
        public void Create_WithBinaryContent_CreatesFile()
        {
            // Arrange
            byte[] content = Encoding.UTF8.GetBytes("Binary test content");

            // Act
            using (var tempFile = TempFile.Create(content, "bin"))
            {
                // Assert
                Assert.IsTrue(File.Exists(tempFile.FullPath), "文件应该被创建");
                CollectionAssert.AreEqual(content, File.ReadAllBytes(tempFile.FullPath), "二进制内容应该匹配");
                Assert.AreEqual(".bin", Path.GetExtension(tempFile.FullPath), "文件扩展名应该匹配");
            }
        }

        [TestMethod]
        public void Create_WithCustomEncoding_CreatesFile()
        {
            // Arrange
            const string content = "测试内容";
            byte[] expectedBytes = Encoding.Unicode.GetBytes(content);

            // Act
            using (var tempFile = TempFile.Create(expectedBytes, "txt", Encoding.Unicode))
            {
                // Assert
                Assert.IsTrue(File.Exists(tempFile.FullPath), "文件应该被创建");
                byte[] actualBytes = File.ReadAllBytes(tempFile.FullPath);
                CollectionAssert.AreEqual(expectedBytes, actualBytes, "编码后的内容应该匹配");
            }
        }

        [TestMethod]
        public void Dispose_DeletesFile()
        {
            // Arrange
            string filePath;

            // Act
            using (var tempFile = TempFile.Create("Test content", "txt"))
            {
                filePath = tempFile.FullPath;
                Assert.IsTrue(File.Exists(filePath), "文件应该在使用期间存在");
            }

            // Assert
            Assert.IsFalse(File.Exists(filePath), "文件应该在Dispose后被删除");
        }

        #endregion

        #region 属性测试

        [TestMethod]
        public void FullPath_ReturnsCorrectPath()
        {
            // Arrange & Act
            using (var tempFile = TempFile.Create("Test", "txt"))
            {
                // Assert
                Assert.IsTrue(tempFile.FullPath.StartsWith(Path.GetTempPath()), "路径应该以临时目录开始");
                Assert.IsTrue(tempFile.FullPath.Contains(TempFile.DirName), "路径应该包含自定义目录名");
                Assert.IsTrue(tempFile.FullPath.EndsWith(".txt"), "路径应该以正确的扩展名结束");
            }
        }

        [TestMethod]
        public void FileName_ReturnsCorrectName()
        {
            // Arrange & Act
            using (var tempFile = TempFile.Create("Test", "txt"))
            {
                // Assert
                string fileName = tempFile.FileName;
                Assert.IsTrue(fileName.EndsWith(".txt"), "文件名应该包含扩展名");
                Assert.AreEqual(Path.GetFileName(tempFile.FullPath), fileName, "FileName应该返回文件名部分");
            }
        }

        [TestMethod]
        public void DirName_ChangesAffectNewFiles()
        {
            // Arrange
            const string newDirName = "CustomTempDir";
            TempFile.DirName = newDirName;

            // Act
            using (var tempFile = TempFile.Create("Test", "txt"))
            {
                // Assert
                Assert.IsTrue(tempFile.FullPath.Contains(newDirName), "新目录名应该被应用");
                Assert.IsFalse(tempFile.FullPath.Contains(_originalDirName), "原目录名不应该被使用");
            }
        }

        #endregion

        #region 异常处理测试

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Create_WithNullStringContent_ThrowsException()
        {
            // Act
            TempFile.Create((string)null, "txt");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Create_WithNullBinaryContent_ThrowsException()
        {
            // Act
            TempFile.Create((byte[])null, "bin");
        }

        [TestMethod]
        public void Create_WithNullExtension_UsesDefaultExtension()
        {
            // Act
            using (var tempFile = TempFile.Create("Test", null))
            {
                // Assert
                Assert.AreEqual(".ps1", Path.GetExtension(tempFile.FullPath), "应该使用默认扩展名");
            }
        }

        [TestMethod]
        public void Create_WithEmptyExtension_UsesDefaultExtension()
        {
            // Act
            using (var tempFile = TempFile.Create("Test", string.Empty))
            {
                // Assert
                Assert.AreEqual(".ps1", Path.GetExtension(tempFile.FullPath), "应该使用默认扩展名");
            }
        }

        [TestMethod]
        public void Create_WithExtensionWithoutDot_AddsDot()
        {
            // Act
            using (var tempFile = TempFile.Create("Test", "ext"))
            {
                // Assert
                Assert.AreEqual(".ext", Path.GetExtension(tempFile.FullPath), "应该添加点号");
            }
        }

        [TestMethod]
        public void Create_WithExtensionWithDot_PreservesDot()
        {
            // Act
            using (var tempFile = TempFile.Create("Test", ".ext"))
            {
                // Assert
                Assert.AreEqual(".ext", Path.GetExtension(tempFile.FullPath), "应该保留点号");
            }
        }

        [TestMethod]
        public void ReadAllText_AfterDispose_ThrowsObjectDisposedException()
        {
            // Arrange
            TempFile tempFile = TempFile.Create("Test", "txt");
            tempFile.Dispose();

            // Act & Assert
            Assert.ThrowsException<ObjectDisposedException>(() => tempFile.ReadAllText());
        }

        [TestMethod]
        public void OpenRead_AfterDispose_ThrowsObjectDisposedException()
        {
            // Arrange
            TempFile tempFile = TempFile.Create("Test", "txt");
            tempFile.Dispose();

            // Act & Assert
            Assert.ThrowsException<ObjectDisposedException>(() => tempFile.OpenRead());
        }

        #endregion

        #region 文件操作测试

        [TestMethod]
        public void ReadAllText_ReturnsCorrectContent()
        {
            // Arrange
            const string content = "Test content for reading";

            // Act
            using (var tempFile = TempFile.Create(content, "txt"))
            {
                // Assert
                Assert.AreEqual(content, tempFile.ReadAllText(), "ReadAllText应该返回正确内容");
            }
        }

        [TestMethod]
        public void ReadAllText_WithEncoding_ReturnsCorrectContent()
        {
            // Arrange
            const string content = "测试内容";
            byte[] contentBytes = Encoding.Unicode.GetBytes(content);

            // Act
            using (var tempFile = TempFile.Create(contentBytes, "txt"))
            {
                // Assert
                Assert.AreEqual(content, tempFile.ReadAllText(Encoding.Unicode), "使用指定编码应该返回正确内容");
            }
        }

        [TestMethod]
        public void OpenRead_ReturnsValidStream()
        {
            // Arrange
            const string content = "Test content for stream";

            // Act
            using (var tempFile = TempFile.Create(content, "txt"))
            {
                using (var stream = tempFile.OpenRead())
                {
                    // Assert
                    Assert.IsTrue(stream.CanRead, "流应该可读");
                    Assert.IsFalse(stream.CanWrite, "流应该不可写");

                    byte[] buffer = new byte[100];
                    int bytesRead = stream.Read(buffer, 0, buffer.Length);
                    string readContent = Encoding.UTF8.GetString(buffer, 0, bytesRead);

                    Assert.AreEqual(content, readContent, "流应该包含正确内容");
                }
            }
        }

        #endregion

        #region 并发和性能测试

        [TestMethod]
        public void Create_MultipleTempFiles_CreatesUniqueFiles()
        {
            // Arrange
            const int fileCount = 10;
            var files = new TempFile[fileCount];
            var paths = new string[fileCount];

            try
            {
                // Act
                for (int i = 0; i < fileCount; i++)
                {
                    files[i] = TempFile.Create($"Test content {i}", "txt");
                    paths[i] = files[i].FullPath;

                    // 验证文件存在
                    Assert.IsTrue(File.Exists(paths[i]), $"文件 {i} 应该被创建");
                }

                // Assert - 验证所有路径都是唯一的
                for (int i = 0; i < fileCount; i++)
                {
                    for (int j = i + 1; j < fileCount; j++)
                    {
                        Assert.AreNotEqual(paths[i], paths[j], $"文件 {i} 和 {j} 应该有不同的路径");
                    }
                }
            }
            finally
            {
                // 清理
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
            new Random().NextBytes(largeContent); // 填充随机数据

            // Act
            using (var tempFile = TempFile.Create(largeContent, "bin"))
            {
                // Assert
                Assert.IsTrue(File.Exists(tempFile.FullPath), "大文件应该被创建");
                Assert.AreEqual(contentSize, new FileInfo(tempFile.FullPath).Length, "文件大小应该正确");
            }
        }

        #endregion

        #region 边界情况测试

        [TestMethod]
        public void Create_WithWhitespaceExtension_UsesDefaultExtension()
        {
            // Act
            using (var tempFile = TempFile.Create("Test", "   "))
            {
                // Assert
                Assert.AreEqual(".ps1", Path.GetExtension(tempFile.FullPath), "应该使用默认扩展名");
            }
        }

        [TestMethod]
        public void Create_EmptyStringContent_CreatesEmptyFile()
        {
            // Act
            using (var tempFile = TempFile.Create(string.Empty, "txt"))
            {
                // Assert
                Assert.IsTrue(File.Exists(tempFile.FullPath), "文件应该被创建");
                Assert.AreEqual(0, new FileInfo(tempFile.FullPath).Length, "文件应该为空");
            }
        }

        [TestMethod]
        public void Create_EmptyByteArray_CreatesEmptyFile()
        {
            // Act
            using (var tempFile = TempFile.Create(new byte[0], "bin"))
            {
                // Assert
                Assert.IsTrue(File.Exists(tempFile.FullPath), "文件应该被创建");
                Assert.AreEqual(0, new FileInfo(tempFile.FullPath).Length, "文件应该为空");
            }
        }

        #endregion
    }
}
