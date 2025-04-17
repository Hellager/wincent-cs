using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class TestScriptStorage
    {
        private string _testFilePath;
        private string _testScriptVersion;

        [TestInitialize]
        public void Initialize()
        {
            // 获取当前程序集版本号，与ScriptStorage中保持一致
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var version = assembly.GetName().Version;
                _testScriptVersion = $"v{version.Major}.{version.Minor}.{version.Build}";
            }
            catch
            {
                _testScriptVersion = "v1.0.0";
            }

            // 创建测试文件，用于测试清理功能
            _testFilePath = Path.Combine(
                ScriptStorage.DynamicScriptDir,
                $"TestScript_{_testScriptVersion}_{Guid.NewGuid():N}.ps1");

            File.WriteAllText(_testFilePath, "Test content");

            // 确保测试文件存在
            Assert.IsTrue(File.Exists(_testFilePath), "测试文件应该被创建");
        }

        [TestCleanup]
        public void Cleanup()
        {
            // 清理测试文件
            if (File.Exists(_testFilePath))
            {
                try
                {
                    File.Delete(_testFilePath);
                }
                catch
                {
                    // 忽略删除错误
                }
            }
        }

        #region 目录结构测试

        [TestMethod]
        public void ScriptRoot_ExistsAndIsInTempFolder()
        {
            // Assert
            Assert.IsTrue(Directory.Exists(ScriptStorage.ScriptRoot), "脚本根目录应该存在");
            Assert.IsTrue(ScriptStorage.ScriptRoot.StartsWith(Path.GetTempPath()), "脚本根目录应该在临时文件夹中");
            Assert.IsTrue(ScriptStorage.ScriptRoot.EndsWith("Wincent"), "脚本根目录应该命名为Wincent");
        }

        [TestMethod]
        public void StaticScriptDir_ExistsAndIsInRootFolder()
        {
            // Assert
            Assert.IsTrue(Directory.Exists(ScriptStorage.StaticScriptDir), "静态脚本目录应该存在");
            Assert.IsTrue(ScriptStorage.StaticScriptDir.StartsWith(ScriptStorage.ScriptRoot), "静态脚本目录应该在根目录中");
            Assert.IsTrue(ScriptStorage.StaticScriptDir.EndsWith("static"), "静态脚本目录应该命名为static");
        }

        [TestMethod]
        public void DynamicScriptDir_ExistsAndIsInRootFolder()
        {
            // Assert
            Assert.IsTrue(Directory.Exists(ScriptStorage.DynamicScriptDir), "动态脚本目录应该存在");
            Assert.IsTrue(ScriptStorage.DynamicScriptDir.StartsWith(ScriptStorage.ScriptRoot), "动态脚本目录应该在根目录中");
            Assert.IsTrue(ScriptStorage.DynamicScriptDir.EndsWith("dynamic"), "动态脚本目录应该命名为dynamic");
        }

        #endregion

        #region 脚本路径测试

        [TestMethod]
        public void GetScriptPath_ForNonParameterizedScript_ReturnsCorrectPath()
        {
            // Arrange
            var script = PSScript.RefreshExplorer;

            // Act
            string path = ScriptStorage.GetScriptPath(script);

            // Assert
            Assert.IsTrue(path.StartsWith(ScriptStorage.StaticScriptDir), "非参数化脚本应该在静态目录中");
            Assert.IsTrue(path.Contains($"{script}_{_testScriptVersion}"), "脚本文件名应该包含版本号");
            Assert.IsTrue(path.EndsWith("ps1"), "脚本文件名应该以ps1为扩展名");
            Assert.IsTrue(File.Exists(path), "脚本文件应该被创建");

            // 验证文件内容
            string content = File.ReadAllText(path);
            Assert.IsTrue(content.Contains("Shell.Application"), "脚本内容应该正确");

            // 验证文件编码
            byte[] bytes = File.ReadAllBytes(path);
            Assert.IsTrue(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF,
                "脚本文件应该使用 UTF8-BOM 编码");

            // 验证文件名格式
            var fileName = Path.GetFileName(path);
            Assert.IsTrue(Regex.IsMatch(fileName, @"^RefreshExplorer_v\d+\.\d+\.\d+\.ps1$"),
                "文件名格式应为 {ScriptName}_{Version}.ps1");
        }

        [TestMethod]
        public void GetScriptPath_ForParameterizedScript_ReturnsCorrectPath()
        {
            // Arrange
            var script = PSScript.RemoveRecentFile;

            // Act
            string path = ScriptStorage.GetScriptPath(script);

            // Assert
            Assert.IsTrue(path.StartsWith(ScriptStorage.DynamicScriptDir), "参数化脚本应该在动态目录中");
            Assert.IsTrue(path.Contains($"{script}_{_testScriptVersion}"), "脚本文件名应该包含版本号");
            Assert.IsTrue(path.EndsWith("ps1"), "脚本文件名应该以ps1为扩展名");
            Assert.IsFalse(File.Exists(path), "参数化脚本不应该被自动创建");
        }

        [TestMethod]
        public void GetDynamicScriptPath_WithValidParameter_ReturnsAndCreatesScript()
        {
            // Arrange
            var script = PSScript.RemoveRecentFile;
            var parameter = @"C:\Test\File.txt";

            // Act
            string path = ScriptStorage.GetDynamicScriptPath(script, parameter);

            // Assert
            Assert.IsTrue(path.StartsWith(ScriptStorage.DynamicScriptDir), "动态脚本应该在动态目录中");
            Assert.IsTrue(path.Contains(script.ToString()), "脚本文件名应该包含脚本类型");
            Assert.IsTrue(path.Contains(_testScriptVersion), "脚本文件名应该包含版本号");
            Assert.IsTrue(File.Exists(path), "动态脚本文件应该被创建");

            // 验证文件内容
            string content = File.ReadAllText(path);
            Assert.IsTrue(content.Contains(parameter), "脚本内容应该包含参数");

            // 验证文件编码
            byte[] bytes = File.ReadAllBytes(path);
            Assert.IsTrue(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF,
                "脚本文件应该使用 UTF8-BOM 编码");

            // 验证文件名格式
            var fileName = Path.GetFileName(path);
            Assert.IsTrue(Regex.IsMatch(fileName, @"^RemoveRecentFile_v\d+\.\d+\.\d+_[0-9A-F]{8}\.ps1$"),
                "文件名格式应为 {ScriptName}_{Version}_{ParamHash}.ps1");
        }

        [TestMethod]
        public void GetDynamicScriptPath_WithQuotesInParameter_HandlesCorrectly()
        {
            // Arrange
            var script = PSScript.RemoveRecentFile;
            var parameter = @"C:\Test\File's.txt";

            // Act
            string path = ScriptStorage.GetDynamicScriptPath(script, parameter);

            // Assert
            Assert.IsTrue(File.Exists(path), "包含单引号的脚本文件应该被创建");

            // 验证文件内容
            string content = File.ReadAllText(path);
            Assert.IsFalse(content.Contains("File's.txt"), "单引号应该被转义");
            Assert.IsTrue(content.Contains("File''s.txt"), "单引号应该被替换为两个单引号");

            // 清理
            File.Delete(path);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void GetDynamicScriptPath_WithNonParameterizedScript_ThrowsException()
        {
            // Arrange
            var script = PSScript.RefreshExplorer;
            var parameter = @"C:\Test\File.txt";

            // Act
            ScriptStorage.GetDynamicScriptPath(script, parameter);

            // Assert - 应该抛出异常
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void GetDynamicScriptPath_WithEmptyParameter_ThrowsException()
        {
            // Arrange
            var script = PSScript.RemoveRecentFile;
            var parameter = "";

            // Act
            ScriptStorage.GetDynamicScriptPath(script, parameter);

            // Assert - 应该抛出异常
        }

        #endregion

        #region UTF8 BOM 测试

        [TestMethod]
        public void AddUtf8Bom_WithNoBom_AddsBom()
        {
            // Arrange
            var content = Encoding.UTF8.GetBytes("test content");

            // Act
            var result = ScriptStorage.AddUtf8Bom(content);

            // Assert
            Assert.AreEqual(content.Length + 3, result.Length, "应该添加3字节的BOM");
            Assert.AreEqual(0xEF, result[0], "第一个字节应该是0xEF");
            Assert.AreEqual(0xBB, result[1], "第二个字节应该是0xBB");
            Assert.AreEqual(0xBF, result[2], "第三个字节应该是0xBF");
        }

        [TestMethod]
        public void AddUtf8Bom_WithExistingBom_DoesNotDuplicate()
        {
            // Arrange
            var bom = Encoding.UTF8.GetPreamble();
            var text = Encoding.UTF8.GetBytes("test content");
            var contentWithBom = new byte[bom.Length + text.Length];
            Buffer.BlockCopy(bom, 0, contentWithBom, 0, bom.Length);
            Buffer.BlockCopy(text, 0, contentWithBom, bom.Length, text.Length);

            // Act
            var result = ScriptStorage.AddUtf8Bom(contentWithBom);

            // Assert
            Assert.AreEqual(contentWithBom.Length, result.Length, "不应该添加重复的BOM");
            // 验证内容相同
            for (int i = 0; i < contentWithBom.Length; i++)
            {
                Assert.AreEqual(contentWithBom[i], result[i], $"字节 {i} 应该相同");
            }
        }

        #endregion

        #region 清理功能测试

        [TestMethod]
        public void CleanupDynamicScripts_RemovesOldScripts()
        {
            // Arrange - 创建一些测试文件，设置不同的最后写入时间
            var testFiles = new string[3];
            for (int i = 0; i < testFiles.Length; i++)
            {
                testFiles[i] = Path.Combine(
                    ScriptStorage.DynamicScriptDir,
                    $"TestCleanup_{_testScriptVersion}_{i}.ps1");

                File.WriteAllText(testFiles[i], $"Test content {i}");

                // 设置不同的最后写入时间
                File.SetLastWriteTime(testFiles[i], DateTime.Now.AddHours(-i * 10 - 1));
            }

            // Act - 清理超过5小时的文件
            ScriptStorage.CleanupDynamicScripts(5);

            // Assert
            Assert.IsTrue(File.Exists(testFiles[0]), "新文件不应该被删除");
            Assert.IsFalse(File.Exists(testFiles[1]), "旧文件应该被删除");
            Assert.IsFalse(File.Exists(testFiles[2]), "旧文件应该被删除");

            // 清理
            if (File.Exists(testFiles[0]))
                File.Delete(testFiles[0]);
        }

        [TestMethod]
        public void CleanupByVersion_RemovesOldVersionScripts()
        {
            // Arrange - 创建使用旧版本号的测试文件
            var oldVersion = "v0.0.1";
            var testFile = Path.Combine(
                ScriptStorage.StaticScriptDir,
                $"RefreshExplorer_{oldVersion}.ps1");

            File.WriteAllText(testFile, "Test content for old version");

            // 确保文件存在
            Assert.IsTrue(File.Exists(testFile), "测试文件应该被创建");

            // Act - 调用CleanupDynamicScripts，里面会调用CleanupByVersion
            ScriptStorage.CleanupDynamicScripts();

            // Assert
            Assert.IsFalse(File.Exists(testFile), "旧版本脚本应该被删除");

            // 查看是否创建了新版本脚本
            var newScript = ScriptStorage.GetScriptPath(PSScript.RefreshExplorer);
            Assert.IsTrue(File.Exists(newScript), "应该创建新版本脚本");
            Assert.IsTrue(newScript.Contains(_testScriptVersion), "新脚本应该使用当前版本号");
        }

        #endregion

        #region 哈希生成测试

        [TestMethod]
        public void GetParameterHash_DifferentParameters_GenerateDifferentHashes()
        {
            // Arrange
            var script = PSScript.RemoveRecentFile;
            var param1 = @"C:\Test\File1.txt";
            var param2 = @"C:\Test\File2.txt";

            // Act
            string path1 = ScriptStorage.GetDynamicScriptPath(script, param1);
            string path2 = ScriptStorage.GetDynamicScriptPath(script, param2);

            // 从路径中提取哈希部分 - 使用正则表达式提取
            string fileName1 = Path.GetFileName(path1);
            string fileName2 = Path.GetFileName(path2);

            // 新的文件名格式为: {ScriptName}_{Version}_{ParamHash}.ps1
            var match1 = Regex.Match(fileName1, @"^[A-Za-z]+_v\d+\.\d+\.\d+_([0-9A-F]{8})\.ps1$");
            var match2 = Regex.Match(fileName2, @"^[A-Za-z]+_v\d+\.\d+\.\d+_([0-9A-F]{8})\.ps1$");

            Assert.IsTrue(match1.Success, "文件名1应该符合格式要求");
            Assert.IsTrue(match2.Success, "文件名2应该符合格式要求");

            string hash1 = match1.Groups[1].Value;
            string hash2 = match2.Groups[1].Value;

            // Assert
            Assert.AreNotEqual(hash1, hash2, "不同参数应该生成不同的哈希值");
            Assert.AreEqual(8, hash1.Length, "哈希值长度应该为8个字符");
            Assert.AreEqual(8, hash2.Length, "哈希值长度应该为8个字符");

            // 清理
            File.Delete(path1);
            File.Delete(path2);
        }

        [TestMethod]
        public void GetParameterHash_SameParameter_GeneratesSameHash()
        {
            // Arrange
            var script = PSScript.RemoveRecentFile;
            var param = @"C:\Test\File.txt";

            // Act
            string path1 = ScriptStorage.GetDynamicScriptPath(script, param);
            string path2 = ScriptStorage.GetDynamicScriptPath(script, param);

            // Assert
            Assert.AreEqual(path1, path2, "相同参数应该生成相同的路径");
            Assert.IsTrue(File.Exists(path1), "脚本文件应该被创建");

            // 清理
            File.Delete(path1);
        }

        #endregion

        #region 集成测试

        [TestMethod]
        public void ScriptStorage_IntegrationTest_CreateAndExecuteScript()
        {
            // Arrange
            var script = PSScript.QueryQuickAccess;

            // Act - 获取脚本路径
            string scriptPath = ScriptStorage.GetScriptPath(script);

            // Assert
            Assert.IsTrue(File.Exists(scriptPath), "脚本文件应该被创建");

            // 验证文件内容
            string content = File.ReadAllText(scriptPath);
            Assert.IsTrue(content.Contains("Shell.Application"), "脚本内容应该正确");
            Assert.IsTrue(content.Contains("Namespace"), "脚本内容应该正确");

            // 验证文件编码
            byte[] bytes = File.ReadAllBytes(scriptPath);
            Assert.IsTrue(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF,
                "脚本文件应该使用 UTF8-BOM 编码");
        }

        [TestMethod]
        public void ScriptStorage_IsParameterizedScript_CorrectlyIdentifies()
        {
            // 参数化脚本
            Assert.IsTrue(ScriptStorage.IsParameterizedScript(PSScript.RemoveRecentFile),
                "RemoveRecentFile 应该被识别为参数化脚本");
            Assert.IsTrue(ScriptStorage.IsParameterizedScript(PSScript.PinToFrequentFolder),
                "PinToFrequentFolder 应该被识别为参数化脚本");
            Assert.IsTrue(ScriptStorage.IsParameterizedScript(PSScript.UnpinFromFrequentFolder),
                "UnpinFromFrequentFolder 应该被识别为参数化脚本");
            Assert.IsTrue(ScriptStorage.IsParameterizedScript(PSScript.AddRecentFile),
                "AddRecentFile 应该被识别为参数化脚本");

            // 非参数化脚本
            Assert.IsFalse(ScriptStorage.IsParameterizedScript(PSScript.RefreshExplorer),
                "RefreshExplorer 不应该被识别为参数化脚本");
            Assert.IsFalse(ScriptStorage.IsParameterizedScript(PSScript.QueryQuickAccess),
                "QueryQuickAccess 不应该被识别为参数化脚本");
            Assert.IsFalse(ScriptStorage.IsParameterizedScript(PSScript.QueryRecentFile),
                "QueryRecentFile 不应该被识别为参数化脚本");
            Assert.IsFalse(ScriptStorage.IsParameterizedScript(PSScript.QueryFrequentFolder),
                "QueryFrequentFolder 不应该被识别为参数化脚本");
            Assert.IsFalse(ScriptStorage.IsParameterizedScript(PSScript.CheckQueryFeasible),
                "CheckQueryFeasible 不应该被识别为参数化脚本");
            Assert.IsFalse(ScriptStorage.IsParameterizedScript(PSScript.CheckPinUnpinFeasible),
                "CheckPinUnpinFeasible 不应该被识别为参数化脚本");
        }

        #endregion
    }
}
