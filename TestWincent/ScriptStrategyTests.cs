using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class TestScriptStrategy
    {
        private IPSScriptStrategyFactory _factory;

        [TestInitialize]
        public void Initialize()
        {
            _factory = new DefaultPSScriptStrategyFactory();
        }

        #region 基本功能测试

        [TestMethod]
        public void RefreshExplorerStrategy_GeneratesValidScript()
        {
            // Arrange
            var strategy = new RefreshExplorerStrategy();

            // Act
            var script = strategy.GenerateScript(null);

            // Assert
            StringAssert.Contains(script, "Shell.Application");
            StringAssert.Contains(script, "$_.Refresh()");
            StringAssert.Contains(script, "UTF8");
        }

        [TestMethod]
        public void QueryRecentFileStrategy_ContainsRecentFilesNamespace()
        {
            // Arrange
            var strategy = new QueryRecentFileStrategy();

            // Act
            var script = strategy.GenerateScript(null);

            // Assert
            StringAssert.Contains(script, ShellNamespaces.QuickAccess);
            StringAssert.Contains(script, "$_.IsFolder -eq $false");
        }

        [TestMethod]
        public void QueryFrequentFolderStrategy_UsesCorrectNamespace()
        {
            // Arrange
            var strategy = new QueryFrequentFolderStrategy();

            // Act
            var script = strategy.GenerateScript(null);

            // Assert
            StringAssert.Contains(script, ShellNamespaces.FrequentFolders);
            StringAssert.Contains(script, "ForEach-Object");
        }

        [TestMethod]
        public void QueryQuickAccessStrategy_GeneratesCorrectScript()
        {
            // Arrange
            var strategy = new QueryQuickAccessStrategy();

            // Act
            var script = strategy.GenerateScript(null);

            // Assert
            StringAssert.Contains(script, ShellNamespaces.QuickAccess);
            StringAssert.Contains(script, "ForEach-Object { $_.Path }");
        }

        #endregion

        #region 参数验证测试

        [TestMethod]
        public void RemoveRecentFileStrategy_ValidatesParameter()
        {
            // Arrange
            var strategy = new RemoveRecentFileStrategy();

            // Act & Assert
            var ex = Assert.ThrowsException<ArgumentException>(
                () => strategy.GenerateScript(null));

            Assert.AreEqual("Valid file path parameter required", ex.Message);
        }

        [TestMethod]
        public void PinToFrequentFolder_ValidatesPathParameter()
        {
            // Arrange
            var strategy = new PinToFrequentFolderStrategy();

            // Act & Assert
            var ex = Assert.ThrowsException<ArgumentException>(
                () => strategy.GenerateScript("  "));

            Assert.AreEqual("Valid file path parameter required", ex.Message);
        }

        [TestMethod]
        public void UnpinFromFrequentFolder_ValidatesPathParameter()
        {
            // Arrange
            var strategy = new UnpinFromFrequentFolderStrategy();

            // Act & Assert
            var ex = Assert.ThrowsException<ArgumentException>(
                () => strategy.GenerateScript(string.Empty));

            Assert.AreEqual("Valid file path parameter required", ex.Message);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void EmptyPath_ThrowsArgumentException()
        {
            // Arrange
            var strategy = new PinToFrequentFolderStrategy();

            // Act
            strategy.GenerateScript(string.Empty);
        }

        #endregion

        #region 工厂测试

        [TestMethod]
        public void Factory_ReturnsCorrectStrategies()
        {
            // Act & Assert
            Assert.IsInstanceOfType(
                _factory.GetStrategy(PSScript.RefreshExplorer),
                typeof(RefreshExplorerStrategy));

            Assert.IsInstanceOfType(
                _factory.GetStrategy(PSScript.QueryRecentFile),
                typeof(QueryRecentFileStrategy));

            Assert.IsInstanceOfType(
                _factory.GetStrategy(PSScript.CheckQueryFeasible),
                typeof(CheckQueryFeasibleStrategy));

            Assert.IsInstanceOfType(
                _factory.GetStrategy(PSScript.PinToFrequentFolder),
                typeof(PinToFrequentFolderStrategy));

            Assert.IsInstanceOfType(
                _factory.GetStrategy(PSScript.UnpinFromFrequentFolder),
                typeof(UnpinFromFrequentFolderStrategy));
        }

        [TestMethod]
        public void Factory_ThrowsOnUnsupportedMethod()
        {
            // Arrange
            var invalidMethod = (PSScript)100;

            // Act & Assert
            var ex = Assert.ThrowsException<NotSupportedException>(
                () => _factory.GetStrategy(invalidMethod));

            StringAssert.Contains(ex.Message, "Unsupported script type: 100");
        }

        #endregion

        #region 路径处理测试

        [TestMethod]
        public void PathWithSpaces_HandlesCorrectly()
        {
            // Arrange
            var strategy = new PinToFrequentFolderStrategy();
            var path = @"C:\Program Files\Test Folder";

            // Act
            var script = strategy.GenerateScript(path);

            // Assert
            StringAssert.Contains(script, path);
        }

        [TestMethod]
        public void PathWithUnicodeCharacters_HandlesCorrectly()
        {
            // Arrange
            var strategy = new RemoveRecentFileStrategy();
            var path = @"C:\测试文件夹\文件.txt";

            // Act
            var script = strategy.GenerateScript(path);

            // Assert
            StringAssert.Contains(script, path);
        }

        [TestMethod]
        public void PathWithDots_HandlesCorrectly()
        {
            // Arrange
            var strategy = new UnpinFromFrequentFolderStrategy();
            var path = @"C:\folder.name\file..txt";

            // Act
            var script = strategy.GenerateScript(path);

            // Assert
            StringAssert.Contains(script, path);
        }

        [TestMethod]
        public void PathWithValidSpecialCharacters_HandlesCorrectly()
        {
            // Arrange
            var strategy = new RemoveRecentFileStrategy();
            var path = @"C:\folder-name\file (1)_[test]~$temp.txt";

            // Act
            var script = strategy.GenerateScript(path);

            // Assert
            StringAssert.Contains(script, path);
        }

        #endregion

        #region 脚本内容测试

        [TestMethod]
        public void ValidPinOperation_GeneratesCorrectVerb()
        {
            // Arrange
            var strategy = new PinToFrequentFolderStrategy();
            const string testPath = @"C:\test";

            // Act
            var script = strategy.GenerateScript(testPath);

            // Assert
            StringAssert.Contains(script, "InvokeVerb('pintohome')");
            StringAssert.Contains(script, testPath);
        }

        [TestMethod]
        public void ValidUnpinOperation_GeneratesCorrectVerb()
        {
            // Arrange
            var strategy = new UnpinFromFrequentFolderStrategy();
            const string testPath = @"C:\test";

            // Act
            var script = strategy.GenerateScript(testPath);

            // Assert
            StringAssert.Contains(script, "InvokeVerb('unpinfromhome')");
            StringAssert.Contains(script, testPath);
        }

        [TestMethod]
        public void CheckPinUnpinFeasible_ContainsValidationLogic()
        {
            // Arrange
            var strategy = new CheckPinUnpinFeasibleStrategy();

            // Act
            var script = strategy.GenerateScript(null);

            // Assert
            StringAssert.Contains(script, "InvokeVerb('pintohome')");
            StringAssert.Contains(script, "InvokeVerb('unpinfromhome')");
        }

        #endregion

        #region 文件生成测试

        [TestMethod]
        public void GenerateAndSaveStrategyScripts()
        {
            // 设置永久保存目录为系统临时目录下的 WincentTemp
            var saveDir = Path.Combine(
                Path.GetTempPath(),
                "WincentTemp",
                "Scripts");  // 添加 Scripts 子目录以区分其他临时文件

            // 确保目录存在，如果已存在则清空
            if (Directory.Exists(saveDir))
            {
                Directory.Delete(saveDir, true);
            }
            Directory.CreateDirectory(saveDir);

            // 测试参数
            var testParams = new Dictionary<PSScript, string>
            {
                [PSScript.RefreshExplorer] = null,
                [PSScript.QueryQuickAccess] = null,
                [PSScript.QueryRecentFile] = null,
                [PSScript.QueryFrequentFolder] = null,
                [PSScript.RemoveRecentFile] = @"C:\test\file.txt",
                [PSScript.PinToFrequentFolder] = @"C:\test\folder",
                [PSScript.UnpinFromFrequentFolder] = @"C:\test\folder",
                [PSScript.CheckQueryFeasible] = null,
                [PSScript.CheckPinUnpinFeasible] = null
            };

            try
            {
                foreach (var pair in testParams)
                {
                    // 获取策略实例
                    var scriptType = pair.Key;
                    var param = pair.Value;
                    var strategy = _factory.GetStrategy(scriptType);
                    var scriptContent = strategy.GenerateScript(param);

                    // 转为 UTF8 编码
                    byte[] scriptBytes = Encoding.UTF8.GetBytes(scriptContent);
                    byte[] contentWithBom = ScriptExecutor.AddUtf8Bom(scriptBytes);

                    // 创建临时文件
                    using (var tempFile = TempFile.Create(contentWithBom, "ps1"))
                    {
                        Assert.IsTrue(File.Exists(tempFile.FullPath), $"临时文件未创建: {scriptType}");

                        // 保存永久文件
                        var permanentPath = Path.Combine(saveDir, $"{scriptType}.ps1");
                        File.Copy(tempFile.FullPath, permanentPath, true);
                        Assert.IsTrue(File.Exists(permanentPath), $"永久文件未创建: {scriptType}");

                        // 验证文件内容
                        var savedContent = File.ReadAllText(permanentPath);
                        Assert.AreEqual(scriptContent, savedContent, $"文件内容不匹配: {scriptType}");

                        // 验证文件编码（应为 UTF8-BOM）
                        var fileBytes = File.ReadAllBytes(permanentPath);
                        var hasBom = fileBytes.Length >= 3 &&
                                     fileBytes[0] == 0xEF &&
                                     fileBytes[1] == 0xBB &&
                                     fileBytes[2] == 0xBF;
                        Assert.IsTrue(hasBom, $"文件应该使用 UTF8-BOM 编码: {scriptType}");
                    }
                }

                // 输出保存位置信息
                Console.WriteLine($"脚本文件已保存到: {saveDir}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"测试过程中发生错误: {ex.Message}");
                throw;
            }
            finally
            {
                // 清理测试目录
                if (Directory.Exists(saveDir))
                {
                    try
                    {
                        Directory.Delete(saveDir, true);
                    }
                    catch
                    {
                        // 忽略清理错误
                    }
                }
            }
        }

        #endregion

        #region 异常情况测试

        [TestMethod]
        public void LongPathParameter_HandlesCorrectly()
        {
            // Arrange
            var strategy = new PinToFrequentFolderStrategy();
            var path = @"C:\" + new string('a', 200) + @"\test.txt"; // 长但有效的路径

            // Act
            var script = strategy.GenerateScript(path);

            // Assert
            StringAssert.Contains(script, path);
        }

        [TestMethod]
        public void MultipleStrategies_GenerateDifferentScripts()
        {
            // Arrange
            var strategy1 = new QueryRecentFileStrategy();
            var strategy2 = new QueryFrequentFolderStrategy();

            // Act
            var script1 = strategy1.GenerateScript(null);
            var script2 = strategy2.GenerateScript(null);

            // Assert
            Assert.AreNotEqual(script1, script2, "不同策略应生成不同的脚本");
        }

        #endregion
    }
}
