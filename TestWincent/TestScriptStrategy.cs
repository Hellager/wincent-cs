﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Text;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class PSScriptStrategyTests
    {
        private DefaultPSScriptStrategyFactory? _factory;

        [TestInitialize]
        public void Initialize()
        {
            _factory = new DefaultPSScriptStrategyFactory();
        }

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
        public void Factory_ReturnsCorrectStrategies()
        {
            // Arrange
            Assert.IsNotNull(_factory, "Factory should be initialized in TestInitialize");
            
            // Act & Assert
            Assert.IsInstanceOfType(
                _factory!.GetStrategy(PSScript.RefreshExplorer),
                typeof(RefreshExplorerStrategy));

            Assert.IsInstanceOfType(
                _factory.GetStrategy(PSScript.QueryRecentFile),
                typeof(QueryRecentFileStrategy));

            Assert.IsInstanceOfType(
                _factory.GetStrategy(PSScript.CheckQueryFeasible),
                typeof(CheckQueryFeasibleStrategy));
        }

        [TestMethod]
        public void Factory_ThrowsOnUnsupportedMethod()
        {
            // Arrange
            var invalidMethod = (PSScript)100;
            Assert.IsNotNull(_factory, "Factory should be initialized in TestInitialize");

            // Act & Assert
            var ex = Assert.ThrowsException<NotSupportedException>(
                () => _factory!.GetStrategy(invalidMethod));

            StringAssert.Contains(ex.Message, "Unsupported script type: 100");
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

        [TestMethod]
        public void QueryFrequentFolder_UsesCorrectNamespace()
        {
            // Arrange
            var strategy = new QueryFrequentFolderStrategy();

            // Act
            var script = strategy.GenerateScript(null);

            // Assert
            StringAssert.Contains(script, ShellNamespaces.FrequentFolders);
            StringAssert.Contains(script, ".Items() | \r\n                ForEach-Object");
        }

        [TestMethod]
        public void LongPathParameter_HandlesCorrectly()
        {
            var longPath = new string('a', 300);
            var strategy = new RemoveRecentFileStrategy();

            var script = strategy.GenerateScript(longPath);

            StringAssert.Contains(script, longPath);
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
        [ExpectedException(typeof(ArgumentException))]
        public void EmptyPath_ThrowsArgumentException()
        {
            // Arrange
            var strategy = new PinToFrequentFolderStrategy();

            // Act
            strategy.GenerateScript(string.Empty);
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

        [TestMethod]
        public void GenerateAndSaveStrategyScripts()
        {
            // Arrange
            Assert.IsNotNull(_factory, "Factory should be initialized in TestInitialize");

            var saveDir = Path.Combine(
                Path.GetTempPath(),
                "WincentTemp",
                "Scripts");

            if (Directory.Exists(saveDir))
            {
                Directory.Delete(saveDir, true);
            }
            Directory.CreateDirectory(saveDir);

            var testParams = new Dictionary<PSScript, string?>
            {
                [PSScript.RefreshExplorer] = null,
                [PSScript.QueryQuickAccess] = null,
                [PSScript.QueryRecentFile] = null,
                [PSScript.QueryFrequentFolder] = null,
                [PSScript.RemoveRecentFile] = @"C:\Users\hp\AppData\Local\Temp\WincentTemp",
                [PSScript.PinToFrequentFolder] = @"C:\Users\hp\AppData\Local\Temp\WincentTemp",
                [PSScript.UnpinFromFrequentFolder] = @"C:\Users\hp\AppData\Local\Temp\WincentTemp",
                [PSScript.CheckQueryFeasible] = null,
                [PSScript.CheckPinUnpinFeasible] = null
            };

            try
            {
                foreach (var (scriptType, param) in testParams)
                {
                    var strategy = _factory!.GetStrategy(scriptType);
                    var scriptContent = strategy.GenerateScript(param);

                    byte[] scriptBytes = Encoding.UTF8.GetBytes(scriptContent);
                    byte[] contentWithBom = ScriptExecutor.AddUtf8Bom(scriptBytes);

                    using var tempFile = TempFile.Create(contentWithBom, "ps1");
                    Assert.IsTrue(File.Exists(tempFile.FullPath), $"Temporary file not created: {scriptType}");

                    var permanentPath = Path.Combine(saveDir, $"{scriptType}.ps1");
                    File.Copy(tempFile.FullPath, permanentPath, true);
                    Assert.IsTrue(File.Exists(permanentPath), $"Permanent file not created: {scriptType}");

                    var savedContent = File.ReadAllText(permanentPath);
                    Assert.AreEqual(scriptContent, savedContent, $"File content does not match: {scriptType}");

                    var fileBytes = File.ReadAllBytes(permanentPath);
                    var hasBom = fileBytes.Length >= 3 &&
                                 fileBytes[0] == 0xEF &&
                                 fileBytes[1] == 0xBB &&
                                 fileBytes[2] == 0xBF;
                    Assert.IsTrue(hasBom, $"File should use UTF8-BOM encoding: {scriptType}");
                }

                Console.WriteLine($"Script files saved to: {saveDir}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred during test: {ex.Message}");
                throw;
            }
        }
    }
}
