using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
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

        #region Basic Functionality Tests

        [TestMethod]
        public void RefreshExplorerStrategy_GeneratesValidScript()
        {
            var strategy = new RefreshExplorerStrategy();

            var script = strategy.GenerateScript(null);

            StringAssert.Contains(script, "Shell.Application");
            StringAssert.Contains(script, "$_.Refresh()");
            StringAssert.Contains(script, "UTF8");
        }

        [TestMethod]
        public void QueryRecentFileStrategy_ContainsRecentFilesNamespace()
        {
            var strategy = new QueryRecentFileStrategy();

            var script = strategy.GenerateScript(null);

            StringAssert.Contains(script, ShellNamespaces.QuickAccess);
            StringAssert.Contains(script, "$_.IsFolder -eq $false");
        }

        [TestMethod]
        public void QueryFrequentFolderStrategy_UsesCorrectNamespace()
        {
            var strategy = new QueryFrequentFolderStrategy();

            var script = strategy.GenerateScript(null);

            StringAssert.Contains(script, ShellNamespaces.FrequentFolders);
            StringAssert.Contains(script, "ForEach-Object");
        }

        [TestMethod]
        public void QueryQuickAccessStrategy_GeneratesCorrectScript()
        {
            var strategy = new QueryQuickAccessStrategy();

            var script = strategy.GenerateScript(null);

            StringAssert.Contains(script, ShellNamespaces.QuickAccess);
            StringAssert.Contains(script, "ForEach-Object { $_.Path }");
        }

        [TestMethod]
        public void AddRecentFileStrategy_GeneratesCorrectScript()
        {
            var strategy = new AddRecentFileStrategy();
            string testPath = @"C:\test\file.txt";

            var script = strategy.GenerateScript(testPath);

            StringAssert.Contains(script, "UTF8");
            StringAssert.Contains(script, "Shell.Application");
            StringAssert.Contains(script, $"Write-Output '{testPath}'");
        }

        #endregion

        #region Parameter Validation Tests

        [TestMethod]
        public void RemoveRecentFileStrategy_ValidatesParameter()
        {
            var strategy = new RemoveRecentFileStrategy();

            var ex = Assert.ThrowsException<ArgumentException>(
                () => strategy.GenerateScript(null));

            Assert.AreEqual("Valid file path parameter required", ex.Message);
        }

        [TestMethod]
        public void PinToFrequentFolder_ValidatesPathParameter()
        {
            var strategy = new PinToFrequentFolderStrategy();

            var ex = Assert.ThrowsException<ArgumentException>(
                () => strategy.GenerateScript("  "));

            Assert.AreEqual("Valid file path parameter required", ex.Message);
        }

        [TestMethod]
        public void UnpinFromFrequentFolder_ValidatesPathParameter()
        {
            var strategy = new UnpinFromFrequentFolderStrategy();

            var ex = Assert.ThrowsException<ArgumentException>(
                () => strategy.GenerateScript(string.Empty));

            Assert.AreEqual("Valid file path parameter required", ex.Message);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void EmptyPath_ThrowsArgumentException()
        {
            var strategy = new PinToFrequentFolderStrategy();

            strategy.GenerateScript(string.Empty);
        }

        [TestMethod]
        public void AddRecentFile_ValidatesPathParameter()
        {
            var strategy = new AddRecentFileStrategy();

            var ex = Assert.ThrowsException<ArgumentException>(
                () => strategy.GenerateScript(null));

            Assert.AreEqual("Valid file path parameter required", ex.Message);
        }

        #endregion

        #region Factory and Mapping Tests

        [TestMethod]
        public void Factory_ReturnsCorrectStrategies()
        {
            Assert.IsInstanceOfType(
                _factory.GetStrategy(PSScript.RefreshExplorer),
                typeof(RefreshExplorerStrategy));

            Assert.IsInstanceOfType(
                _factory.GetStrategy(PSScript.QueryRecentFile),
                typeof(QueryRecentFileStrategy));

            Assert.IsInstanceOfType(
                _factory.GetStrategy(PSScript.QueryFrequentFolder),
                typeof(QueryFrequentFolderStrategy));

            Assert.IsInstanceOfType(
                _factory.GetStrategy(PSScript.QueryQuickAccess),
                typeof(QueryQuickAccessStrategy));

            Assert.IsInstanceOfType(
                _factory.GetStrategy(PSScript.AddRecentFile),
                typeof(AddRecentFileStrategy));

            Assert.IsInstanceOfType(
                _factory.GetStrategy(PSScript.RemoveRecentFile),
                typeof(RemoveRecentFileStrategy));

            Assert.IsInstanceOfType(
                _factory.GetStrategy(PSScript.PinToFrequentFolder),
                typeof(PinToFrequentFolderStrategy));

            Assert.IsInstanceOfType(
                _factory.GetStrategy(PSScript.UnpinFromFrequentFolder),
                typeof(UnpinFromFrequentFolderStrategy));

            Assert.IsInstanceOfType(
                _factory.GetStrategy(PSScript.EmptyPinnedFolders),
                typeof(EmptyPinnedFoldersStrategy));
        }

        [TestMethod]
        public void Factory_ThrowsOnUnsupportedMethod()
        {
            var invalidMethod = (PSScript)100;

            var ex = Assert.ThrowsException<NotSupportedException>(
                () => _factory.GetStrategy(invalidMethod));

            StringAssert.Contains(ex.Message, "Unsupported script type: 100");
        }

        [TestMethod]
        public void ToPowerShellOperation_MapsAllSupportedScripts()
        {
            Assert.AreEqual(PowerShellOperation.RefreshExplorer, PSScript.RefreshExplorer.ToPowerShellOperation());
            Assert.AreEqual(PowerShellOperation.QueryQuickAccess, PSScript.QueryQuickAccess.ToPowerShellOperation());
            Assert.AreEqual(PowerShellOperation.QueryRecentFiles, PSScript.QueryRecentFile.ToPowerShellOperation());
            Assert.AreEqual(PowerShellOperation.QueryFrequentFolders, PSScript.QueryFrequentFolder.ToPowerShellOperation());
            Assert.AreEqual(PowerShellOperation.AddRecentFile, PSScript.AddRecentFile.ToPowerShellOperation());
            Assert.AreEqual(PowerShellOperation.RemoveRecentFile, PSScript.RemoveRecentFile.ToPowerShellOperation());
            Assert.AreEqual(PowerShellOperation.PinFrequentFolder, PSScript.PinToFrequentFolder.ToPowerShellOperation());
            Assert.AreEqual(PowerShellOperation.UnpinFrequentFolder, PSScript.UnpinFromFrequentFolder.ToPowerShellOperation());
            Assert.AreEqual(PowerShellOperation.ClearPinnedFolders, PSScript.EmptyPinnedFolders.ToPowerShellOperation());
        }

        [TestMethod]
        public void ToPowerShellOperation_ThrowsOnUnsupportedScript()
        {
            var ex = Assert.ThrowsException<NotSupportedException>(
                () => ((PSScript)100).ToPowerShellOperation());

            StringAssert.Contains(ex.Message, "Unsupported script type: 100");
        }

        [TestMethod]
        public void Factory_ReturnsNewStrategyInstancesEachTime()
        {
            var instance1 = _factory.GetStrategy(PSScript.RefreshExplorer);
            var instance2 = _factory.GetStrategy(PSScript.RefreshExplorer);

            Assert.AreNotSame(instance1, instance2, "Factory should return new strategy instances each time");
        }

        #endregion

        #region Path Handling Tests

        [TestMethod]
        public void PathWithSpaces_HandlesCorrectly()
        {
            var strategy = new PinToFrequentFolderStrategy();
            var path = @"C:\Program Files\Test Folder";

            var script = strategy.GenerateScript(path);

            StringAssert.Contains(script, path);
        }

        [TestMethod]
        public void PathWithUnicodeCharacters_HandlesCorrectly()
        {
            var strategy = new RemoveRecentFileStrategy();
            var path = @"C:\test\file.txt";

            var script = strategy.GenerateScript(path);

            StringAssert.Contains(script, path);
        }

        [TestMethod]
        public void PathWithDots_HandlesCorrectly()
        {
            var strategy = new UnpinFromFrequentFolderStrategy();
            var path = @"C:\folder.name\file..txt";

            var script = strategy.GenerateScript(path);

            StringAssert.Contains(script, path);
        }

        [TestMethod]
        public void PathWithValidSpecialCharacters_HandlesCorrectly()
        {
            var strategy = new RemoveRecentFileStrategy();
            var path = @"C:\folder-name\file (1)_[test]~$temp.txt";

            var script = strategy.GenerateScript(path);

            StringAssert.Contains(script, path);
        }

        [TestMethod]
        public void PathWithSingleQuotes_EscapesCorrectly()
        {
            var strategy = new RemoveRecentFileStrategy();
            var path = @"C:\folder\file's_name.txt";

            var script = strategy.GenerateScript(path);

            StringAssert.Contains(script, @"C:\folder\file''s_name.txt");
            Assert.IsFalse(script.Contains(@"file's_name"), "Single quote should be escaped");
            Assert.IsTrue(script.Contains(@"file''s_name"), "Single quote should be replaced with two single quotes");
        }

        #endregion

        #region Script Content Tests

        [TestMethod]
        public void ValidPinOperation_GeneratesCorrectVerb()
        {
            var strategy = new PinToFrequentFolderStrategy();
            const string testPath = @"C:\test";

            var script = strategy.GenerateScript(testPath);

            StringAssert.Contains(script, "InvokeVerb('pintohome')");
            StringAssert.Contains(script, testPath);
        }

        [TestMethod]
        public void ValidUnpinOperation_GeneratesCorrectVerb()
        {
            var strategy = new UnpinFromFrequentFolderStrategy();
            const string testPath = @"C:\test";

            var script = strategy.GenerateScript(testPath);

            StringAssert.Contains(script, "InvokeVerb('unpinfromhome')");
            StringAssert.Contains(script, testPath);
        }

        [TestMethod]
        public void UnpinOperation_DoesNotReferenceUndefinedScriptPath()
        {
            var strategy = new UnpinFromFrequentFolderStrategy();
            const string testPath = @"C:\test";

            var script = strategy.GenerateScript(testPath);

            Assert.IsFalse(script.Contains("$scriptPath"), "Unpin script should not reference undefined $scriptPath variable");
            StringAssert.Contains(script, $@"Namespace('{testPath}')");
            StringAssert.Contains(script, $@"Path -eq '{testPath}'");
        }

        [TestMethod]
        public void UnpinStrategyContainsOSVersionCheck()
        {
            var strategy = new UnpinFromFrequentFolderStrategy();
            const string testPath = @"C:\test";

            var script = strategy.GenerateScript(testPath);

            StringAssert.Contains(script, "Get-CimInstance -Class Win32_OperatingSystem");
            StringAssert.Contains(script, "Windows 11");
            StringAssert.Contains(script, "if ($isWin11)");
        }

        [TestMethod]
        public void UnpinStrategy_VerifiesRemovalAndPinsUnpinnedFrequentFoldersBeforeUnpin()
        {
            var strategy = new UnpinFromFrequentFolderStrategy();
            const string testPath = @"C:\test";

            var script = strategy.GenerateScript(testPath);

            StringAssert.Contains(script, "$target.InvokeVerb('unpinfromhome')");
            StringAssert.Contains(script, $@"$shellApplication.Namespace('{testPath}').Self.InvokeVerb('pintohome')");
            StringAssert.Contains(script, "Start-Sleep -Milliseconds 1000");
            StringAssert.Contains(script, "if ($null -eq $target) { return }");
            StringAssert.Contains(script, "Failed to remove frequent folder");
        }

        #endregion

        #region File Generation Tests

        [TestMethod]
        public void GenerateAndSaveStrategyScripts()
        {
            var saveDir = Path.Combine(
                Path.GetTempPath(),
                "WincentTemp",
                "Scripts");

            if (Directory.Exists(saveDir))
            {
                Directory.Delete(saveDir, true);
            }
            Directory.CreateDirectory(saveDir);

            var testParams = new Dictionary<PSScript, string>
            {
                [PSScript.RefreshExplorer] = null,
                [PSScript.QueryQuickAccess] = null,
                [PSScript.QueryRecentFile] = null,
                [PSScript.QueryFrequentFolder] = null,
                [PSScript.AddRecentFile] = @"C:\test\file.txt",
                [PSScript.RemoveRecentFile] = @"C:\test\file.txt",
                [PSScript.PinToFrequentFolder] = @"C:\test\folder",
                [PSScript.UnpinFromFrequentFolder] = @"C:\test\folder",
                [PSScript.EmptyPinnedFolders] = null
            };

            try
            {
                foreach (var pair in testParams)
                {
                    var scriptType = pair.Key;
                    var param = pair.Value;
                    var strategy = _factory.GetStrategy(scriptType);
                    var scriptContent = strategy.GenerateScript(param);

                    byte[] scriptBytes = Encoding.UTF8.GetBytes(scriptContent);
                    byte[] contentWithBom = ScriptStorage.AddUtf8Bom(scriptBytes);

                    using (var tempFile = TempFile.Create(contentWithBom, "ps1"))
                    {
                        Assert.IsTrue(File.Exists(tempFile.FullPath), $"Temp file not created: {scriptType}");

                        var permanentPath = Path.Combine(saveDir, $"{scriptType}.ps1");
                        File.Copy(tempFile.FullPath, permanentPath, true);
                        Assert.IsTrue(File.Exists(permanentPath), $"Permanent file not created: {scriptType}");

                        var savedContent = File.ReadAllText(permanentPath);
                        Assert.AreEqual(scriptContent, savedContent, $"File content mismatch: {scriptType}");

                        var fileBytes = File.ReadAllBytes(permanentPath);
                        var hasBom = fileBytes.Length >= 3 &&
                                     fileBytes[0] == 0xEF &&
                                     fileBytes[1] == 0xBB &&
                                     fileBytes[2] == 0xBF;
                        Assert.IsTrue(hasBom, $"File should use UTF8-BOM encoding: {scriptType}");
                    }
                }

                Console.WriteLine($"Script files saved to: {saveDir}");
            }
            finally
            {
                if (Directory.Exists(saveDir))
                {
                    try
                    {
                        Directory.Delete(saveDir, true);
                    }
                    catch
                    {
                        // Ignore cleanup errors
                    }
                }
            }
        }

        #endregion

        #region Edge Case Tests

        [TestMethod]
        public void LongPathParameter_HandlesCorrectly()
        {
            var strategy = new PinToFrequentFolderStrategy();
            var path = @"C:\" + new string('a', 200) + @"\test.txt";

            var script = strategy.GenerateScript(path);

            StringAssert.Contains(script, path);
        }

        [TestMethod]
        public void MultipleStrategies_GenerateDifferentScripts()
        {
            var strategy1 = new QueryRecentFileStrategy();
            var strategy2 = new QueryFrequentFolderStrategy();

            var script1 = strategy1.GenerateScript(null);
            var script2 = strategy2.GenerateScript(null);

            Assert.AreNotEqual(script1, script2, "Different strategies should generate different scripts");
        }

        [TestMethod]
        public void CommonConstants_ShouldBeConsistent()
        {
            var quickAccessStrategy = new QueryQuickAccessStrategy();
            var recentFileStrategy = new QueryRecentFileStrategy();

            var script1 = quickAccessStrategy.GenerateScript(null);
            var script2 = recentFileStrategy.GenerateScript(null);

            Assert.IsTrue(script1.Contains(ShellNamespaces.QuickAccess));
            Assert.IsTrue(script2.Contains(ShellNamespaces.QuickAccess));
            Assert.IsTrue(script1.Contains("UTF8") && script2.Contains("UTF8"));
        }

        [TestMethod]
        public void ScriptGenerationException_CanBeCreated()
        {
            var innerException = new IOException("Inner error");
            var exception = new ScriptGenerationException("Script generation failed", innerException);

            Assert.AreEqual("Script generation failed", exception.Message);
            Assert.AreEqual("Inner error", exception.InnerException.Message);
        }

        #endregion

        #region Single Quote Escape Tests

        [TestMethod]
        public void AddRecentFile_EscapesSingleQuotes()
        {
            var strategy = new AddRecentFileStrategy();
            var path = @"C:\Users\John's\Documents\file.txt";

            var script = strategy.GenerateScript(path);

            Assert.IsFalse(script.Contains(@"John's"), "Single quote was not escaped");
            Assert.IsTrue(script.Contains(@"John''s"), "Single quote should be replaced with two single quotes");
        }

        [TestMethod]
        public void PinToFrequentFolder_EscapesSingleQuotes()
        {
            var strategy = new PinToFrequentFolderStrategy();
            var path = @"C:\User's\Data";

            var script = strategy.GenerateScript(path);

            Assert.IsFalse(script.Contains(@"User's"), "Single quote was not escaped");
            Assert.IsTrue(script.Contains(@"User''s"), "Single quote should be replaced with two single quotes");
        }

        [TestMethod]
        public void UnpinFromFrequentFolder_EscapesSingleQuotes()
        {
            var strategy = new UnpinFromFrequentFolderStrategy();
            var path = @"C:\My's Documents\Data";

            var script = strategy.GenerateScript(path);

            Assert.IsFalse(script.Contains(@"My's"), "Single quote was not escaped");
            Assert.IsTrue(script.Contains(@"My''s"), "Single quote should be replaced with two single quotes");
        }

        #endregion

        [TestMethod]
        public void EscapePowerShellString_HandlesAllCases()
        {
            var methodInfo = typeof(PSScriptStrategyBase).GetMethod(
                "EscapePowerShellString",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);

            Assert.IsNotNull(methodInfo, "EscapePowerShellString method should exist");

            var testCases = new Dictionary<string, string>
            {
                { "", "" },
                { "NoQuotes", "NoQuotes" },
                { "Single'Quote", "Single''Quote" },
                { "Multiple'Quotes'Here", "Multiple''Quotes''Here" },
                { "'StartAndEnd'", "''StartAndEnd''" },
                { "''", "''''" },
                { "'", "''" }
            };

            foreach (var testCase in testCases)
            {
                var actual = methodInfo.Invoke(null, new object[] { testCase.Key });

                Assert.AreEqual(
                    testCase.Value,
                    actual,
                    $"Input '{testCase.Key ?? "null"}' should escape to '{testCase.Value ?? "null"}'");
            }
        }
    }
}
