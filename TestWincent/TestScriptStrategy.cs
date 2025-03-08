using Microsoft.VisualStudio.TestTools.UnitTesting;
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
            StringAssert.Contains(script, ShellNamespaces.RecentFiles);
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
        public void CheckQueryFeasible_ContainsTimeoutHandler()
        {
            // Arrange
            var strategy = new CheckQueryFeasibleStrategy();

            // Act
            var script = strategy.GenerateScript(null);

            // Assert
            StringAssert.Contains(script, "WaitForExit(5 * 1000)");
            StringAssert.Contains(script, "Write-Error \"Operation timeout (5 seconds)\"");
            StringAssert.Contains(script, "Start-Process powershell");
        }

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
            StringAssert.Contains(script, "WaitForExit(5 * 1000)");
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
    }
}
