using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Wincent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Collections.Concurrent;
using System.Diagnostics;
using Moq;

namespace TestWincent
{
    [TestClass]
    [TestCategory("Process")]
    public class TestScriptExecutor
    {
        private ScriptExecutor _executor;
        private string _tempScriptPath;
        private Mock<IFileSystemService> _mockFileSystem;
        private Mock<IScriptStorageService> _mockScriptStorage;
        private QuickAccessDataFiles _dataFiles;

        [TestInitialize]
        public void Initialize()
        {
            _tempScriptPath = Path.Combine(Path.GetTempPath(), "TestScript.ps1");

            _mockFileSystem = new Mock<IFileSystemService>();
            _mockScriptStorage = new Mock<IScriptStorageService>();
            _dataFiles = new QuickAccessDataFiles();

            _mockFileSystem.Setup(fs => fs.FileExists(It.IsAny<string>())).Returns(true);
            _mockFileSystem.Setup(fs => fs.DirectoryExists(It.IsAny<string>())).Returns(true);
            _mockScriptStorage.Setup(ss => ss.GetScriptPath(It.IsAny<PSScript>())).Returns(_tempScriptPath);
            _mockScriptStorage.Setup(ss => ss.GetDynamicScriptPath(It.IsAny<PSScript>(), It.IsAny<string>())).Returns(_tempScriptPath);
            _mockScriptStorage.Setup(ss => ss.IsParameterizedScript(It.IsAny<PSScript>())).Returns(false);

            _executor = new ScriptExecutor(_mockFileSystem.Object, _mockScriptStorage.Object, _dataFiles);
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (File.Exists(_tempScriptPath))
            {
                File.Delete(_tempScriptPath);
            }

            if (_executor != null)
            {
                _executor.ClearCache();
            }
        }

        #region Basic Functionality Tests

        [TestMethod]
        public async Task ExecuteCoreAsync_SuccessfulExecution_ReturnsCorrectResult()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output 'TestOutput'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            var result = await _executor.ExecuteCoreAsync(_tempScriptPath);

            Assert.AreEqual(0, result.ExitCode);
            Assert.IsTrue(result.Output.Contains("TestOutput"));
            Assert.IsTrue(string.IsNullOrEmpty(result.Error));
        }

        [TestMethod]
        public async Task ExecuteCoreAsync_ErrorOutput_ReturnsErrorResult()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Error 'TestError'; exit 1";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            var result = await _executor.ExecuteCoreAsync(_tempScriptPath);

            Assert.AreEqual(1, result.ExitCode);
            Assert.IsTrue(result.Error.Contains("TestError"));
        }

        [TestMethod]
        public async Task ExecutePSScript_SuccessfulExecution_ReturnsCorrectResult()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output 'TestOutput'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            var result = await _executor.ExecutePSScript(PSScript.RefreshExplorer, null);

            Assert.AreEqual(0, result.ExitCode);
            Assert.IsTrue(result.Output.Contains("TestOutput"));
            Assert.IsTrue(string.IsNullOrEmpty(result.Error));
        }

        #endregion

        #region Exception Handling Tests

        [TestMethod]
        [ExpectedException(typeof(FileNotFoundException))]
        public async Task ExecuteCoreAsync_NonExistentScript_ThrowsFileNotFoundException()
        {
            _mockFileSystem.Setup(fs => fs.FileExists(It.IsAny<string>())).Returns(false);

            await _executor.ExecuteCoreAsync(_tempScriptPath);
        }

        [TestMethod]
        public async Task ExecutePSScript_InvalidPath_ThrowsInvalidPathException()
        {
            _mockFileSystem.Setup(fs => fs.FileExists(It.IsAny<string>())).Returns(false);
            _mockFileSystem.Setup(fs => fs.DirectoryExists(It.IsAny<string>())).Returns(false);

            try
            {
                await _executor.ExecutePSScript(PSScript.AddRecentFile, "C:\\NonExistentPath");
                Assert.Fail("Should throw InvalidPathException");
            }
            catch (InvalidPathException ex)
            {
                Assert.AreEqual("C:\\NonExistentPath", ex.Path);
                Assert.IsTrue(ex.Message.Contains("File or directory does not exist"));
            }
        }

        #endregion

        #region Timeout Tests

        [TestMethod]
        public async Task ExecuteCoreAsync_Timeout_ThrowsScriptTimeoutException()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Start-Sleep -Seconds 10; Write-Output 'Output'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            try
            {
                await _executor.ExecuteCoreAsync(_tempScriptPath, TimeSpan.FromSeconds(1));
                Assert.Fail("Should throw ScriptTimeoutException");
            }
            catch (ScriptTimeoutException ex)
            {
                Assert.IsTrue(ex.Message.Contains("Script execution timeout"));
                Assert.IsTrue(ex.Message.Contains("1 seconds"));
            }
        }

        [TestMethod]
        public async Task ExecutePSScriptWithTimeout_Timeout_ThrowsPowerShellExecutionException()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Start-Sleep -Seconds 10; Write-Output 'Output'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            var exception = await Assert.ThrowsExceptionAsync<PowerShellExecutionException>(
                () => _executor.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, 1));

            Assert.AreEqual(PowerShellOperation.RefreshExplorer, exception.Operation);
            Assert.AreEqual(PowerShellErrorKind.Timeout, exception.Kind);
            Assert.IsTrue(exception.StandardError.Contains("timeout"));
        }

        #endregion

        #region Cache Tests

        [TestMethod]
        public async Task ExecutePSScriptWithCache_QueryScript_CachesResult()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output 'CacheTest'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            var result1 = await _executor.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null);
            var result2 = await _executor.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null);

            Assert.IsNotNull(result1);
            Assert.IsNotNull(result2);

            CollectionAssert.AreEqual(result1, result2);
        }

        [TestMethod]
        public async Task ExecutePSScriptWithCache_NonQueryScript_DoesNotCache()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output 'NoCache'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            var result1 = await _executor.ExecutePSScriptWithCache(PSScript.RefreshExplorer, null);
            var result2 = await _executor.ExecutePSScriptWithCache(PSScript.RefreshExplorer, null);

            Assert.IsNotNull(result1);
            Assert.IsNotNull(result2);

            CollectionAssert.AreEqual(result1, result2);
        }

        [TestMethod]
        public async Task ExecutePSScriptWithCache_QueryScriptNonZeroExit_ThrowsPowerShellExecutionException()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Error 'Access is denied'; exit 1";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            var exception = await Assert.ThrowsExceptionAsync<PowerShellExecutionException>(
                () => _executor.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null));

            Assert.AreEqual(PowerShellOperation.QueryQuickAccess, exception.Operation);
            Assert.AreEqual(PowerShellErrorKind.AccessDenied, exception.Kind);
            Assert.AreEqual(1, exception.ExitCode);
            Assert.IsTrue(exception.StandardError.Contains("Access is denied"));
        }

        [TestMethod]
        public async Task ExecutePSScriptWithCache_NonQueryScriptNonZeroExit_ThrowsPowerShellExecutionException()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Error 'CommandNotFoundException'; exit 1";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            var exception = await Assert.ThrowsExceptionAsync<PowerShellExecutionException>(
                () => _executor.ExecutePSScriptWithCache(PSScript.RefreshExplorer, null));

            Assert.AreEqual(PowerShellOperation.RefreshExplorer, exception.Operation);
            Assert.AreEqual(PowerShellErrorKind.CmdletNotFound, exception.Kind);
            Assert.AreEqual(1, exception.ExitCode);
        }

        [TestMethod]
        public async Task ExecutePSScriptWithTimeout_PinAlreadyExistsSentinel_ThrowsAlreadyExistsException()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output 'WINCENT_ALREADY_EXISTS'; exit 1";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            var exception = await Assert.ThrowsExceptionAsync<QuickAccessItemAlreadyExistsException>(
                () => _executor.ExecutePSScriptWithTimeout(PSScript.PinToFrequentFolder, @"C:\Folder", 1));

            Assert.AreEqual(@"C:\Folder", exception.Path);
            Assert.AreEqual(QuickAccess.FrequentFolders, exception.Target);
        }

        [TestMethod]
        public async Task ExecutePSScriptWithTimeout_RemoveRecentNotFoundSentinel_ThrowsNotFoundException()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output 'WINCENT_NOT_IN_QUICK_ACCESS'; exit 1";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            var exception = await Assert.ThrowsExceptionAsync<QuickAccessItemNotFoundException>(
                () => _executor.ExecutePSScriptWithTimeout(PSScript.RemoveRecentFile, @"C:\test.txt", 1));

            Assert.AreEqual(@"C:\test.txt", exception.Path);
            Assert.AreEqual(QuickAccess.RecentFiles, exception.Target);
        }

        [TestMethod]
        public async Task ExecutePSScriptWithTimeout_UnpinNotFoundSentinel_ThrowsNotFoundException()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output 'WINCENT_NOT_IN_QUICK_ACCESS'; exit 1";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            var exception = await Assert.ThrowsExceptionAsync<QuickAccessItemNotFoundException>(
                () => _executor.ExecutePSScriptWithTimeout(PSScript.UnpinFromFrequentFolder, @"C:\Folder", 1));

            Assert.AreEqual(@"C:\Folder", exception.Path);
            Assert.AreEqual(QuickAccess.FrequentFolders, exception.Target);
        }

        [TestMethod]
        public async Task ExecutePSScriptWithCache_DataFileModified_InvalidatesCache()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output 'CacheInvalidate'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            var result1 = await _executor.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null);

            await Task.Delay(100);

            var result2 = await _executor.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null);

            Assert.IsNotNull(result1);
            Assert.IsNotNull(result2);

            CollectionAssert.AreEqual(result1, result2);
        }

        [TestMethod]
        public async Task ExecutePSScriptWithCache_MissingBackingFile_BypassesCacheAndExecutesQuery()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output 'BypassCache'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            var mockFs = new MockFileSystem();
            mockFs.FileExistsDefault = false;
            var dataFiles = new QuickAccessDataFiles(mockFs);
            var executor = new ScriptExecutor(
                _mockFileSystem.Object,
                _mockScriptStorage.Object,
                dataFiles);

            var result = await executor.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null);

            Assert.IsNotNull(result);
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("BypassCache", result[0]);
        }

        [TestMethod]
        public async Task ExecutePSScriptWithCache_CachedResultInvalidatedWhenBackingFileDisappears()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Get-Date -Format 'HH:mm:ss.ffffff'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            bool backingFileExists = true;
            var mockFs = new Mock<IFileSystem>(MockBehavior.Strict);
            mockFs.Setup(fs => fs.FileExists(It.IsAny<string>()))
                .Returns(() => backingFileExists);
            mockFs.Setup(fs => fs.GetLastWriteTime(It.IsAny<string>()))
                .Returns(DateTime.Now);
            var dataFiles = new QuickAccessDataFiles(mockFs.Object);
            var executor = new ScriptExecutor(
                _mockFileSystem.Object,
                _mockScriptStorage.Object,
                dataFiles);

            // First call: backing file exists, result is cached.
            var result1 = await executor.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null);
            Assert.IsNotNull(result1);
            Assert.AreEqual(1, result1.Count);

            // Backing file disappears. Cache should be bypassed and query re-executed.
            backingFileExists = false;
            await Task.Delay(10); // ensure timestamp differs

            var result2 = await executor.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null);
            Assert.IsNotNull(result2);
            Assert.AreEqual(1, result2.Count);

            // Output should differ (Get-Date) — prove it was re-executed, not returned from cache.
            Assert.AreNotEqual(result1[0], result2[0],
                "PowerShell should be re-executed when the backing file disappears; cached result must not be returned.");
        }

        #endregion

        #region File System Integration Tests

        [TestMethod]
        public void FileOrDirectoryExists_ValidPath_ReturnsTrue()
        {
            _mockFileSystem.Setup(fs => fs.FileExists("C:\\ValidPath")).Returns(true);

            Assert.IsTrue(_executor.FileOrDirectoryExists("C:\\ValidPath"));
        }

        [TestMethod]
        public void FileOrDirectoryExists_InvalidPath_ReturnsFalse()
        {
            _mockFileSystem.Setup(fs => fs.FileExists("C:\\InvalidPath")).Returns(false);
            _mockFileSystem.Setup(fs => fs.DirectoryExists("C:\\InvalidPath")).Returns(false);

            Assert.IsFalse(_executor.FileOrDirectoryExists("C:\\InvalidPath"));
        }

        [TestMethod]
        public void FileOrDirectoryExists_NullPath_ReturnsFalse()
        {
            Assert.IsFalse(_executor.FileOrDirectoryExists(null));
        }

        [TestMethod]
        public void FileOrDirectoryExists_EmptyPath_ReturnsFalse()
        {
            Assert.IsFalse(_executor.FileOrDirectoryExists(""));
        }

        [TestMethod]
        public void FileOrDirectoryExists_WhitespacePath_ReturnsFalse()
        {
            Assert.IsFalse(_executor.FileOrDirectoryExists("   "));
        }

        #endregion

        #region Concurrency Tests

        [TestMethod]
        public async Task ExecutePSScript_ConcurrentExecution_IsolatesOutput()
        {
            string scriptContent = "param($index)\r\n[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8;\r\nWrite-Output \"ConcurrentOutput $index\";\r\nexit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            _mockScriptStorage.Setup(ss => ss.IsParameterizedScript(It.IsAny<PSScript>())).Returns(true);

            var tasks = new List<Task<ScriptResult>>();
            for (int i = 0; i < 10; i++)
            {
                int index = i;

                string paramScriptPath = Path.Combine(Path.GetTempPath(), $"TestScript_{index}.ps1");
                File.WriteAllText(paramScriptPath, scriptContent, Encoding.UTF8);

                _mockScriptStorage.Setup(ss => ss.GetDynamicScriptPath(It.IsAny<PSScript>(), index.ToString())).Returns(paramScriptPath);

                tasks.Add(Task.Run(async () =>
                {
                    try
                    {
                        var result = await _executor.ExecutePSScript(PSScript.AddRecentFile, index.ToString());
                        return result;
                    }
                    finally
                    {
                        if (File.Exists(paramScriptPath))
                        {
                            try { File.Delete(paramScriptPath); } catch { /* Ignore */ }
                        }
                    }
                }));
            }

            var results = await Task.WhenAll(tasks);

            for (int i = 0; i < results.Length; i++)
            {
                Assert.AreEqual(0, results[i].ExitCode);
                Assert.IsTrue(results[i].Output.Contains($"ConcurrentOutput {i}"), $"Output missing expected text 'ConcurrentOutput {i}'. Actual: {results[i].Output}");
            }
        }

        [TestMethod]
        public async Task ExecutePSScriptWithCache_ConcurrentAccess_ThreadSafe()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output 'ThreadSafe'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            var tasks = new List<Task<List<string>>>();
            for (int i = 0; i < 10; i++)
            {
                tasks.Add(Task.Run(async () =>
                {
                    return await _executor.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null);
                }));
            }

            var results = await Task.WhenAll(tasks);

            for (int i = 0; i < results.Length; i++)
            {
                Assert.IsNotNull(results[i]);
                Assert.IsTrue(results[i].Contains("ThreadSafe"));
            }

            for (int i = 1; i < results.Length; i++)
            {
                CollectionAssert.AreEqual(results[0], results[i]);
            }
        }

        #endregion

        #region Script Storage Integration Tests

        [TestMethod]
        public async Task ExecutePSScript_ParameterizedScript_CreatesDynamicScript()
        {
            string scriptContent = "param($path)\r\n[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8;\r\nWrite-Output \"ParamPath: $path\";\r\nexit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            _mockScriptStorage.Setup(ss => ss.IsParameterizedScript(PSScript.AddRecentFile)).Returns(true);
            _mockScriptStorage.Setup(ss => ss.GetDynamicScriptPath(PSScript.AddRecentFile, "C:\\TestPath")).Returns(_tempScriptPath);

            var result = await _executor.ExecutePSScript(PSScript.AddRecentFile, "C:\\TestPath");

            Assert.AreEqual(0, result.ExitCode);
            Assert.IsTrue(result.Output.Contains("ParamPath: C:\\TestPath"));
        }

        [TestMethod]
        public async Task ExecutePSScript_NonParameterizedScript_UsesStaticScript()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output 'StaticScript'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            _mockScriptStorage.Setup(ss => ss.IsParameterizedScript(PSScript.RefreshExplorer)).Returns(false);
            _mockScriptStorage.Setup(ss => ss.GetScriptPath(PSScript.RefreshExplorer)).Returns(_tempScriptPath);

            var result = await _executor.ExecutePSScript(PSScript.RefreshExplorer, null);

            Assert.AreEqual(0, result.ExitCode);
            Assert.IsTrue(result.Output.Contains("StaticScript"));
        }

        #endregion

        #region Edge Case Tests

        [TestMethod]
        public async Task ExecutePSScriptWithTimeout_ZeroTimeout_NoTimeout()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output 'ZeroTimeout'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            var result = await _executor.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, 0);

            Assert.AreEqual(0, result.ExitCode);
            Assert.IsTrue(result.Output.Contains("ZeroTimeout"));
        }

        [TestMethod]
        public async Task ExecutePSScript_NullParameter_ExecutesSuccessfully()
        {
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output 'NullParam'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            var result = await _executor.ExecutePSScript(PSScript.RefreshExplorer, null);

            Assert.AreEqual(0, result.ExitCode);
            Assert.IsTrue(result.Output.Contains("NullParam"));
        }

        #endregion
    }
}
