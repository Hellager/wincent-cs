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

            // 创建模拟文件系统服务
            _mockFileSystem = new Mock<IFileSystemService>();

            // 创建模拟脚本存储服务
            _mockScriptStorage = new Mock<IScriptStorageService>();

            // 创建实际的 QuickAccessDataFiles 实例
            _dataFiles = new QuickAccessDataFiles();

            // 设置默认行为
            _mockFileSystem.Setup(fs => fs.FileExists(It.IsAny<string>())).Returns(true);
            _mockFileSystem.Setup(fs => fs.DirectoryExists(It.IsAny<string>())).Returns(true);
            _mockScriptStorage.Setup(ss => ss.GetScriptPath(It.IsAny<PSScript>())).Returns(_tempScriptPath);
            _mockScriptStorage.Setup(ss => ss.GetDynamicScriptPath(It.IsAny<PSScript>(), It.IsAny<string>())).Returns(_tempScriptPath);
            _mockScriptStorage.Setup(ss => ss.IsParameterizedScript(It.IsAny<PSScript>())).Returns(false);

            // 创建 ScriptExecutor 实例
            _executor = new ScriptExecutor(_mockFileSystem.Object, _mockScriptStorage.Object, _dataFiles);
        }

        [TestCleanup]
        public void Cleanup()
        {
            // 清理临时文件
            if (File.Exists(_tempScriptPath))
            {
                File.Delete(_tempScriptPath);
            }

            // 清理缓存
            if (_executor != null)
            {
                _executor.ClearCache();
            }
        }

        #region 基础功能测试

        [TestMethod]
        public async Task ExecuteCoreAsync_SuccessfulExecution_ReturnsCorrectResult()
        {
            // 准备测试脚本
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output '测试输出'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            // 执行脚本
            var result = await _executor.ExecuteCoreAsync(_tempScriptPath);

            // 验证结果
            Assert.AreEqual(0, result.ExitCode);
            Assert.IsTrue(result.Output.Contains("测试输出"));
            Assert.IsTrue(string.IsNullOrEmpty(result.Error));
        }

        [TestMethod]
        public async Task ExecuteCoreAsync_ErrorOutput_ReturnsErrorResult()
        {
            // 准备测试脚本
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Error '测试错误'; exit 1";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            // 执行脚本
            var result = await _executor.ExecuteCoreAsync(_tempScriptPath);

            // 验证结果
            Assert.AreEqual(1, result.ExitCode);
            Assert.IsTrue(result.Error.Contains("测试错误"));
        }

        [TestMethod]
        public async Task ExecutePSScript_SuccessfulExecution_ReturnsCorrectResult()
        {
            // 准备测试脚本
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output '测试输出'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            // 执行脚本
            var result = await _executor.ExecutePSScript(PSScript.RefreshExplorer, null);

            // 验证结果
            Assert.AreEqual(0, result.ExitCode);
            Assert.IsTrue(result.Output.Contains("测试输出"));
            Assert.IsTrue(string.IsNullOrEmpty(result.Error));
        }

        #endregion

        #region 异常处理测试

        [TestMethod]
        [ExpectedException(typeof(FileNotFoundException))]
        public async Task ExecuteCoreAsync_NonExistentScript_ThrowsFileNotFoundException()
        {
            // 设置文件系统服务返回文件不存在
            _mockFileSystem.Setup(fs => fs.FileExists(It.IsAny<string>())).Returns(false);

            // 执行脚本
            await _executor.ExecuteCoreAsync(_tempScriptPath);
        }

        [TestMethod]
        public async Task ExecutePSScript_InvalidPath_ThrowsInvalidPathException()
        {
            // 设置文件系统服务返回文件不存在
            _mockFileSystem.Setup(fs => fs.FileExists(It.IsAny<string>())).Returns(false);
            _mockFileSystem.Setup(fs => fs.DirectoryExists(It.IsAny<string>())).Returns(false);

            // 执行脚本并捕获异常
            try
            {
                await _executor.ExecutePSScript(PSScript.AddRecentFile, "C:\\NonExistentPath");
                Assert.Fail("应该抛出 InvalidPathException 异常");
            }
            catch (InvalidPathException ex)
            {
                Assert.AreEqual("C:\\NonExistentPath", ex.Path);
                Assert.IsTrue(ex.Message.Contains("File or directory does not exist"));
            }
        }

        #endregion

        #region 超时机制测试

        [TestMethod]
        public async Task ExecuteCoreAsync_Timeout_ThrowsScriptTimeoutException()
        {
            // 准备长时间运行的测试脚本
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Start-Sleep -Seconds 10; Write-Output '测试输出'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            // 执行脚本并设置超时
            try
            {
                await _executor.ExecuteCoreAsync(_tempScriptPath, TimeSpan.FromSeconds(1));
                Assert.Fail("应该抛出 ScriptTimeoutException 异常");
            }
            catch (ScriptTimeoutException ex)
            {
                Assert.IsTrue(ex.Message.Contains("Script execution timeout"));
                Assert.IsTrue(ex.Message.Contains("1 seconds"));
            }
        }

        [TestMethod]
        public async Task ExecutePSScriptWithTimeout_Timeout_ReturnsTimeoutResult()
        {
            // 准备长时间运行的测试脚本
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Start-Sleep -Seconds 10; Write-Output '测试输出'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            // 执行脚本并设置超时
            var result = await _executor.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, 1);

            // 验证结果
            Assert.AreEqual(-1, result.ExitCode);
            Assert.IsTrue(result.Error.Contains("timeout"));
        }

        #endregion

        #region 缓存功能测试

        [TestMethod]
        public async Task ExecutePSScriptWithCache_QueryScript_CachesResult()
        {
            // 准备测试脚本
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output '测试输出'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            // 执行查询脚本两次
            var result1 = await _executor.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null);
            var result2 = await _executor.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null);

            // 验证结果
            Assert.IsNotNull(result1);
            Assert.IsNotNull(result2);

            // 由于缓存机制，第二次调用应该返回相同的结果
            CollectionAssert.AreEqual(result1, result2);
        }

        [TestMethod]
        public async Task ExecutePSScriptWithCache_NonQueryScript_DoesNotCache()
        {
            // 准备测试脚本
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output '测试输出'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            // 执行非查询脚本两次
            var result1 = await _executor.ExecutePSScriptWithCache(PSScript.RefreshExplorer, null);
            var result2 = await _executor.ExecutePSScriptWithCache(PSScript.RefreshExplorer, null);

            // 验证结果
            Assert.IsNotNull(result1);
            Assert.IsNotNull(result2);

            // 非查询脚本不应该被缓存，但结果应该相同
            CollectionAssert.AreEqual(result1, result2);
        }

        [TestMethod]
        public async Task ExecutePSScriptWithCache_DataFileModified_InvalidatesCache()
        {
            // 准备测试脚本
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output '测试输出'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            // 执行查询脚本两次
            var result1 = await _executor.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null);

            // 等待一段时间以确保修改时间不同
            await Task.Delay(100);

            var result2 = await _executor.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null);

            // 验证结果
            Assert.IsNotNull(result1);
            Assert.IsNotNull(result2);

            // 由于数据文件修改时间变化，缓存应该失效，但结果应该相同
            CollectionAssert.AreEqual(result1, result2);
        }

        #endregion

        #region 文件系统集成测试

        [TestMethod]
        public void FileOrDirectoryExists_ValidPath_ReturnsTrue()
        {
            // 设置文件系统服务返回文件存在
            _mockFileSystem.Setup(fs => fs.FileExists("C:\\ValidPath")).Returns(true);

            // 验证结果
            Assert.IsTrue(_executor.FileOrDirectoryExists("C:\\ValidPath"));
        }

        [TestMethod]
        public void FileOrDirectoryExists_InvalidPath_ReturnsFalse()
        {
            // 设置文件系统服务返回文件不存在
            _mockFileSystem.Setup(fs => fs.FileExists("C:\\InvalidPath")).Returns(false);
            _mockFileSystem.Setup(fs => fs.DirectoryExists("C:\\InvalidPath")).Returns(false);

            // 验证结果
            Assert.IsFalse(_executor.FileOrDirectoryExists("C:\\InvalidPath"));
        }

        [TestMethod]
        public void FileOrDirectoryExists_NullPath_ReturnsFalse()
        {
            // 验证结果
            Assert.IsFalse(_executor.FileOrDirectoryExists(null));
        }

        [TestMethod]
        public void FileOrDirectoryExists_EmptyPath_ReturnsFalse()
        {
            // 验证结果
            Assert.IsFalse(_executor.FileOrDirectoryExists(""));
        }

        [TestMethod]
        public void FileOrDirectoryExists_WhitespacePath_ReturnsFalse()
        {
            // 验证结果
            Assert.IsFalse(_executor.FileOrDirectoryExists("   "));
        }

        #endregion

        #region 并发执行测试

        [TestMethod]
        public async Task ExecutePSScript_ConcurrentExecution_IsolatesOutput()
        {
            // 准备测试脚本 - 使用参数化的方式
            string scriptContent = "param($index)\r\n[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8;\r\nWrite-Output \"测试输出 $index\";\r\nexit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            // 设置脚本存储服务，使其返回我们的测试脚本路径，并标记为参数化脚本
            _mockScriptStorage.Setup(ss => ss.IsParameterizedScript(It.IsAny<PSScript>())).Returns(true);

            // 创建并发任务
            var tasks = new List<Task<ScriptResult>>();
            for (int i = 0; i < 10; i++)
            {
                int index = i; // 捕获循环变量

                // 为每个任务创建唯一的脚本路径，避免并发冲突
                string paramScriptPath = Path.Combine(Path.GetTempPath(), $"TestScript_{index}.ps1");
                File.WriteAllText(paramScriptPath, scriptContent, Encoding.UTF8);

                // 设置每个特定参数的脚本路径
                _mockScriptStorage.Setup(ss => ss.GetDynamicScriptPath(It.IsAny<PSScript>(), index.ToString())).Returns(paramScriptPath);

                // 添加任务
                tasks.Add(Task.Run(async () =>
                {
                    try
                    {
                        // 执行脚本，使用索引作为参数
                        var result = await _executor.ExecutePSScript(PSScript.AddRecentFile, index.ToString());
                        return result;
                    }
                    finally
                    {
                        // 清理临时文件
                        if (File.Exists(paramScriptPath))
                        {
                            try { File.Delete(paramScriptPath); } catch { /* 忽略删除失败 */ }
                        }
                    }
                }));
            }

            // 等待所有任务完成
            var results = await Task.WhenAll(tasks);

            // 验证结果
            for (int i = 0; i < results.Length; i++)
            {
                Assert.AreEqual(0, results[i].ExitCode);
                Assert.IsTrue(results[i].Output.Contains($"测试输出 {i}"), $"输出不包含预期的文本 '测试输出 {i}'。实际输出: {results[i].Output}");
            }
        }

        [TestMethod]
        public async Task ExecutePSScriptWithCache_ConcurrentAccess_ThreadSafe()
        {
            // 准备测试脚本
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output '测试输出'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            // 创建并发任务
            var tasks = new List<Task<List<string>>>();
            for (int i = 0; i < 10; i++)
            {
                tasks.Add(Task.Run(async () =>
                {
                    return await _executor.ExecutePSScriptWithCache(PSScript.QueryQuickAccess, null);
                }));
            }

            // 等待所有任务完成
            var results = await Task.WhenAll(tasks);

            // 验证结果
            for (int i = 0; i < results.Length; i++)
            {
                Assert.IsNotNull(results[i]);
                Assert.IsTrue(results[i].Contains("测试输出"));
            }

            // 验证所有结果相同（缓存应该确保这一点）
            for (int i = 1; i < results.Length; i++)
            {
                CollectionAssert.AreEqual(results[0], results[i]);
            }
        }

        #endregion

        #region 脚本存储集成测试

        [TestMethod]
        public async Task ExecutePSScript_ParameterizedScript_CreatesDynamicScript()
        {
            // 准备测试脚本
            string scriptContent = "param($path)\r\n[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8;\r\nWrite-Output \"参数路径: $path\";\r\nexit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            // 设置脚本存储服务
            _mockScriptStorage.Setup(ss => ss.IsParameterizedScript(PSScript.AddRecentFile)).Returns(true);
            _mockScriptStorage.Setup(ss => ss.GetDynamicScriptPath(PSScript.AddRecentFile, "C:\\TestPath")).Returns(_tempScriptPath);

            // 执行参数化脚本
            var result = await _executor.ExecutePSScript(PSScript.AddRecentFile, "C:\\TestPath");

            // 验证结果
            Assert.AreEqual(0, result.ExitCode);
            Assert.IsTrue(result.Output.Contains("参数路径: C:\\TestPath"));
        }

        [TestMethod]
        public async Task ExecutePSScript_NonParameterizedScript_UsesStaticScript()
        {
            // 准备测试脚本
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output '静态脚本测试'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            // 设置脚本存储服务
            _mockScriptStorage.Setup(ss => ss.IsParameterizedScript(PSScript.RefreshExplorer)).Returns(false);
            _mockScriptStorage.Setup(ss => ss.GetScriptPath(PSScript.RefreshExplorer)).Returns(_tempScriptPath);

            // 执行非参数化脚本
            var result = await _executor.ExecutePSScript(PSScript.RefreshExplorer, null);

            // 验证结果
            Assert.AreEqual(0, result.ExitCode);
            Assert.IsTrue(result.Output.Contains("静态脚本测试"));
        }

        #endregion

        #region 边界条件测试

        [TestMethod]
        public async Task ExecutePSScriptWithTimeout_ZeroTimeout_NoTimeout()
        {
            // 准备测试脚本
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output '测试输出'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            // 执行脚本并设置超时为0
            var result = await _executor.ExecutePSScriptWithTimeout(PSScript.RefreshExplorer, null, 0);

            // 验证结果
            Assert.AreEqual(0, result.ExitCode);
            Assert.IsTrue(result.Output.Contains("测试输出"));
        }

        [TestMethod]
        public async Task ExecutePSScript_NullParameter_ExecutesSuccessfully()
        {
            // 准备测试脚本
            string scriptContent = "[System.Console]::OutputEncoding=[System.Text.Encoding]::UTF8; Write-Output '测试输出'; exit 0";
            File.WriteAllText(_tempScriptPath, scriptContent, Encoding.UTF8);

            // 执行脚本并传入null参数
            var result = await _executor.ExecutePSScript(PSScript.RefreshExplorer, null);

            // 验证结果
            Assert.AreEqual(0, result.ExitCode);
            Assert.IsTrue(result.Output.Contains("测试输出"));
        }

        #endregion
    }
}
