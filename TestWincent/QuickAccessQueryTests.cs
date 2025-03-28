using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Security;
using System.Threading.Tasks;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class TestQuickAccessQuery
    {
        private MockScriptExecutorService _mockExecutor;
        private MockExecutionFeasibleService _mockFeasible;

        [TestInitialize]
        public void Initialize()
        {
            // 创建模拟服务
            _mockExecutor = new MockScriptExecutorService();
            _mockFeasible = new MockExecutionFeasibleService();

            // 替换服务实现
            QuickAccessQuery.SetServices(_mockExecutor, _mockFeasible);

            // 设置默认行为
            _mockFeasible.SetScriptFeasible(true);
            _mockFeasible.SetIsAdministrator(false);
        }

        [TestCleanup]
        public void Cleanup()
        {
            // 恢复默认服务
            QuickAccessQuery.ResetServices();
        }

        #region 基本功能测试

        [TestMethod]
        public async Task GetRecentFilesAsync_Success_ReturnsCorrectFiles()
        {
            // Arrange
            var mockFiles = new List<string>
            {
                @"C:\Users\Test\Documents\file1.txt",
                @"C:\Users\Test\Documents\file2.docx",
                @"C:\Users\Test\Pictures\image.jpg"
            };

            _mockExecutor.SetupResult(PSScript.QueryRecentFile, 0, string.Join(Environment.NewLine, mockFiles), "");

            // Act
            var result = await QuickAccessQuery.GetRecentFilesAsync();

            // Assert
            CollectionAssert.AreEqual(mockFiles, result, "返回的最近文件应该与模拟数据匹配");
        }

        [TestMethod]
        public async Task GetFrequentFoldersAsync_Success_ReturnsCorrectFolders()
        {
            // Arrange
            var mockFolders = new List<string>
            {
                @"C:\Users\Test\Documents",
                @"C:\Users\Test\Downloads",
                @"C:\Program Files"
            };

            _mockExecutor.SetupResult(PSScript.QueryFrequentFolder, 0, string.Join(Environment.NewLine, mockFolders), "");

            // Act
            var result = await QuickAccessQuery.GetFrequentFoldersAsync();

            // Assert
            CollectionAssert.AreEqual(mockFolders, result, "返回的常用文件夹应该与模拟数据匹配");
        }

        [TestMethod]
        public async Task GetAllItemsAsync_Success_ReturnsCorrectItems()
        {
            // Arrange
            var mockItems = new List<string>
            {
                @"C:\Users\Test\Documents",
                @"C:\Users\Test\Pictures",
                @"C:\Users\Test\Desktop\Project"
            };

            _mockExecutor.SetupResult(PSScript.QueryQuickAccess, 0, string.Join(Environment.NewLine, mockItems), "");

            // Act
            var result = await QuickAccessQuery.GetAllItemsAsync();

            // Assert
            CollectionAssert.AreEqual(mockItems, result, "返回的所有项目应该与模拟数据匹配");
        }

        #endregion

        #region 异常处理测试

        [TestMethod]
        [ExpectedException(typeof(SecurityException))]
        public async Task GetRecentFilesAsync_ScriptNotFeasible_ThrowsSecurityException()
        {
            // Arrange
            _mockFeasible.SetScriptFeasible(false);

            // Act - 应该抛出 SecurityException
            await QuickAccessQuery.GetRecentFilesAsync();
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public async Task GetFrequentFoldersAsync_RepeatedFailures_ThrowsInvalidOperationException()
        {
            // Arrange
            _mockExecutor.SetupResult(PSScript.QueryFrequentFolder, 1, "", "Error occurred");

            // Act - 应该在多次重试后抛出 InvalidOperationException
            await QuickAccessQuery.GetFrequentFoldersAsync(2);
        }

        [TestMethod]
        public async Task GetAllItemsAsync_FailOnceSucceedLater_ReturnsCorrectItems()
        {
            // Arrange
            var mockItems = new List<string> { @"C:\Users\Test\Documents" };

            // 设置第一次失败，第二次成功
            _mockExecutor.SetupSequence(PSScript.QueryQuickAccess, new[]
            {
                new ScriptResult(1, "", "First attempt error"),
                new ScriptResult(0, string.Join(Environment.NewLine, mockItems), "")
            });

            // Act
            var result = await QuickAccessQuery.GetAllItemsAsync(2);

            // Assert
            CollectionAssert.AreEqual(mockItems, result, "应该返回第二次尝试的成功结果");
            Assert.AreEqual(2, _mockExecutor.GetCallCount(PSScript.QueryQuickAccess), "应该调用脚本两次");
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public async Task GetFrequentFoldersAsync_ExceedsMaxRetries_ThrowsException()
        {
            // Arrange - 设置所有尝试都失败
            _mockExecutor.SetupSequence(PSScript.QueryFrequentFolder, new[]
            {
                new ScriptResult(1, "", "First attempt error"),
                new ScriptResult(1, "", "Second attempt error"),
                new ScriptResult(1, "", "Third attempt error")
            });

            // Act - 应该在超过最大重试次数后抛出 InvalidOperationException
            await QuickAccessQuery.GetFrequentFoldersAsync(2);
        }

        #endregion

        #region 重试机制测试

        [TestMethod]
        public async Task GetRecentFilesAsync_AdminRetryWithFix_FixesExecutionPolicy()
        {
            // Arrange
            var mockFiles = new List<string> { @"C:\Users\Test\Documents\file.txt" };

            // 设置第一次失败，第二次成功
            _mockExecutor.SetupSequence(PSScript.QueryRecentFile, new[]
            {
                new ScriptResult(1, "", "First attempt error"),
                new ScriptResult(0, string.Join(Environment.NewLine, mockFiles), "")
            });

            // 设置为管理员
            _mockFeasible.SetIsAdministrator(true);

            // Act
            await QuickAccessQuery.GetRecentFilesAsync(2);

            // Assert
            Assert.IsTrue(_mockFeasible.WasFixExecutionPolicyCalled, "应该尝试修复执行策略");
        }

        #endregion

        #region 输出解析测试

        [TestMethod]
        public void ParsePowerShellOutput_EmptyString_ReturnsEmptyList()
        {
            // Arrange
            string output = "";

            // Act
            var result = InvokePrivateMethod<List<string>>("ParsePowerShellOutput", output);

            // Assert
            Assert.IsNotNull(result, "结果不应为null");
            Assert.AreEqual(0, result.Count, "空字符串应返回空列表");
        }

        [TestMethod]
        public void ParsePowerShellOutput_WhitespaceOnly_ReturnsEmptyList()
        {
            // Arrange
            string output = "   \r\n  \t  ";

            // Act
            var result = InvokePrivateMethod<List<string>>("ParsePowerShellOutput", output);

            // Assert
            Assert.IsNotNull(result, "结果不应为null");
            Assert.AreEqual(0, result.Count, "只有空白字符应返回空列表");
        }

        [TestMethod]
        public void ParsePowerShellOutput_MixedContent_ReturnsCleanedList()
        {
            // Arrange
            string output = "  Line1  \r\n\r\n  Line2\t\r\n   \r\nLine3";

            // Act
            var result = InvokePrivateMethod<List<string>>("ParsePowerShellOutput", output);

            // Assert
            Assert.AreEqual(3, result.Count, "应返回3个非空行");
            Assert.AreEqual("Line1", result[0], "应去除前后空白");
            Assert.AreEqual("Line2", result[1], "应去除前后空白");
            Assert.AreEqual("Line3", result[2], "应去除前后空白");
        }

        #endregion

        #region 辅助方法

        private T InvokePrivateMethod<T>(string methodName, params object[] parameters)
        {
            var method = typeof(QuickAccessQuery).GetMethod(methodName,
                BindingFlags.NonPublic | BindingFlags.Static);

            if (method == null)
                throw new ArgumentException($"Method {methodName} not found");

            return (T)method.Invoke(null, parameters);
        }

        #endregion
    }

    #region 测试替身类

    /// <summary>
    /// 用于测试的 ScriptExecutor 模拟服务
    /// </summary>
    internal class MockScriptExecutorService : QuickAccessQuery.IScriptExecutorService
    {
        private Dictionary<PSScript, ScriptResult> _results = new Dictionary<PSScript, ScriptResult>();
        private Dictionary<PSScript, Queue<ScriptResult>> _sequences = new Dictionary<PSScript, Queue<ScriptResult>>();
        private Dictionary<PSScript, int> _callCounts = new Dictionary<PSScript, int>();

        public void SetupResult(PSScript scriptType, int exitCode, string output, string error)
        {
            _results[scriptType] = new ScriptResult(exitCode, output, error);
        }

        public void SetupSequence(PSScript scriptType, ScriptResult[] results)
        {
            _sequences[scriptType] = new Queue<ScriptResult>(results);
        }

        public int GetCallCount(PSScript scriptType)
        {
            return _callCounts.TryGetValue(scriptType, out int count) ? count : 0;
        }

        public Task<ScriptResult> ExecutePSScript(PSScript scriptType, string parameter)
        {
            // 记录调用次数
            if (!_callCounts.ContainsKey(scriptType))
                _callCounts[scriptType] = 0;

            _callCounts[scriptType]++;

            // 检查是否有序列结果
            if (_sequences.TryGetValue(scriptType, out var sequence) && sequence.Count > 0)
            {
                return Task.FromResult(sequence.Dequeue());
            }

            // 检查是否有单一结果
            if (_results.TryGetValue(scriptType, out var result))
            {
                return Task.FromResult(result);
            }

            // 默认返回成功结果
            return Task.FromResult(new ScriptResult(0, "", ""));
        }
    }

    /// <summary>
    /// 用于测试的 ExecutionFeasible 模拟服务
    /// </summary>
    internal class MockExecutionFeasibleService : QuickAccessQuery.IExecutionFeasibleService
    {
        private bool _scriptFeasible = true;
        private bool _isAdmin = false;

        public bool WasFixExecutionPolicyCalled { get; private set; }

        public void SetScriptFeasible(bool feasible)
        {
            _scriptFeasible = feasible;
        }

        public void SetIsAdministrator(bool isAdmin)
        {
            _isAdmin = isAdmin;
        }

        public bool CheckScriptFeasible()
        {
            return _scriptFeasible;
        }

        public bool IsAdministrator()
        {
            return _isAdmin;
        }

        public void FixExecutionPolicy()
        {
            WasFixExecutionPolicyCalled = true;
        }
    }

    #endregion
}
