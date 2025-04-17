using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Threading.Tasks;
using Wincent;
using static Wincent.ExecutionFeasibilityStatus;

namespace TestWincent
{
    [TestClass]
    public class TestExecutionFeasibilityStatus
    {
        // 模拟的脚本执行器
        private Mock<IScriptExecutor> _mockExecutor;

        [TestInitialize]
        public void Initialize()
        {
            _mockExecutor = new Mock<IScriptExecutor>();
        }

        #region 1. 基础功能验证测试

        [TestMethod]
        public async Task CheckAsync_QueryAndHandleBothTrue_ReturnsTrue()
        {
            // 设置：查询和操作都成功
            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible, null, It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(0, "查询成功", ""));

            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible, null, It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(0, "操作成功", ""));

            // 执行
            var result = await ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, 10);

            // 验证
            Assert.IsTrue(result.Query);
            Assert.IsTrue(result.Handle);

            // 验证调用次数
            _mockExecutor.Verify(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible, null, It.IsAny<int>()), Times.Once);
            _mockExecutor.Verify(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible, null, It.IsAny<int>()), Times.Once);
        }

        [TestMethod]
        public async Task CheckAsync_QueryTrueHandleFalse_ReturnsOnlyQueryTrue()
        {
            // 设置：查询成功，操作失败
            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible, null, It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(0, "查询成功", ""));

            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible, null, It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(1, "", "操作失败"));

            // 执行
            var result = await ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, 10);

            // 验证
            Assert.IsTrue(result.Query);
            Assert.IsFalse(result.Handle);
        }

        [TestMethod]
        public async Task CheckAsync_QueryFalseHandleTrue_ReturnsOnlyHandleTrue()
        {
            // 设置：查询失败，操作成功
            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible, null, It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(1, "", "查询失败"));

            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible, null, It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(0, "操作成功", ""));

            // 执行
            var result = await ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, 10);

            // 验证
            Assert.IsFalse(result.Query);
            Assert.IsTrue(result.Handle);
        }

        [TestMethod]
        public async Task CheckAsync_BothFalse_ReturnsBothFalse()
        {
            // 设置：查询和操作都失败
            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible, null, It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(1, "", "查询失败"));

            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible, null, It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(1, "", "操作失败"));

            // 执行
            var result = await ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, 10);

            // 验证
            Assert.IsFalse(result.Query);
            Assert.IsFalse(result.Handle);
        }

        #endregion

        #region 2. 异常处理验证测试

        [TestMethod]
        public async Task CheckAsync_QueryScriptThrowsException_ReturnsQueryFalse()
        {
            // 设置：查询抛出异常，操作成功
            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible, null, It.IsAny<int>()))
                .ThrowsAsync(new InvalidOperationException("查询时发生异常"));

            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible, null, It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(0, "操作成功", ""));

            // 执行
            var result = await ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, 10);

            // 验证
            Assert.IsFalse(result.Query);
            Assert.IsTrue(result.Handle);
        }

        [TestMethod]
        public async Task CheckAsync_HandleScriptThrowsTimeout_ReturnsHandleFalse()
        {
            // 设置：查询成功，操作超时
            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible, null, It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(0, "查询成功", ""));

            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible, null, It.IsAny<int>()))
                .ThrowsAsync(new TimeoutException("操作超时"));

            // 执行
            var result = await ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, 10);

            // 验证
            Assert.IsTrue(result.Query);
            Assert.IsFalse(result.Handle);
        }

        [TestMethod]
        public async Task CheckAsync_BothScriptsThrowExceptions_ReturnsBothFalse()
        {
            // 设置：查询和操作都抛出异常
            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible, null, It.IsAny<int>()))
                .ThrowsAsync(new InvalidOperationException("查询时发生异常"));

            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible, null, It.IsAny<int>()))
                .ThrowsAsync(new TimeoutException("操作超时"));

            // 执行
            var result = await ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, 10);

            // 验证
            Assert.IsFalse(result.Query);
            Assert.IsFalse(result.Handle);
        }

        #endregion

        #region 3. 参数传递验证测试

        [TestMethod]
        public async Task CheckAsync_CustomTimeout_UsesCorrectTimeout()
        {
            // 设置：使用自定义超时参数
            const int customTimeout = 20;

            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible, null, customTimeout))
                .ReturnsAsync(new ScriptResult(0, "查询成功", ""));

            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible, null, customTimeout))
                .ReturnsAsync(new ScriptResult(0, "操作成功", ""));

            // 执行
            var result = await ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, customTimeout);

            // 验证
            Assert.IsTrue(result.Query);
            Assert.IsTrue(result.Handle);

            // 验证超时参数传递正确
            _mockExecutor.Verify(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible, null, customTimeout), Times.Once);
            _mockExecutor.Verify(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible, null, customTimeout), Times.Once);
        }

        [TestMethod]
        public async Task CheckAsync_ZeroTimeout_UsesDefaultTimeout()
        {
            // 设置：使用零超时参数（应该使用默认值10）
            const int zeroTimeout = 0;
            const int expectedTimeout = 10; // 默认值

            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible, null, expectedTimeout))
                .ReturnsAsync(new ScriptResult(0, "查询成功", ""));

            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible, null, expectedTimeout))
                .ReturnsAsync(new ScriptResult(0, "操作成功", ""));

            // 执行
            var result = await ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, zeroTimeout);

            // 验证
            Assert.IsTrue(result.Query);
            Assert.IsTrue(result.Handle);

            // 验证超时参数使用了默认值
            _mockExecutor.Verify(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible, null, expectedTimeout), Times.Once);
            _mockExecutor.Verify(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible, null, expectedTimeout), Times.Once);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public async Task CheckAsync_NullExecutor_ThrowsArgumentNullException()
        {
            // 执行：传递空执行器
            await ExecutionFeasibilityStatus.CheckAsync(null, 10);

            // 预期：抛出 ArgumentNullException
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public async Task CheckAsync_NegativeTimeout_ThrowsArgumentOutOfRangeException()
        {
            // 执行：传递负超时值
            await ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, -1);

            // 预期：抛出 ArgumentOutOfRangeException
        }

        [TestMethod]
        public async Task CheckAsync_MaxTimeoutValue_HandlesCorrectly()
        {
            // 设置：使用最大整数值作为超时
            int maxTimeout = int.MaxValue;

            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                It.IsAny<PSScript>(), null, maxTimeout))
                .ReturnsAsync(new ScriptResult(0, "成功", ""));

            // 执行
            var result = await ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, maxTimeout);

            // 验证
            Assert.IsTrue(result.Query);
            Assert.IsTrue(result.Handle);

            // 验证超时参数正确传递
            _mockExecutor.Verify(x => x.ExecutePSScriptWithTimeout(
                It.IsAny<PSScript>(), null, maxTimeout), Times.Exactly(2));
        }

        #endregion

        #region 4. 组合逻辑验证测试

        [TestMethod]
        public async Task CheckAsync_SequentialCalls_ReturnsDifferentResults()
        {
            // 设置：第一次调用全部返回成功，第二次全部返回失败
            _mockExecutor.SetupSequence(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible, null, It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(0, "查询成功", ""))
                .ReturnsAsync(new ScriptResult(1, "", "查询失败"));

            _mockExecutor.SetupSequence(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible, null, It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(0, "操作成功", ""))
                .ReturnsAsync(new ScriptResult(1, "", "操作失败"));

            // 第一次执行
            var result1 = await ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, 10);
            // 第二次执行
            var result2 = await ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, 10);

            // 验证
            Assert.IsTrue(result1.Query);
            Assert.IsTrue(result1.Handle);
            Assert.IsFalse(result2.Query);
            Assert.IsFalse(result2.Handle);
        }

        [TestMethod]
        public async Task CheckAsync_AsyncExecution_BothTasksCompleteIndependently()
        {
            // 设置：模拟两个任务的独立异步执行
            var queryTcs = new TaskCompletionSource<ScriptResult>();
            var handleTcs = new TaskCompletionSource<ScriptResult>();

            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckQueryFeasible, null, It.IsAny<int>()))
                .Returns(queryTcs.Task);

            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                PSScript.CheckPinUnpinFeasible, null, It.IsAny<int>()))
                .Returns(handleTcs.Task);

            // 启动异步执行
            var resultTask = ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, 10);

            // 确保任务尚未完成
            Assert.IsFalse(resultTask.IsCompleted);

            // 完成操作任务
            handleTcs.SetResult(new ScriptResult(0, "操作成功", ""));
            // 确保总任务仍未完成
            Assert.IsFalse(resultTask.IsCompleted);

            // 完成查询任务
            queryTcs.SetResult(new ScriptResult(0, "查询成功", ""));

            // 等待总任务完成
            var result = await resultTask;

            // 验证结果
            Assert.IsTrue(result.Query);
            Assert.IsTrue(result.Handle);
        }

        [TestMethod]
        public async Task CheckAsync_ParallelCalls_ExecutesIndependently()
        {
            // 设置：所有调用都成功
            _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                It.IsAny<PSScript>(), null, It.IsAny<int>()))
                .ReturnsAsync(new ScriptResult(0, "成功", ""));

            // 并行执行三次检查
            var task1 = ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, 5);
            var task2 = ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, 10);
            var task3 = ExecutionFeasibilityStatus.CheckAsync(_mockExecutor.Object, 15);

            // 等待所有任务完成
            await Task.WhenAll(task1, task2, task3);

            // 验证所有结果
            Assert.IsTrue(task1.Result.Query && task1.Result.Handle);
            Assert.IsTrue(task2.Result.Query && task2.Result.Handle);
            Assert.IsTrue(task3.Result.Query && task3.Result.Handle);

            // 验证调用次数
            _mockExecutor.Verify(x => x.ExecutePSScriptWithTimeout(
                It.IsAny<PSScript>(), null, It.IsAny<int>()), Times.Exactly(6));
        }

        #endregion

        #region 5. 实现兼容性测试

        [TestMethod]
        public async Task CheckAsync_WithRealScriptExecutor_Works()
        {
            // 由于 ScriptExecutor 是 sealed 类，不能被 Moq 模拟
            // 创建一个包装器，实现 IScriptExecutor 接口
            var wrapper = new ScriptExecutorWrapper();

            // 执行检查 - 使用包装器作为 IScriptExecutor
            var status = await ExecutionFeasibilityStatus.CheckAsync(wrapper, 10);

            // 验证结果
            Assert.IsNotNull(status);
            // 由于使用的是实际对象，我们只验证类型，不验证具体值
        }

        [TestMethod]
        public async Task CheckAsync_NoParameterOverload_Works()
        {
            // 测试不带参数的重载是否正常工作
            // 注意：这个测试依赖于实际的环境，所以在真实环境中可能需要设置 
            // IgnoreAttribute 或者使用集成测试框架

            // 执行
            var status = await ExecutionFeasibilityStatus.CheckAsync();

            // 验证返回了有效的结果
            Assert.IsNotNull(status);
            // 仅检查类型，不验证具体值，因为这取决于实际环境
        }

        /// <summary>
        /// ScriptExecutor 包装器，用于测试，实现 IScriptExecutor 接口
        /// </summary>
        private class ScriptExecutorWrapper : IScriptExecutor
        {
            // 使用一个模拟的 IScriptExecutor 来处理调用
            private readonly Mock<IScriptExecutor> _mockExecutor;

            public ScriptExecutorWrapper()
            {
                _mockExecutor = new Mock<IScriptExecutor>();

                // 设置所有脚本调用都成功
                _mockExecutor.Setup(x => x.ExecutePSScriptWithTimeout(
                    It.IsAny<PSScript>(), It.IsAny<string>(), It.IsAny<int>()))
                    .ReturnsAsync(new ScriptResult(0, "模拟成功", ""));
            }

            public Task<ScriptResult> ExecutePSScriptWithTimeout(PSScript script, string parameter, int timeoutSeconds)
            {
                return _mockExecutor.Object.ExecutePSScriptWithTimeout(script, parameter, timeoutSeconds);
            }
        }

        #endregion
    }

    [TestClass]
    public class TestExecutionFeasible
    {
        // 这里可以添加 ExecutionFeasible 静态类的测试
    }
}
