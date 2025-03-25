using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class ScriptExecutorTests
    {
        private ScriptExecutor _executor;
        [TestInitialize]
        public void Initialize()
        {
            // 使用自定义 Mock 工厂
            _executor = new ScriptExecutor(new MockScriptStrategyFactory());
        }
        [TestMethod]
        public async Task ExecutePSScript_ReturnsValidResult()
        {
            // 测试基础脚本执行
            var result = await _executor.ExecutePowerShellScriptAsync("Write-Output 'test'");
            Assert.AreEqual(0, result.ExitCode);
            StringAssert.Contains(result.Output, "test");
        }
        [TestMethod]
        [ExpectedException(typeof(ScriptExecutionException))]
        public async Task ExecutePSScript_HandlesErrors()
        {
            // 测试错误脚本执行
            await _executor.ExecutePowerShellScriptAsync("Throw 'test error'");
        }
        #region Mock 实现
        // 实现缺失的 Mock 工厂
        public class MockScriptStrategyFactory : IPSScriptStrategyFactory
        {
            private readonly Dictionary<PSScript, IPSScriptStrategy> _strategyCache
                = new Dictionary<PSScript, IPSScriptStrategy>();
            public IPSScriptStrategy GetStrategy(PSScript method)
            {
                if (_strategyCache.TryGetValue(method, out var cachedStrategy))
                {
                    return cachedStrategy;
                }
                IPSScriptStrategy strategy;
                switch (method)
                {
                    case PSScript.RefreshExplorer:
                        strategy = new MockRefreshExplorerStrategy();
                        break;
                    case PSScript.QueryRecentFile:
                        strategy = new MockQueryRecentFileStrategy();
                        break;
                    case PSScript.CheckQueryFeasible:
                        strategy = new MockCheckFeasibleStrategy();
                        break;
                    default:
                        throw new NotSupportedException($"Unsupported script type: {method}");
                }
                _strategyCache[method] = strategy;
                return strategy;
            }
        }

        // 示例 Mock 策略实现
        public class MockRefreshExplorerStrategy : IPSScriptStrategy
        {
            public string GenerateScript(string parameter)
            {
                return "Write-Output 'MockRefreshExplorer'";
            }
        }
        public class MockQueryRecentFileStrategy : IPSScriptStrategy
        {
            public string GenerateScript(string parameter)
            {
                return "Write-Output 'MockQueryRecentFile'";
            }
        }

        public class MockCheckFeasibleStrategy : IPSScriptStrategy
        {
            public string GenerateScript(string parameter)
            {
                return "Write-Output 'MockCheckFeasibleStrategy'";
            }
        }
        #endregion
        [TestCleanup]
        public void Cleanup()
        {
            _executor?.Dispose();

            // 清理生成的脚本文件
            foreach (var file in Directory.GetFiles(ScriptStorage.ScriptRoot))
            {
                File.Delete(file);
            }
        }
    }
}
