using System.Text;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class TestScriptExecutor
    {
        [TestMethod]
        public async Task TestBasicScriptExecution()
        {
            var executor = new ScriptExecutor();
            string script = @"
                Write-Output 'Hello from PowerShell'
                exit 0
            ";

            var result = await executor.ExecutePowerShellScriptAsync(script);

            Assert.AreEqual(0, result.ExitCode, "Exit code should be 0");
            StringAssert.Contains(result.Output, "Hello from PowerShell", "Output should contain expected message");
            Assert.AreEqual("", result.Error.Trim(), "Error output should be empty");
        }

        [TestMethod]
        public async Task TestScriptWithError()
        {
            var executor = new ScriptExecutor();
            string script = @"
                Write-Output 'Normal output'
                Write-Error 'Error message'
                exit 1
            ";

            var result = await executor.ExecutePowerShellScriptAsync(script);

            Assert.AreEqual(1, result.ExitCode, "Exit code should be 1");
            StringAssert.Contains(result.Output, "Normal output", "Output should contain normal message");
            StringAssert.Contains(result.Error, "Error message", "Error output should contain error message");
        }

        [TestMethod]
        public async Task TestScriptWithUTF8Encoding()
        {
            var executor = new ScriptExecutor();
            string script = @"
                $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
                Write-Output '中文输出测试'
                Write-Error '中文错误测试'
            ";

            var result = await executor.ExecutePowerShellScriptAsync(script);

            StringAssert.Contains(result.Output, "中文输出测试", "Output should contain UTF-8 characters");
            StringAssert.Contains(result.Error, "中文错误测试", "Error should contain UTF-8 characters");
        }

        [TestMethod]
        public async Task TestScriptTimeout()
        {
            var executor = new ScriptExecutor();
            string script = @"
                Write-Output 'Starting long operation'
                Start-Sleep -Seconds 10
                Write-Output 'This should not be seen'
            ";

            try
            {
                await executor.ExecutePowerShellScriptAsync(script, TimeSpan.FromSeconds(1));
                Assert.Fail("Should have thrown a timeout exception");
            }
            catch (ScriptTimeoutException ex)
            {
                StringAssert.Contains(ex.Output, "Starting long operation", "Output should contain initial message");
                StringAssert.Contains(ex.Message, "Script execution timed out", "Exception message should mention timeout");
            }
        }

        [TestMethod]
        public async Task TestBinaryScriptExecution()
        {
            var executor = new ScriptExecutor();
            string script = "Write-Output 'Binary script test'";
            byte[] scriptBytes = Encoding.UTF8.GetBytes(script);

            var result = await executor.ExecutePowerShellScriptAsync(scriptBytes);

            Assert.AreEqual(0, result.ExitCode, "Exit code should be 0");
            StringAssert.Contains(result.Output, "Binary script test", "Output should contain expected message");
        }

        [TestMethod]
        public async Task TestPredefinedScripts()
        {
            var executor = new ScriptExecutor();

            var result = await executor.ExecutePSScript(PSScript.QueryRecentFile, null);

            Assert.IsNotNull(result, "Result should not be null");
        }

        [TestMethod]
        public async Task TestInvalidScript()
        {
            var executor = new ScriptExecutor();
            string script = @"
                This-Is-Not-A-Valid-Command
                exit
            ";

            var result = await executor.ExecutePowerShellScriptAsync(script);

            Assert.AreNotEqual(0, result.ExitCode, "Exit code should not be 0 for invalid script");
            Assert.IsTrue(!string.IsNullOrEmpty(result.Error), "Error output should not be empty");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public async Task TestNullScript()
        {
            var executor = new ScriptExecutor();
            string script = null;

            await executor.ExecutePowerShellScriptAsync(script);
        }

        [TestMethod]
        public async Task TestCustomStrategyFactory()
        {
            var mockFactory = new MockPSScriptStrategyFactory();
            var executor = new ScriptExecutor(mockFactory);

            var result = await executor.ExecutePSScript(PSScript.RefreshExplorer, null);

            Assert.IsNotNull(result, "Result should not be null");
        }
    }

    public class MockPSScriptStrategyFactory : IPSScriptStrategyFactory
    {
        public IPSScriptStrategy GetStrategy(PSScript method)
        {
            return new MockPSScriptStrategy();
        }
    }

    public class MockPSScriptStrategy : IPSScriptStrategy
    {
        public string GenerateScript(string? parameter)
        {
            return "Write-Output 'Mock script executed'";
        }
    }
}
