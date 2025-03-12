using Moq;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class TestQuickAccessQuery
    {
        [TestInitialize]
        public void Setup()
        {
        }

        [TestCleanup]
        public void Cleanup()
        {
            QuickAccessQueryProxy.Reset();
        }

        [TestMethod]
        public async Task GetRecentFilesAsync_WhenAllChecksPass_ReturnsFiles()
        {
            // Arrange
            QuickAccessQueryProxy.EnableMock(
                checkScriptFeasible: () => true,
                executeScript: (_, __) => Task.FromResult(
                    new ScriptResult(0, "C:\\file1.txt\nC:\\file2.doc", "")
                )
            );

            // Act
            var result = await QuickAccessQueryProxy.GetRecentFilesAsync();

            // Assert
            Assert.AreEqual(2, result.Count);
            Assert.IsTrue(result.Contains("C:\\file1.txt"));
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public async Task GetRecentFilesAsync_WhenScriptNotFeasible_Throws()
        {
            // Arrange
            QuickAccessQueryProxy.EnableMock(checkScriptFeasible: () => false);

            // Act & Assert
            await QuickAccessQueryProxy.GetRecentFilesAsync();
        }
    }

    public static class QuickAccessQueryProxy
    {
        private static bool _useMock;

        private static Func<bool> MockCheckScriptFeasible { get; set; }
            = FeasibleChecker.CheckScriptFeasible;

        public static Func<PSScript, string?, Task<ScriptResult>> MockExecuteScript { get; set; } =
            (script, param) => ScriptExecutor.ExecutePSScript(script, param ?? string.Empty);

        private static readonly char[] separator = ['\r', '\n'];

        public static async Task<List<string>> GetRecentFilesAsync()
        {
            if (!_useMock)
                return await QuickAccessQuery.GetRecentFilesAsync();

            if (!MockCheckScriptFeasible())
                throw new InvalidOperationException("PowerShell script execution is not available");

            var result = await MockExecuteScript(PSScript.QueryRecentFile, string.Empty);

            if (result.ExitCode != 0)
                throw new InvalidOperationException(result.Error);

            return ProcessResult(result);
        }

        private static List<string> ProcessResult(ScriptResult result)
        {
            return result.Output.Split(separator, StringSplitOptions.RemoveEmptyEntries)
                              .Select(s => s.Trim())
                              .Where(s => !string.IsNullOrEmpty(s))
                              .ToList();
        }

        public static void EnableMock(
            Func<bool>? checkScriptFeasible = null,
            Func<PSScript, string?, Task<ScriptResult>>? executeScript = null)
        {
            _useMock = true;

            MockCheckScriptFeasible = checkScriptFeasible ?? MockCheckScriptFeasible!;
            MockExecuteScript = executeScript ?? MockExecuteScript!;
        }

        public static void Reset()
        {
            _useMock = false;
            MockCheckScriptFeasible = FeasibleChecker.CheckScriptFeasible;
            MockExecuteScript = ScriptExecutor.ExecutePSScript;
        }
    }
}
