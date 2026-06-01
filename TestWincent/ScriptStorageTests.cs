using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class TestScriptStorage
    {
        private string _testFilePath;
        private string _testScriptVersion;

        [TestInitialize]
        public void Initialize()
        {
            ScriptStorage.CleanupDynamicScripts();

            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var version = assembly.GetName().Version;
                _testScriptVersion = $"v{version.Major}.{version.Minor}.{version.Build}";
            }
            catch
            {
                _testScriptVersion = "v0.1.0";
            }

            _testFilePath = Path.Combine(
                ScriptStorage.DynamicScriptDir,
                $"TestScript_{_testScriptVersion}_{Guid.NewGuid():N}.ps1");

            File.WriteAllText(_testFilePath, "Test content");

            Assert.IsTrue(File.Exists(_testFilePath), "Test file should have been created");
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (File.Exists(_testFilePath))
            {
                try
                {
                    File.Delete(_testFilePath);
                }
                catch
                {
                    // Ignore deletion errors
                }
            }
        }

        #region Directory Structure Tests

        [TestMethod]
        public void ScriptRoot_ExistsAndIsInTempFolder()
        {
            // Act - trigger lazy initialization
            ScriptStorage.GetScriptPath(PSScript.RefreshExplorer);

            // Assert
            Assert.IsTrue(Directory.Exists(ScriptStorage.ScriptRoot), "Script root directory should exist");
            Assert.IsTrue(ScriptStorage.ScriptRoot.StartsWith(Path.GetTempPath()), "Script root should be in temp folder");
            Assert.IsTrue(ScriptStorage.ScriptRoot.EndsWith("Wincent"), "Script root should be named Wincent");
        }

        [TestMethod]
        public void StaticScriptDir_ExistsAndIsInRootFolder()
        {
            // Act - trigger lazy initialization
            ScriptStorage.GetScriptPath(PSScript.RefreshExplorer);

            // Assert
            Assert.IsTrue(Directory.Exists(ScriptStorage.StaticScriptDir), "Static script directory should exist");
            Assert.IsTrue(ScriptStorage.StaticScriptDir.StartsWith(ScriptStorage.ScriptRoot), "Static script directory should be in root");
            Assert.IsTrue(ScriptStorage.StaticScriptDir.EndsWith("static"), "Static script directory should be named static");
        }

        [TestMethod]
        public void DynamicScriptDir_ExistsAndIsInRootFolder()
        {
            // Act - trigger lazy initialization
            ScriptStorage.CleanupDynamicScripts();

            // Assert
            Assert.IsTrue(Directory.Exists(ScriptStorage.DynamicScriptDir), "Dynamic script directory should exist");
            Assert.IsTrue(ScriptStorage.DynamicScriptDir.StartsWith(ScriptStorage.ScriptRoot), "Dynamic script directory should be in root");
            Assert.IsTrue(ScriptStorage.DynamicScriptDir.EndsWith("dynamic"), "Dynamic script directory should be named dynamic");
        }

        #endregion

        #region Script Path Tests

        [TestMethod]
        public void GetScriptPath_ForNonParameterizedScript_ReturnsCorrectPath()
        {
            var script = PSScript.RefreshExplorer;

            string path = ScriptStorage.GetScriptPath(script);

            Assert.IsTrue(path.StartsWith(ScriptStorage.StaticScriptDir), "Non-parameterized script should be in static directory");
            Assert.IsTrue(path.Contains($"{script}_{_testScriptVersion}"), "Script filename should contain version");
            Assert.IsTrue(path.EndsWith("ps1"), "Script filename should have ps1 extension");
            Assert.IsTrue(File.Exists(path), "Script file should have been created");

            string content = File.ReadAllText(path);
            Assert.IsTrue(content.Contains("Shell.Application"), "Script content should be correct");

            byte[] bytes = File.ReadAllBytes(path);
            Assert.IsTrue(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF,
                "Script file should use UTF8-BOM encoding");

            var fileName = Path.GetFileName(path);
            Assert.IsTrue(Regex.IsMatch(fileName, @"^RefreshExplorer_v\d+\.\d+\.\d+\.ps1$"),
                "Filename format should be {ScriptName}_{Version}.ps1");
        }

        [TestMethod]
        public void GetScriptPath_ForParameterizedScript_ReturnsCorrectPath()
        {
            var script = PSScript.RemoveRecentFile;

            string path = ScriptStorage.GetScriptPath(script);

            Assert.IsTrue(path.StartsWith(ScriptStorage.DynamicScriptDir), "Parameterized script should be in dynamic directory");
            Assert.IsTrue(path.Contains($"{script}_{_testScriptVersion}"), "Script filename should contain version");
            Assert.IsTrue(path.EndsWith("ps1"), "Script filename should have ps1 extension");
            Assert.IsFalse(File.Exists(path), "Parameterized script should not be auto-created");
        }

        [TestMethod]
        public void GetDynamicScriptPath_WithValidParameter_ReturnsAndCreatesScript()
        {
            var script = PSScript.RemoveRecentFile;
            var parameter = @"C:\Test\File.txt";

            string path = ScriptStorage.GetDynamicScriptPath(script, parameter);

            Assert.IsTrue(path.StartsWith(ScriptStorage.DynamicScriptDir), "Dynamic script should be in dynamic directory");
            Assert.IsTrue(path.Contains(script.ToString()), "Script filename should contain script type");
            Assert.IsTrue(path.Contains(_testScriptVersion), "Script filename should contain version");
            Assert.IsTrue(File.Exists(path), "Dynamic script file should have been created");

            string content = File.ReadAllText(path);
            Assert.IsTrue(content.Contains(parameter), "Script content should contain parameter");

            byte[] bytes = File.ReadAllBytes(path);
            Assert.IsTrue(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF,
                "Script file should use UTF8-BOM encoding");

            var fileName = Path.GetFileName(path);
            Assert.IsTrue(Regex.IsMatch(fileName, @"^RemoveRecentFile_v\d+\.\d+\.\d+_[0-9A-F]{8}\.ps1$"),
                "Filename format should be {ScriptName}_{Version}_{ParamHash}.ps1");
        }

        [TestMethod]
        public void GetDynamicScriptPath_WithQuotesInParameter_HandlesCorrectly()
        {
            var script = PSScript.RemoveRecentFile;
            var parameter = @"C:\Test\File's.txt";

            string path = ScriptStorage.GetDynamicScriptPath(script, parameter);

            Assert.IsTrue(File.Exists(path), "Script file with single quotes should have been created");

            string content = File.ReadAllText(path);
            Assert.IsFalse(content.Contains("File's.txt"), "Single quotes should be escaped");
            Assert.IsTrue(content.Contains("File''s.txt"), "Single quote should be replaced with two single quotes");

            File.Delete(path);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void GetDynamicScriptPath_WithNonParameterizedScript_ThrowsException()
        {
            var script = PSScript.RefreshExplorer;
            var parameter = @"C:\Test\File.txt";

            ScriptStorage.GetDynamicScriptPath(script, parameter);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void GetDynamicScriptPath_WithEmptyParameter_ThrowsException()
        {
            var script = PSScript.RemoveRecentFile;
            var parameter = "";

            ScriptStorage.GetDynamicScriptPath(script, parameter);
        }

        #endregion

        #region UTF8 BOM Tests

        [TestMethod]
        public void AddUtf8Bom_WithNoBom_AddsBom()
        {
            var content = Encoding.UTF8.GetBytes("test content");

            var result = ScriptStorage.AddUtf8Bom(content);

            Assert.AreEqual(content.Length + 3, result.Length, "Should add 3-byte BOM");
            Assert.AreEqual(0xEF, result[0], "First byte should be 0xEF");
            Assert.AreEqual(0xBB, result[1], "Second byte should be 0xBB");
            Assert.AreEqual(0xBF, result[2], "Third byte should be 0xBF");
        }

        [TestMethod]
        public void AddUtf8Bom_WithExistingBom_DoesNotDuplicate()
        {
            var bom = Encoding.UTF8.GetPreamble();
            var text = Encoding.UTF8.GetBytes("test content");
            var contentWithBom = new byte[bom.Length + text.Length];
            Buffer.BlockCopy(bom, 0, contentWithBom, 0, bom.Length);
            Buffer.BlockCopy(text, 0, contentWithBom, bom.Length, text.Length);

            var result = ScriptStorage.AddUtf8Bom(contentWithBom);

            Assert.AreEqual(contentWithBom.Length, result.Length, "Should not add duplicate BOM");
            for (int i = 0; i < contentWithBom.Length; i++)
            {
                Assert.AreEqual(contentWithBom[i], result[i], $"Byte {i} should be identical");
            }
        }

        #endregion

        #region Cleanup Tests

        [TestMethod]
        public void CleanupDynamicScripts_RemovesOldScripts()
        {
            var testFiles = new string[3];
            for (int i = 0; i < testFiles.Length; i++)
            {
                testFiles[i] = Path.Combine(
                    ScriptStorage.DynamicScriptDir,
                    $"TestCleanup_{_testScriptVersion}_{i}.ps1");

                File.WriteAllText(testFiles[i], $"Test content {i}");

                File.SetLastWriteTime(testFiles[i], DateTime.Now.AddHours(-i * 10 - 1));
            }

            ScriptStorage.CleanupDynamicScripts(5);

            Assert.IsTrue(File.Exists(testFiles[0]), "Recent file should not be deleted");
            Assert.IsFalse(File.Exists(testFiles[1]), "Old file should be deleted");
            Assert.IsFalse(File.Exists(testFiles[2]), "Old file should be deleted");

            if (File.Exists(testFiles[0]))
                File.Delete(testFiles[0]);
        }

        [TestMethod]
        public void CleanupByVersion_RemovesOldVersionScripts()
        {
            var oldVersion = "v0.0.1";
            var testFile = Path.Combine(
                ScriptStorage.StaticScriptDir,
                $"RefreshExplorer_{oldVersion}.ps1");

            File.WriteAllText(testFile, "Test content for old version");

            Assert.IsTrue(File.Exists(testFile), "Test file should have been created");

            ScriptStorage.CleanupStaticScripts();

            Assert.IsFalse(File.Exists(testFile), "Old version script should be deleted");

            var newScript = ScriptStorage.GetScriptPath(PSScript.RefreshExplorer);
            Assert.IsTrue(File.Exists(newScript), "New version script should be created");
            Assert.IsTrue(newScript.Contains(_testScriptVersion), "New script should use current version");
        }

        #endregion

        #region Hash Generation Tests

        [TestMethod]
        public void GetParameterHash_DifferentParameters_GenerateDifferentHashes()
        {
            var script = PSScript.RemoveRecentFile;
            var param1 = @"C:\Test\File1.txt";
            var param2 = @"C:\Test\File2.txt";

            string path1 = ScriptStorage.GetDynamicScriptPath(script, param1);
            string path2 = ScriptStorage.GetDynamicScriptPath(script, param2);

            string fileName1 = Path.GetFileName(path1);
            string fileName2 = Path.GetFileName(path2);

            var match1 = Regex.Match(fileName1, @"^[A-Za-z]+_v\d+\.\d+\.\d+_([0-9A-F]{8})\.ps1$");
            var match2 = Regex.Match(fileName2, @"^[A-Za-z]+_v\d+\.\d+\.\d+_([0-9A-F]{8})\.ps1$");

            Assert.IsTrue(match1.Success, "Filename 1 should match expected format");
            Assert.IsTrue(match2.Success, "Filename 2 should match expected format");

            string hash1 = match1.Groups[1].Value;
            string hash2 = match2.Groups[1].Value;

            Assert.AreNotEqual(hash1, hash2, "Different parameters should generate different hashes");
            Assert.AreEqual(8, hash1.Length, "Hash length should be 8 characters");
            Assert.AreEqual(8, hash2.Length, "Hash length should be 8 characters");

            File.Delete(path1);
            File.Delete(path2);
        }

        [TestMethod]
        public void GetParameterHash_SameParameter_GeneratesSameHash()
        {
            var script = PSScript.RemoveRecentFile;
            var param = @"C:\Test\File.txt";

            string path1 = ScriptStorage.GetDynamicScriptPath(script, param);
            string path2 = ScriptStorage.GetDynamicScriptPath(script, param);

            Assert.AreEqual(path1, path2, "Same parameter should generate same path");
            Assert.IsTrue(File.Exists(path1), "Script file should have been created");

            File.Delete(path1);
        }

        [TestMethod]
        public async Task GetDynamicScriptPath_ConcurrentSameParameter_CreatesOneCompleteFile()
        {
            var script = PSScript.RemoveRecentFile;
            var param = $@"C:\Test\Concurrent_{Guid.NewGuid():N}.txt";

            var tasks = Enumerable.Range(0, 20)
                .Select(_ => Task.Run(() => ScriptStorage.GetDynamicScriptPath(script, param)))
                .ToArray();

            string[] paths = await Task.WhenAll(tasks);
            string expectedPath = paths[0];

            Assert.AreEqual(1, paths.Distinct(StringComparer.OrdinalIgnoreCase).Count(), "Concurrent calls should return same path");
            Assert.IsTrue(File.Exists(expectedPath), "Script file should have been created");

            string fileName = Path.GetFileName(expectedPath);
            string[] matchingFiles = Directory.GetFiles(ScriptStorage.DynamicScriptDir, fileName);
            Assert.AreEqual(1, matchingFiles.Length, "Only one script file should be created for same parameter");

            string content = File.ReadAllText(expectedPath);
            Assert.IsTrue(content.Contains(param), "Script content should contain complete parameter");
            Assert.IsTrue(content.Contains("Shell.Application"), "Script content should be fully generated");

            DateTime firstWriteTime = File.GetLastWriteTimeUtc(expectedPath);
            Thread.Sleep(100);

            string repeatedPath = ScriptStorage.GetDynamicScriptPath(script, param);
            DateTime secondWriteTime = File.GetLastWriteTimeUtc(repeatedPath);

            Assert.AreEqual(expectedPath, repeatedPath, "Repeated call should return same path");
            Assert.AreEqual(firstWriteTime, secondWriteTime, "Existing script should not be overwritten");

            File.Delete(expectedPath);
        }

        #endregion

        #region Integration Tests

        [TestMethod]
        public void ScriptStorage_IntegrationTest_CreateAndExecuteScript()
        {
            var script = PSScript.QueryQuickAccess;

            string scriptPath = ScriptStorage.GetScriptPath(script);

            Assert.IsTrue(File.Exists(scriptPath), "Script file should have been created");

            string content = File.ReadAllText(scriptPath);
            Assert.IsTrue(content.Contains("Shell.Application"), "Script content should be correct");
            Assert.IsTrue(content.Contains("Namespace"), "Script content should be correct");

            byte[] bytes = File.ReadAllBytes(scriptPath);
            Assert.IsTrue(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF,
                "Script file should use UTF8-BOM encoding");
        }

        [TestMethod]
        public void ScriptStorage_IsParameterizedScript_CorrectlyIdentifies()
        {
            Assert.IsTrue(ScriptStorage.IsParameterizedScript(PSScript.RemoveRecentFile),
                "RemoveRecentFile should be identified as parameterized");
            Assert.IsTrue(ScriptStorage.IsParameterizedScript(PSScript.PinToFrequentFolder),
                "PinToFrequentFolder should be identified as parameterized");
            Assert.IsTrue(ScriptStorage.IsParameterizedScript(PSScript.UnpinFromFrequentFolder),
                "UnpinFromFrequentFolder should be identified as parameterized");
            Assert.IsTrue(ScriptStorage.IsParameterizedScript(PSScript.AddRecentFile),
                "AddRecentFile should be identified as parameterized");

            Assert.IsFalse(ScriptStorage.IsParameterizedScript(PSScript.RefreshExplorer),
                "RefreshExplorer should not be identified as parameterized");
            Assert.IsFalse(ScriptStorage.IsParameterizedScript(PSScript.QueryQuickAccess),
                "QueryQuickAccess should not be identified as parameterized");
            Assert.IsFalse(ScriptStorage.IsParameterizedScript(PSScript.QueryRecentFile),
                "QueryRecentFile should not be identified as parameterized");
            Assert.IsFalse(ScriptStorage.IsParameterizedScript(PSScript.QueryFrequentFolder),
                "QueryFrequentFolder should not be identified as parameterized");
        }

        #endregion
    }
}
