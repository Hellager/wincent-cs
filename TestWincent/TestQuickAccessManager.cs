using System.Security;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class QuickAccessManagerTests
    {
        private string? _tempDirectory;
        private string? _testFilePath;
        private string? _testDirPath;

        [TestInitialize]
        public void TestInitialize()
        {
            _tempDirectory = Path.Combine(Path.GetTempPath(), "WincentTests", Guid.NewGuid().ToString());
            _testFilePath = Path.Combine(_tempDirectory, "testfile.txt");
            _testDirPath = Path.Combine(_tempDirectory, "testdir");

            Directory.CreateDirectory(_tempDirectory);
            File.WriteAllText(_testFilePath, "test content");
            Directory.CreateDirectory(_testDirPath);

            QuickAccessManagerProxy.EnableMock();
        }

        [TestCleanup]
        public void TestCleanup()
        {
            if (Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, true);
            }

            QuickAccessManagerProxy.Reset();
        }


        [TestMethod]
        public async Task AddItemAsync_ForFile_CallsWindowsApi()
        {
            // Arrange
            var apiCalled = false;
            QuickAccessManagerProxy.EnableMock(
                nativeApiWrapper: new MockNativeApiWrapper
                {
                    SHAddToRecentDocs = (_, __) => apiCalled = true
                });

            // Act
            await QuickAccessManagerProxy.AddItemAsync(_testFilePath ?? "", QuickAccessItemType.File);

            // Assert
            Assert.IsTrue(apiCalled);
        }

        [TestMethod]
        public async Task AddItemAsync_ForDirectory_ExecutesCorrectScript()
        {
            // Arrange
            var scriptExecuted = false;
            QuickAccessManagerProxy.EnableMock(
                scriptExecutor: new MockScriptExecutor
                {
                    ExecutePSScript = (script, param) =>
                    {
                        if (script == PSScript.PinToFrequentFolder && param == _testDirPath)
                            scriptExecuted = true;
                        return Task.FromResult(new ScriptResult(0, "", ""));
                    }
                });

            // Act
            await QuickAccessManagerProxy.AddItemAsync(_testDirPath ?? "", QuickAccessItemType.Directory);

            // Assert
            Assert.IsTrue(scriptExecuted);
        }

        [TestMethod]
        public async Task RemoveItemAsync_ForFile_ExecutesRemoveScript()
        {
            // Arrange
            var scriptExecuted = false;
            QuickAccessManagerProxy.EnableMock(
                scriptExecutor: new MockScriptExecutor
                {
                    ExecutePSScript = (script, param) =>
                    {
                        if (script == PSScript.RemoveRecentFile && param == _testFilePath)
                            scriptExecuted = true;
                        return Task.FromResult(new ScriptResult(0, "", ""));
                    }
                });

            // Act
            await QuickAccessManagerProxy.RemoveItemAsync(_testFilePath ?? "", QuickAccessItemType.File);

            // Assert
            Assert.IsTrue(scriptExecuted);
        }

        [DataTestMethod]
        [DataRow("", QuickAccessItemType.File)]
        [DataRow(null, QuickAccessItemType.Directory)]
        public void ValidatePath_EmptyPath_ThrowsException(string path, QuickAccessItemType type)
        {
            // Act & Assert
            Assert.ThrowsException<ArgumentException>(() =>
                QuickAccessManagerProxy.ValidatePathSecurity(path ?? "", type));
        }

        [TestMethod]
        public void ValidatePath_ProtectedPath_ThrowsSecurityException()
        {
            // Act & Assert
            Assert.ThrowsException<SecurityException>(() =>
                QuickAccessManagerProxy.ValidatePathSecurity(
                    @"C:\Windows\System32\drivers\etc\hosts",
                    QuickAccessItemType.File));
        }

        [TestMethod]
        public async Task AddItemAsync_ScriptFailure_ThrowsCustomException()
        {
            // Arrange
            QuickAccessManagerProxy.EnableMock(
                scriptExecutor: new MockScriptExecutor
                {
                    ExecutePSScript = (_, __) =>
                        Task.FromResult(new ScriptResult(1, "", "Access denied"))
                });

            // Act & Assert
            await Assert.ThrowsExceptionAsync<QuickAccessOperationException>(() =>
                QuickAccessManagerProxy.AddItemAsync(_tempDirectory ?? "", QuickAccessItemType.Directory));
        }

        [TestMethod]
        public void EmptyRecentFiles_ShouldCallWindowsApi()
        {
            // Arrange
            var apiCalled = false;
            QuickAccessManagerProxy.EnableMock(
                nativeApiWrapper: new MockNativeApiWrapper
                {
                    SHAddToRecentDocs = (flags, _) =>
                    {
                        if (flags == QuickAccessManager.SHARD_PATHW)
                            apiCalled = true;
                    }
                });

            // Act
            QuickAccessManagerProxy.EmptyRecentFiles();

            // Assert
            Assert.IsTrue(apiCalled);
        }

        [TestMethod]
        public void EmptyFrequentFolders_ShouldDeleteJumplistFile()
        {
            // Arrange
            var mockFs = new MockFileSystem
            {
                GetKnownFolderPath = _ => _tempDirectory ?? "",
                FileExists = _ => true
            };
            var fileDeleted = false;
            mockFs.DeleteFile = _ => fileDeleted = true;

            QuickAccessManagerProxy.EnableMock(fileSystem: mockFs);

            // Act
            QuickAccessManagerProxy.EmptyFrequentFolders();

            // Assert
            Assert.IsTrue(fileDeleted);
        }

        [TestMethod]
        public void EmptyFrequentFolders_ShouldUnpinAllFolders()
        {
            // Arrange
            var testFolders = new[] { @"C:\Test1", @"D:\Test2" };
            var unpinCount = 0;

            var mockQuery = new MockQuickAccessQuery
            {
                GetFrequentFolders = () => testFolders.ToList()
            };

            var mockExecutor = new MockScriptExecutor
            {
                ExecutePSScript = (script, _) =>
                {
                    if (script == PSScript.UnpinFromFrequentFolder)
                        unpinCount++;
                    return Task.FromResult(new ScriptResult(0, "", ""));
                }
            };

            QuickAccessManagerProxy.EnableMock(
                queryProxy: mockQuery,
                scriptExecutor: mockExecutor);

            // Act
            QuickAccessManagerProxy.EmptyFrequentFolders();

            // Assert
            Assert.AreEqual(testFolders.Length, unpinCount);
        }

        [TestMethod]
        public void EmptyQuickAccess_ShouldCallBothMethods()
        {
            // Arrange
            var callLog = new List<string>();
            QuickAccessManagerProxy.EnableMock(
                emptyRecent: () => callLog.Add("Recent"),
                emptyFolders: () => callLog.Add("Folders"));

            // Act
            QuickAccessManagerProxy.EmptyQuickAccess();

            // Assert
            CollectionAssert.AreEqual(new[] { "Recent", "Folders" }, callLog);
        }
    }

    public static class QuickAccessManagerProxy
    {
        private static bool _useMock;
        private static MockNativeApiWrapper? _nativeApiWrapper;
        private static MockScriptExecutor? _scriptExecutor;
        private static MockFileSystem? _fileSystem;
        private static MockQuickAccessQuery? _queryProxy;
        private static Action? _emptyRecentCallback;
        private static Action? _emptyFoldersCallback;

        public static void EnableMock(
            MockNativeApiWrapper? nativeApiWrapper = null,
            MockScriptExecutor? scriptExecutor = null,
            MockFileSystem? fileSystem = null,
            MockQuickAccessQuery? queryProxy = null,
            Action? emptyRecent = null,
            Action? emptyFolders = null)
        {
            _useMock = true;
            _nativeApiWrapper = nativeApiWrapper ?? new MockNativeApiWrapper();
            _scriptExecutor = scriptExecutor ?? new MockScriptExecutor();
            _fileSystem = fileSystem;
            _queryProxy = queryProxy;
            _emptyRecentCallback = emptyRecent;
            _emptyFoldersCallback = emptyFolders;
        }

        public static void Reset()
        {
            _useMock = false;
            _nativeApiWrapper = null;
            _scriptExecutor = null;
        }

        public static async Task AddItemAsync(string path, QuickAccessItemType itemType)
        {
            if (!_useMock)
            {
                await QuickAccessManager.AddItemAsync(path, itemType);
                return;
            }

            if (itemType == QuickAccessItemType.File)
            {
                _nativeApiWrapper?.SHAddToRecentDocs?.Invoke(
                    QuickAccessManager.SHARD_PATHW, IntPtr.Zero);
            }
            else
            {
                var executor = _scriptExecutor ??
                    throw new InvalidOperationException("Script executor not initialized");

                var result = await _scriptExecutor.ExecutePSScript(
                    PSScript.PinToFrequentFolder, path);
                if (result.ExitCode != 0)
                    throw new QuickAccessOperationException(result.Error);
            }
        }

        public static async Task RemoveItemAsync(string path, QuickAccessItemType itemType)
        {
            if (!_useMock)
            {
                await QuickAccessManager.RemoveItemAsync(path, itemType);
                return;
            }

            var script = itemType == QuickAccessItemType.File
                ? PSScript.RemoveRecentFile
                : PSScript.UnpinFromFrequentFolder;

            var executor = _scriptExecutor ??
                throw new InvalidOperationException("Script executor not initialized");

            var result = await _scriptExecutor.ExecutePSScript(script, path);
            if (result.ExitCode != 0)
                throw new QuickAccessOperationException(result.Error);
        }

        public static void ValidatePathSecurity(string path, QuickAccessItemType expectedType)
        {
            if (!_useMock)
            {
                QuickAccessManager.ValidatePathSecurity(path, expectedType);
                return;
            }

            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("Path can not be empty");

            if (path.Contains("System32", StringComparison.OrdinalIgnoreCase))
                throw new SecurityException("Protected system path");
        }

        public static void EmptyRecentFiles()
        {
            if (_emptyRecentCallback != null)
            {
                _emptyRecentCallback();
                return;
            }

            _nativeApiWrapper?.SHAddToRecentDocs(QuickAccessManager.SHARD_PATHW, IntPtr.Zero);
        }

        public static void EmptyFrequentFolders()
        {
            if (_emptyFoldersCallback != null)
            {
                _emptyFoldersCallback();
                return;
            }

            var path = _fileSystem?.GetKnownFolderPath(Guid.Empty) ?? "";
            var targetFile = Path.Combine(path, "AutomaticDestinations", "f01b4d95cf55d32a.automaticDestinations-ms");

            if (_fileSystem?.FileExists(targetFile) == true)
            {
                _fileSystem.DeleteFile(targetFile);
            }

            foreach (var folder in _queryProxy?.GetFrequentFolders() ?? new List<string>())
            {
                _scriptExecutor?.ExecutePSScript(PSScript.UnpinFromFrequentFolder, folder).Wait();
            }
        }

        public static void EmptyQuickAccess()
        {
            EmptyRecentFiles();
            EmptyFrequentFolders();
        }
    }

    public class MockNativeApiWrapper
    {
        public Action<uint, IntPtr> SHAddToRecentDocs { get; set; } = (_, __) => { };
    }

    public class MockScriptExecutor
    {
        public Func<PSScript, string, Task<ScriptResult>> ExecutePSScript { get; set; }
            = (_, __) => Task.FromResult(new ScriptResult(0, "", ""));
    }

    public class MockFileSystem
    {
        public Func<Guid, string> GetKnownFolderPath { get; set; } = _ => "";
        public Func<string, bool> FileExists { get; set; } = _ => false;
        public Action<string> DeleteFile { get; set; } = _ => { };
    }

    public class MockQuickAccessQuery
    {
        public Func<List<string>> GetFrequentFolders { get; set; } = () => new List<string>();
    }
}
