using Microsoft.Win32;
using Moq;
using System.Diagnostics;
using System.Reflection;
using System.Security;
using Wincent;

namespace TestWincent
{
    [TestClass]
    [DoNotParallelize]
    public class TestFeasibleChecker
    {
        private const string TestRegistryPath = @"HKEY_CURRENT_USER\Software\WincentTest";
        private Mock<IRegistryOperations>? _mockRegistry;
        private Mock<IRegistryKeyProxy>? _mockKey;

        [TestInitialize]
        public void Initialize()
        {
            string subKeyPath = TestRegistryPath.Split(['\\'], 2)[1];

            var field = typeof(FeasibleChecker).GetField("_executionPolicyKeyPath",
                BindingFlags.NonPublic | BindingFlags.Static);
            if (field != null)
            {
                field.SetValue(null, subKeyPath);
            }

            _mockRegistry = new Mock<IRegistryOperations>();
            _mockKey = new Mock<IRegistryKeyProxy>();

            _mockRegistry.Setup(r => r.OpenCurrentUserSubKey(It.IsAny<string>(), It.IsAny<bool>()))
                         .Returns(_mockKey.Object);

            _mockRegistry.Setup(r => r.CreateCurrentUserSubKey(It.IsAny<string>()))
                         .Returns(_mockKey.Object);


            FeasibleChecker.InjectDependencies(_mockRegistry.Object);
        }

        [TestCleanup]
        public void Cleanup()
        {
            var field = typeof(FeasibleChecker).GetField("_executionPolicyKeyPath",
                BindingFlags.NonPublic | BindingFlags.Static);
            if (field != null)
            {
                field.SetValue(null, "Software\\WincentTest");
            }
            FeasibleChecker.ResetDependencies();

            try
            {
                string subKeyPath = TestRegistryPath.Split(['\\'], 2)[1];

                using (var testKey = Registry.CurrentUser.OpenSubKey(subKeyPath))
                {
                    if (testKey != null)
                    {
                        Registry.CurrentUser.DeleteSubKeyTree(subKeyPath, throwOnMissingSubKey: false);
                        Debug.WriteLine($"[Cleanup] Test registry path deleted: {TestRegistryPath}");
                    }
                }
            }
            catch (SecurityException ex)
            {
                Debug.WriteLine($"[Warning] Insufficient permissions to clean up registry: {ex.Message}");
            }
            catch (UnauthorizedAccessException ex)
            {
                Debug.WriteLine($"[Warning] Access denied: {ex.Message}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Error] Unexpected error during registry cleanup: {ex}");
            }
        }

        [TestMethod]
        public void OpenCurrentUserSubKey_ReturnsValidKey_WhenPathExists()
        {
            var registry = FeasibleChecker.GetCurrentRegistry();
            var key = registry.OpenCurrentUserSubKey(TestRegistryPath, false);
            Assert.IsNotNull(key);
        }

        [TestMethod]
        public void CheckDependencyInjection_ConfiguresMockCorrectly()
        {
            var registry = FeasibleChecker.GetCurrentRegistry();

            Assert.IsNotNull(registry, "Dependency injection failed");
            Assert.IsTrue(
                Mock.Get(registry).Behavior == MockBehavior.Default,
                "Default mock configuration not used"
            );
        }

        [TestMethod]
        public void CheckExecutionPolicyPath_ReflectionInjection_Success()
        {
            // Arrange
            const string expectedPath = "Software\\WincentTest";
            var privateField = typeof(FeasibleChecker)
                .GetField("_executionPolicyKeyPath", BindingFlags.NonPublic | BindingFlags.Static);

            Assert.IsNotNull(privateField, "_executionPolicyKeyPath field not found");

            // Act
            var actualValue = privateField!.GetValue(null);
            Assert.IsNotNull(actualValue, "Field value is null");
            var actualPath = (string)actualValue;

            // Assert
            Assert.AreEqual(expectedPath, actualPath, "Registry path injection failed");
        }

        [DataTestMethod]
        [DataRow("RemoteSigned", true)]
        [DataRow("Unrestricted", true)]
        [DataRow(null, false)]
        public void CheckScriptFeasible_PolicyVariants(string policyValue, bool expected)
        {
            // Arrange
            var mockKey = new Mock<IRegistryKeyProxy>();
            mockKey.Setup(k => k.GetValue(
                "ExecutionPolicy",
                It.Is<object>(v =>
                    v != null &&
                    v.GetType() == typeof(string) &&
                    (string)v == "NotSet"
                )
            )).Returns(policyValue);

            var mockRegistry = new Mock<IRegistryOperations>();
            mockRegistry.Setup(r => r.OpenCurrentUserSubKey(It.IsAny<string>(), false))
                        .Returns(mockKey.Object);

            FeasibleChecker.InjectDependencies(mockRegistry.Object);

            // Act
            var result = FeasibleChecker.CheckScriptFeasible();

            // Assert
            Assert.AreEqual(expected, FeasibleChecker.CheckScriptFeasible());
        }

        [TestMethod]
        public void CheckScriptFeasible_InvalidPolicyType_ReturnsFalse()
        {
            // Arrange
            var mockKey = new Mock<IRegistryKeyProxy>();
            mockKey.Setup(k => k.GetValue("ExecutionPolicy", It.IsAny<object>()))
                   .Returns(123);

            var mockRegistry = new Mock<IRegistryOperations>();

            mockRegistry.Setup(r => r.OpenCurrentUserSubKey(It.IsAny<string>(), false))
                       .Returns((IRegistryKeyProxy?)null);

            mockRegistry.Setup(r => r.CreateCurrentUserSubKey(It.IsAny<string>()))
                       .Returns(mockKey.Object);

            FeasibleChecker.InjectDependencies(mockRegistry.Object);

            // Act
            var result = FeasibleChecker.CheckScriptFeasible();

            // Assert
            Assert.IsFalse(result, "Non-string policy type not handled correctly");
        }

        [TestMethod]
        public void FixExecutionPolicy_RealWorldTest()
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree("Software\\WincentTest", false);
            }
            catch (Exception ex) when (
                ex is SecurityException ||
                ex is UnauthorizedAccessException
            )
            {
                Console.WriteLine($"Insufficient permissions: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Cleanup exception: {ex.Message}");
            }

            FeasibleChecker.ResetDependencies();
            FeasibleChecker.FixExecutionPolicy();

            using var key = Registry.CurrentUser.OpenSubKey("Software\\WincentTest");
            Assert.AreEqual("RemoteSigned", key?.GetValue("ExecutionPolicy"));
        }

        [TestMethod]
        public void FixExecutionPolicy_MockTest()
        {
            var mockKey = new Mock<IRegistryKeyProxy>();
            var mockRegistry = new Mock<IRegistryOperations>();

            mockRegistry.Setup(r => r.OpenCurrentUserSubKey(It.IsAny<string>(), true))
                       .Returns(mockKey.Object);

            mockRegistry.Setup(r => r.CreateCurrentUserSubKey(It.IsAny<string>()))
                       .Returns(mockKey.Object);

            FeasibleChecker.InjectDependencies(mockRegistry.Object);

            FeasibleChecker.FixExecutionPolicy();

            mockKey.Verify(
                k => k.SetValue(
                    "ExecutionPolicy",
                    "RemoteSigned",
                    It.Is<RegistryValueKind>(kind => kind == RegistryValueKind.String)
                ),
                Times.Once()
            );
        }

        [TestMethod]
        public void RegistryPathExists_ValidCustomPath_ReturnsTrue()
        {
            FeasibleChecker.ResetDependencies();
            const string subKey = "Software\\WincentTest\\ValidSubKey";

            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(subKey))
                {
                    key.SetValue("TestValue", 1);
                }

                bool exists = FeasibleChecker.RegistryPathExists(subKey);
                Assert.IsTrue(exists, "Valid path not correctly identified");
            }
            finally
            {
                Registry.CurrentUser.DeleteSubKeyTree(subKey, throwOnMissingSubKey: false);
            }
        }

        [TestMethod]
        public void RegistryPathExists_InvalidRootKey_ThrowsException()
        {
            Assert.ThrowsException<InvalidRegistryPathException>(() =>
            {
                FeasibleChecker.ResetDependencies();
                FeasibleChecker.RegistryPathExists(
                    "Software\\Test",
                    rootKey: "HKEY_INVALID_ROOT"
                );
            });
        }

        [TestMethod]
        public void RegistryPathExists_DefaultRootKey_Works()
        {
            FeasibleChecker.ResetDependencies();
            const string subKey = "Software\\WincentTest\\DefaultRootTest";

            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(subKey))
                {
                    key.SetValue("Test", 1);
                }

                bool exists = FeasibleChecker.RegistryPathExists(subKey);
                Assert.IsTrue(exists, "Default root key check failed");
            }
            finally
            {
                Registry.CurrentUser.DeleteSubKeyTree(subKey, throwOnMissingSubKey: false);
            }
        }
    }
}
