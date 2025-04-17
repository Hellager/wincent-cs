//using Microsoft.VisualStudio.TestTools.UnitTesting;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Security;
//using System.Text;
//using System.Threading.Tasks;
//using Wincent;

//namespace TestWincent
//{
//    [TestClass]
//    public class TestExecutionFeasible
//    {
//        private MockRegistryService _mockRegistry;

//        [TestInitialize]
//        public void Initialize()
//        {
//            // 创建模拟服务
//            _mockRegistry = new MockRegistryService();

//            // 替换服务实现
//            ExecutionFeasible.SetRegistryService(_mockRegistry);

//            // 设置默认行为
//            _mockRegistry.SetExecutionPolicyValue("Restricted");
//            _mockRegistry.SetIsAdministrator(false);
//        }

//        [TestCleanup]
//        public void Cleanup()
//        {
//            // 恢复默认服务
//            ExecutionFeasible.ResetRegistryService();
//        }

//        #region 基本功能测试

//        [TestMethod]
//        public void CheckScriptFeasible_WithValidPolicy_ReturnsTrue()
//        {
//            // Arrange
//            string[] validPolicies = { "AllSigned", "Bypass", "RemoteSigned", "Unrestricted" };

//            foreach (var policy in validPolicies)
//            {
//                // 设置有效的执行策略
//                _mockRegistry.SetExecutionPolicyValue(policy);

//                // Act
//                bool result = ExecutionFeasible.CheckScriptFeasible();

//                // Assert
//                Assert.IsTrue(result, $"执行策略 {policy} 应该被视为可行");
//            }
//        }

//        [TestMethod]
//        public void CheckScriptFeasible_WithInvalidPolicy_ReturnsFalse()
//        {
//            // Arrange
//            string[] invalidPolicies = { "Restricted", "Default", "Undefined", "NotSet", "" };

//            foreach (var policy in invalidPolicies)
//            {
//                // 设置无效的执行策略
//                _mockRegistry.SetExecutionPolicyValue(policy);

//                // Act
//                bool result = ExecutionFeasible.CheckScriptFeasible();

//                // Assert
//                Assert.IsFalse(result, $"执行策略 {policy} 应该被视为不可行");
//            }
//        }

//        [TestMethod]
//        public void GetExecutionPolicy_ReturnsCorrectValue()
//        {
//            // Arrange
//            string expectedPolicy = "RemoteSigned";
//            _mockRegistry.SetExecutionPolicyValue(expectedPolicy);

//            // Act
//            string result = ExecutionFeasible.GetExecutionPolicy();

//            // Assert
//            Assert.AreEqual(expectedPolicy, result, "应该返回正确的执行策略值");
//        }

//        [TestMethod]
//        public void IsAdministrator_ReturnsCorrectValue()
//        {
//            // Arrange - 设置为非管理员
//            _mockRegistry.SetIsAdministrator(false);

//            // Act
//            bool result = ExecutionFeasible.IsAdministrator();

//            // Assert
//            Assert.IsFalse(result, "非管理员应该返回 false");

//            // Arrange - 设置为管理员
//            _mockRegistry.SetIsAdministrator(true);

//            // Act
//            result = ExecutionFeasible.IsAdministrator();

//            // Assert
//            Assert.IsTrue(result, "管理员应该返回 true");
//        }

//        #endregion

//        #region 异常处理测试

//        [TestMethod]
//        public void CheckScriptFeasible_WithSecurityException_ReturnsFalse()
//        {
//            // Arrange
//            _mockRegistry.ThrowOnGetPolicy = true;

//            // Act
//            bool result = ExecutionFeasible.CheckScriptFeasible();

//            // Assert
//            Assert.IsFalse(result, "发生安全异常时应该返回 false");
//        }

//        [TestMethod]
//        [ExpectedException(typeof(SecurityException))]
//        public void FixExecutionPolicy_WithSecurityException_ThrowsException()
//        {
//            // Arrange
//            _mockRegistry.ThrowOnSetPolicy = true;

//            // Act - 应该抛出 SecurityException
//            ExecutionFeasible.FixExecutionPolicy();
//        }

//        [TestMethod]
//        public void GetExecutionPolicy_WithSecurityException_ReturnsAccessDenied()
//        {
//            // Arrange
//            _mockRegistry.ThrowOnGetPolicy = true;

//            // Act
//            string result = ExecutionFeasible.GetExecutionPolicy();

//            // Assert
//            Assert.AreEqual("AccessDenied", result, "发生安全异常时应该返回 AccessDenied");
//        }

//        #endregion

//        #region 修复功能测试

//        [TestMethod]
//        public void FixExecutionPolicy_SetsCorrectPolicy()
//        {
//            // Arrange
//            _mockRegistry.SetExecutionPolicyValue("Restricted");

//            // Act
//            ExecutionFeasible.FixExecutionPolicy();

//            // Assert
//            Assert.AreEqual("RemoteSigned", _mockRegistry.LastSetPolicy, "应该设置为 RemoteSigned 策略");
//        }

//        [TestMethod]
//        public void FixExecutionPolicy_AsAdmin_Succeeds()
//        {
//            // Arrange
//            _mockRegistry.SetIsAdministrator(true);

//            // Act
//            ExecutionFeasible.FixExecutionPolicy();

//            // Assert
//            Assert.AreEqual("RemoteSigned", _mockRegistry.LastSetPolicy, "管理员应该能够成功修复执行策略");
//        }

//        #endregion
//    }

//    #region 测试替身类

//    /// <summary>
//    /// 用于测试的注册表服务模拟类
//    /// </summary>
//    internal class MockRegistryService : ExecutionFeasible.IRegistryService
//    {
//        private string _executionPolicy = "Restricted";
//        private bool _isAdmin = false;

//        public bool ThrowOnGetPolicy { get; set; } = false;
//        public bool ThrowOnSetPolicy { get; set; } = false;
//        public string LastSetPolicy { get; private set; }

//        public void SetExecutionPolicyValue(string policy)
//        {
//            _executionPolicy = policy;
//        }

//        public void SetIsAdministrator(bool isAdmin)
//        {
//            _isAdmin = isAdmin;
//        }

//        public string GetExecutionPolicy()
//        {
//            if (ThrowOnGetPolicy)
//                return "AccessDenied";

//            return _executionPolicy;
//        }

//        public void SetExecutionPolicy(string policy)
//        {
//            if (ThrowOnSetPolicy)
//                throw new SecurityException("模拟的安全异常");

//            LastSetPolicy = policy;
//            _executionPolicy = policy;
//        }

//        public bool IsAdministrator()
//        {
//            return _isAdmin;
//        }
//    }

//    #endregion
//}
