using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Wincent;

namespace TestWincent
{
    // ExecutionFeasibleTests.cs
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Microsoft.Win32;
    using System.Security;

    [TestClass]
    public class ExecutionFeasibleTests
    {
        [DataTestMethod]
        [DataRow("RemoteSigned", true)]
        [DataRow("restrictED", false)]
        public void CheckScriptFeasible_ValidatesPolicy(string policy, bool expected)
        {
            var result = ExecutionFeasible.CheckScriptFeasible();
            Assert.AreEqual(expected, result);
        }

        [TestMethod]
        [ExpectedException(typeof(SecurityException))]
        public void FixExecutionPolicy_RequiresAdmin()
        {
            if (!ExecutionFeasible.IsAdministrator())
            {
                ExecutionFeasible.FixExecutionPolicy();
            }
            else
            {
                Assert.Inconclusive("Run test in non-admin mode");
            }
        }

        [TestMethod]
        public void GetExecutionPolicy_ReturnsValidValues()
        {
            var policy = ExecutionFeasible.GetExecutionPolicy();
            string[] validPolicies = new[] { "AllSigned", "Bypass", "RemoteSigned", "NotSet", "AccessDenied" };

            bool containsValidPolicy = validPolicies.Any(policy.Contains);
            Assert.IsTrue(containsValidPolicy);
        }
    }

}
