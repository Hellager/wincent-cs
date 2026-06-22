using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Reflection;

namespace TestWincent
{
    [TestClass]
    public class TestCoverageTests
    {
        [TestMethod]
        public void IntegrationTests_AreIgnoredAndCategorized()
        {
            var offenders = typeof(TestCoverageTests).Assembly
                .GetTypes()
                .SelectMany(type => type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static))
                .Where(method => method.Name.StartsWith("Integration_", StringComparison.Ordinal))
                .Where(method => HasAttribute<TestMethodAttribute>(method))
                .Where(method => !HasAttribute<IgnoreAttribute>(method) || !HasIntegrationCategory(method))
                .Select(method => $"{method.DeclaringType.FullName}.{method.Name}")
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToList();

            if (offenders.Count > 0)
            {
                Assert.Fail(
                    "Integration tests must be both ignored and categorized as Integration: " +
                    string.Join(", ", offenders));
            }
        }

        private static bool HasAttribute<TAttribute>(MethodInfo method)
            where TAttribute : Attribute
        {
            return method.GetCustomAttributes(typeof(TAttribute), inherit: false).Any();
        }

        private static bool HasIntegrationCategory(MethodInfo method)
        {
            return method
                .GetCustomAttributes(typeof(TestCategoryAttribute), inherit: false)
                .Cast<TestCategoryAttribute>()
                .Any(attribute => attribute.TestCategories.Contains("Integration"));
        }
    }
}
