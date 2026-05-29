using Microsoft.VisualStudio.TestTools.UnitTesting;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class WindowsPathComparerTests
    {
        [TestMethod]
        public void Equals_IgnoresCaseSlashStyleAndTrailingSlash()
        {
            Assert.IsTrue(WindowsPathComparer.Equals(@"C:\Folder\", "c:/folder"));
        }

        [TestMethod]
        public void Equals_UsesFullPathWhenAvailable()
        {
            Assert.IsTrue(WindowsPathComparer.Equals(@"C:\Folder\Sub\..\File.txt", @"c:\folder\file.txt"));
        }

        [TestMethod]
        public void Equals_FallsBackForInvalidPaths()
        {
            Assert.IsTrue(WindowsPathComparer.Equals(@"C:\Invalid|Path\", @"c:/invalid|path"));
        }

        [TestMethod]
        public void Equals_DifferentPaths_ReturnsFalse()
        {
            Assert.IsFalse(WindowsPathComparer.Equals(@"C:\One", @"C:\Two"));
        }
    }
}
