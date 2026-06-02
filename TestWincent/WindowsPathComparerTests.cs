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

        [TestMethod]
        public void Equals_DriveRoots_CompareCorrectly()
        {
            Assert.IsTrue(WindowsPathComparer.Equals(@"C:\", @"C:\"));
            Assert.IsTrue(WindowsPathComparer.Equals(@"C:\", @"C:/"));
            Assert.IsFalse(WindowsPathComparer.Equals(@"C:\", @"D:\"));
        }

        [TestMethod]
        public void Equals_UncRoots_CompareCorrectly()
        {
            Assert.IsTrue(WindowsPathComparer.Equals(@"\\server\share\", @"\\server\share"));
            Assert.IsTrue(WindowsPathComparer.Equals(@"\\server\share\", @"//server/share/"));
            Assert.IsFalse(WindowsPathComparer.Equals(@"\\server\share\", @"\\server\other\"));
        }

        [TestMethod]
        public void Equals_ResolvesToRoot()
        {
            Assert.IsTrue(WindowsPathComparer.Equals(@"C:\Folder\..", @"C:\"));
            Assert.IsFalse(WindowsPathComparer.Equals(@"C:\Folder\..", @"C:\not-root"));
        }
    }
}
