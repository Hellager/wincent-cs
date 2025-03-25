using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class QuickAccessManagerTests
    {
        private const string TestFilePath = @"C:\TestFile.txt";
        [TestInitialize]
        public void Setup()
        {
            File.WriteAllText(TestFilePath, "test");
        }
        [TestCleanup]
        public void Cleanup()
        {
            File.Delete(TestFilePath);
        }
        [TestMethod]
        [ExpectedException(typeof(SecurityException))]
        public async Task AddItemAsync_ValidatesProtectedPaths()
        {
            await QuickAccessManager.AddItemAsync(@"C:\Windows\System32\test.txt", QuickAccessItemType.File);
        }
        [TestMethod]
        public async Task RemoveItemAsync_HandlesInvalidPaths()
        {
            try
            {
                await QuickAccessManager.RemoveItemAsync("invalid_path", QuickAccessItemType.File);
                Assert.Fail("Expected exception not thrown");
            }
            catch (FileNotFoundException ex)
            {
                StringAssert.Contains(ex.Message, "invalid_path");
            }
        }
    }

}
