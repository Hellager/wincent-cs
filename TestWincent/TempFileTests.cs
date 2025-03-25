using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class TempFileTests
    {
        [TestMethod]
        public void TempFile_AutoDeletesOnDispose()
        {
            string path;
            using (var temp = TempFile.Create("test", "tmp"))
            {
                path = temp.FullPath;
                Assert.IsTrue(File.Exists(path));
            }
            Assert.IsFalse(File.Exists(path));
        }
        [DataTestMethod]
        [DataRow("noext")]
        [DataRow(".txt")]
        public void Create_HandlesVariousExtensions(string ext)
        {
            var file = TempFile.Create("content", ext);
            StringAssert.EndsWith(file.FullPath, ext);
        }
    }
}
