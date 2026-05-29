using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class NativeFileHandleTests
    {
        [TestMethod]
        public void OpenExistingForBackingFileLock_ExistingFile_ReturnsValidHandle()
        {
            string path = Path.GetTempFileName();
            try
            {
                using (var handle = NativeFileHandle.OpenExistingForBackingFileLock(path))
                {
                    Assert.IsFalse(handle.Handle.IsInvalid);
                    Assert.IsFalse(handle.Handle.IsClosed);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [TestMethod]
        public void OpenExistingForBackingFileLock_MissingFile_ThrowsFileNotFoundException()
        {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".tmp");

            Assert.ThrowsException<FileNotFoundException>(
                () => NativeFileHandle.OpenExistingForBackingFileLock(path));
        }

        [TestMethod]
        public void OpenExistingForBackingFileLock_EmptyPath_ThrowsArgumentException()
        {
            Assert.ThrowsException<ArgumentException>(
                () => NativeFileHandle.OpenExistingForBackingFileLock(string.Empty));
        }
    }
}
