using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Runtime.InteropServices;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class WindowsRecentFolderTests
    {
        private delegate void SHGetKnownFolderPathCallback(Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr ppszPath);

        [TestMethod]
        public void GetPath_UsesKnownFolderPathWhenAvailable()
        {
            var nativeMethods = new Mock<INativeMethods>(MockBehavior.Strict);
            var fileSystem = new Mock<IFileSystemOperations>(MockBehavior.Strict);
            IntPtr recentFolder = Marshal.StringToHGlobalUni(@"C:\Users\Test\Recent");

            nativeMethods.Setup(n => n.SHGetKnownFolderPath(
                    It.IsAny<Guid>(),
                    It.IsAny<uint>(),
                    It.IsAny<IntPtr>(),
                    out It.Ref<IntPtr>.IsAny))
                .Callback(new SHGetKnownFolderPathCallback((Guid id, uint flags, IntPtr token, out IntPtr path) =>
                {
                    path = recentFolder;
                }))
                .Returns(0);
            nativeMethods.Setup(n => n.CoTaskMemFree(recentFolder)).Callback<IntPtr>(Marshal.FreeHGlobal);
            fileSystem.Setup(f => f.DirectoryExists(@"C:\Users\Test\Recent")).Returns(true);

            var result = new WindowsRecentFolder(nativeMethods.Object, fileSystem.Object).GetPath();

            Assert.AreEqual(@"C:\Users\Test\Recent", result);
            nativeMethods.Verify(n => n.CoTaskMemFree(recentFolder), Times.Once);
        }

        [TestMethod]
        public void GetPath_KnownFolderFailure_FallsBackToEnvironmentRecentFolder()
        {
            var nativeMethods = new Mock<INativeMethods>(MockBehavior.Strict);
            var fileSystem = new Mock<IFileSystemOperations>(MockBehavior.Strict);
            IntPtr noPath = IntPtr.Zero;

            nativeMethods.Setup(n => n.SHGetKnownFolderPath(
                    It.IsAny<Guid>(),
                    It.IsAny<uint>(),
                    It.IsAny<IntPtr>(),
                    out noPath))
                .Returns(unchecked((int)0x80070002));
            fileSystem.Setup(f => f.DirectoryExists(It.IsAny<string>())).Returns(true);

            var result = new WindowsRecentFolder(nativeMethods.Object, fileSystem.Object).GetPath();

            Assert.IsFalse(string.IsNullOrWhiteSpace(result));
        }

        [TestMethod]
        public void GetPath_PositiveKnownFolderInformationCode_UsesReturnedPath()
        {
            var nativeMethods = new Mock<INativeMethods>(MockBehavior.Strict);
            var fileSystem = new Mock<IFileSystemOperations>(MockBehavior.Strict);
            IntPtr recentFolder = Marshal.StringToHGlobalUni(@"C:\Users\Test\Recent");

            nativeMethods.Setup(n => n.SHGetKnownFolderPath(
                    It.IsAny<Guid>(),
                    It.IsAny<uint>(),
                    It.IsAny<IntPtr>(),
                    out It.Ref<IntPtr>.IsAny))
                .Callback(new SHGetKnownFolderPathCallback((Guid id, uint flags, IntPtr token, out IntPtr path) =>
                {
                    path = recentFolder;
                }))
                .Returns(2);
            nativeMethods.Setup(n => n.CoTaskMemFree(recentFolder)).Callback<IntPtr>(Marshal.FreeHGlobal);
            fileSystem.Setup(f => f.DirectoryExists(@"C:\Users\Test\Recent")).Returns(true);

            var result = new WindowsRecentFolder(nativeMethods.Object, fileSystem.Object).GetPath();

            Assert.AreEqual(@"C:\Users\Test\Recent", result);
        }

        [TestMethod]
        public void GetPath_NonexistentKnownFolder_ThrowsInvalidPathException()
        {
            var nativeMethods = new Mock<INativeMethods>(MockBehavior.Strict);
            var fileSystem = new Mock<IFileSystemOperations>(MockBehavior.Strict);
            IntPtr recentFolder = Marshal.StringToHGlobalUni(@"C:\Users\Test\Recent");

            nativeMethods.Setup(n => n.SHGetKnownFolderPath(
                    It.IsAny<Guid>(),
                    It.IsAny<uint>(),
                    It.IsAny<IntPtr>(),
                    out It.Ref<IntPtr>.IsAny))
                .Callback(new SHGetKnownFolderPathCallback((Guid id, uint flags, IntPtr token, out IntPtr path) =>
                {
                    path = recentFolder;
                }))
                .Returns(0);
            nativeMethods.Setup(n => n.CoTaskMemFree(recentFolder)).Callback<IntPtr>(Marshal.FreeHGlobal);
            fileSystem.Setup(f => f.DirectoryExists(@"C:\Users\Test\Recent")).Returns(false);

            Assert.ThrowsException<InvalidPathException>(
                () => new WindowsRecentFolder(nativeMethods.Object, fileSystem.Object).GetPath());
        }
    }
}
