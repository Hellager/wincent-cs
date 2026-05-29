using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Runtime.InteropServices;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class ComGuardTests
    {
        [TestMethod]
        public void ClassifyCoInitializeResult_MapsKnownResults()
        {
            Assert.AreEqual(ComInitStatus.Success, ComGuard.ClassifyCoInitializeResult(0));
            Assert.AreEqual(ComInitStatus.AlreadyInitialized, ComGuard.ClassifyCoInitializeResult(1));
            Assert.AreEqual(ComInitStatus.ApartmentMismatch, ComGuard.ClassifyCoInitializeResult(unchecked((int)0x80010106)));
            Assert.AreEqual(ComInitStatus.Failed, ComGuard.ClassifyCoInitializeResult(unchecked((int)0x80070057)));
            Assert.AreEqual(ComInitStatus.Success, ComGuard.ClassifyCoInitializeResult(2));
        }

        [TestMethod]
        public void InitializeSta_SOk_BalancesCoUninitialize()
        {
            var nativeMethods = CreateNativeMethods(0);

            using (ComGuard.InitializeSta(nativeMethods.Object))
            {
            }

            nativeMethods.Verify(n => n.CoUninitialize(), Times.Once);
        }

        [TestMethod]
        public void InitializeSta_SFalse_BalancesCoUninitialize()
        {
            var nativeMethods = CreateNativeMethods(1);

            using (ComGuard.InitializeSta(nativeMethods.Object))
            {
            }

            nativeMethods.Verify(n => n.CoUninitialize(), Times.Once);
        }

        [TestMethod]
        public void InitializeSta_ApartmentMismatch_ThrowsWithoutUninitialize()
        {
            var nativeMethods = CreateNativeMethods(unchecked((int)0x80010106));

            Assert.ThrowsException<ComApartmentMismatchException>(() => ComGuard.InitializeSta(nativeMethods.Object));
            nativeMethods.Verify(n => n.CoUninitialize(), Times.Never);
        }

        [TestMethod]
        public void InitializeSta_OtherFailure_ThrowsComExceptionWithoutUninitialize()
        {
            int hResult = unchecked((int)0x80070057);
            var nativeMethods = CreateNativeMethods(hResult);

            var exception = Assert.ThrowsException<COMException>(() => ComGuard.InitializeSta(nativeMethods.Object));
            Assert.AreEqual(hResult, exception.ErrorCode);
            nativeMethods.Verify(n => n.CoUninitialize(), Times.Never);
        }

        private static Mock<INativeMethods> CreateNativeMethods(int hResult)
        {
            var nativeMethods = new Mock<INativeMethods>(MockBehavior.Strict);
            nativeMethods.Setup(n => n.CoInitializeEx(IntPtr.Zero, It.IsAny<uint>())).Returns(hResult);
            nativeMethods.Setup(n => n.CoUninitialize());
            return nativeMethods;
        }
    }
}
