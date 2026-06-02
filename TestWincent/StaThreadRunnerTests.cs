using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Threading;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class StaThreadRunnerTests
    {
        [TestMethod]
        public void Run_ReturnsResultFromStaThread()
        {
            var nativeMethods = CreateNativeMethods();

            var result = StaThreadRunner.Run(
                () => Thread.CurrentThread.GetApartmentState(),
                TimeSpan.FromSeconds(1),
                nativeMethods.Object);

            Assert.AreEqual(ApartmentState.STA, result);
            nativeMethods.Verify(n => n.CoUninitialize(), Times.Once);
        }

        [TestMethod]
        public void Run_PropagatesWorkerException()
        {
            var nativeMethods = CreateNativeMethods();

            Assert.ThrowsException<InvalidOperationException>(
                () => StaThreadRunner.Run(
                    () => { throw new InvalidOperationException("boom"); },
                    TimeSpan.FromSeconds(1),
                    nativeMethods.Object));
        }

        [TestMethod]
        public void Run_ZeroTimeout_ThrowsArgumentOutOfRangeException()
        {
            Assert.ThrowsException<ArgumentOutOfRangeException>(
                () => StaThreadRunner.Run(() => true, TimeSpan.Zero));
        }

        [TestMethod]
        public void Run_Timeout_ThrowsTimeoutException()
        {
            var nativeMethods = CreateNativeMethods();

            Assert.ThrowsException<TimeoutException>(
                () => StaThreadRunner.Run(
                    () =>
                    {
                        Thread.Sleep(500);
                        return true;
                    },
                    TimeSpan.FromMilliseconds(50),
                    nativeMethods.Object));
        }

        [TestMethod]
        public void Run_DisableOle1Dde_PassesFlagToComInitialization()
        {
            uint? coInit = null;
            var nativeMethods = new Mock<INativeMethods>(MockBehavior.Strict);
            nativeMethods.Setup(n => n.CoInitializeEx(IntPtr.Zero, It.IsAny<uint>()))
                .Callback<IntPtr, uint>((reserved, flags) => coInit = flags)
                .Returns(0);
            nativeMethods.Setup(n => n.CoUninitialize());

            StaThreadRunner.Run(
                () => { },
                TimeSpan.FromSeconds(1),
                nativeMethods.Object,
                disableOle1Dde: true);

            Assert.AreEqual(
                (uint)(NativeMethods.COINIT_APARTMENTTHREADED | NativeMethods.COINIT_DISABLE_OLE1DDE),
                coInit.Value);
        }

        private static Mock<INativeMethods> CreateNativeMethods()
        {
            var nativeMethods = new Mock<INativeMethods>(MockBehavior.Strict);
            nativeMethods.Setup(n => n.CoInitializeEx(IntPtr.Zero, It.IsAny<uint>())).Returns(0);
            nativeMethods.Setup(n => n.CoUninitialize());
            return nativeMethods;
        }
    }
}
