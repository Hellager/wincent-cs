using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.Threading;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class StaThreadRunnerTests
    {
        [TestMethod]
        public void MaxActiveStaWorkers_MatchesUpstreamWorkerCap()
        {
            Assert.AreEqual(4, StaThreadRunner.MaxActiveStaWorkers);
        }

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
            using (var release = new ManualResetEventSlim(false))
            {
                try
                {
                    Assert.ThrowsException<TimeoutException>(
                        () => StaThreadRunner.Run(
                            () =>
                            {
                                release.Wait();
                                return true;
                            },
                            TimeSpan.FromMilliseconds(50),
                            nativeMethods.Object));
                }
                finally
                {
                    release.Set();
                }

                Assert.IsTrue(WaitForActiveWorkerCount(0, TimeSpan.FromSeconds(2)));
            }
        }

        [TestMethod]
        public void Run_ActiveWorkerLimit_RejectsExtraWorkerAndRecovers()
        {
            Assert.IsTrue(WaitForActiveWorkerCount(0, TimeSpan.FromSeconds(2)));

            var nativeMethods = CreateNativeMethods();
            var release = new ManualResetEventSlim(false);
            var entered = new CountdownEvent(StaThreadRunner.MaxActiveStaWorkers);
            var callers = new List<Thread>();

            try
            {
                for (int i = 0; i < StaThreadRunner.MaxActiveStaWorkers; i++)
                {
                    var caller = new Thread(() =>
                    {
                        Assert.ThrowsException<TimeoutException>(
                            () => StaThreadRunner.Run(
                                () =>
                                {
                                    entered.Signal();
                                    release.Wait();
                                    return true;
                                },
                                TimeSpan.FromMilliseconds(250),
                                nativeMethods.Object));
                    });
                    caller.IsBackground = true;
                    callers.Add(caller);
                    caller.Start();
                }

                Assert.IsTrue(entered.Wait(TimeSpan.FromSeconds(5)), "All STA workers should enter the blocking closure.");
                Assert.IsTrue(WaitForActiveWorkerCount(StaThreadRunner.MaxActiveStaWorkers, TimeSpan.FromSeconds(2)));

                foreach (var caller in callers)
                    Assert.IsTrue(caller.Join(TimeSpan.FromSeconds(5)), "Caller should time out while worker remains active.");

                bool extraRan = false;
                var ex = Assert.ThrowsException<TimeoutException>(
                    () => StaThreadRunner.Run(
                        () =>
                        {
                            extraRan = true;
                            return true;
                        },
                        TimeSpan.FromSeconds(1),
                        nativeMethods.Object));

                StringAssert.Contains(ex.Message, "Too many Shell COM operations");
                Assert.IsFalse(extraRan, "The extra closure must not run when the active worker limit is reached.");
            }
            finally
            {
                release.Set();
                foreach (var caller in callers)
                    caller.Join(TimeSpan.FromSeconds(5));
                entered.Dispose();
                release.Dispose();
            }

            Assert.IsTrue(WaitForActiveWorkerCount(0, TimeSpan.FromSeconds(5)));
            Assert.IsTrue(StaThreadRunner.Run(() => true, TimeSpan.FromSeconds(1), nativeMethods.Object));
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

        private static bool WaitForActiveWorkerCount(int expected, TimeSpan timeout)
        {
            var started = DateTime.UtcNow;
            while (DateTime.UtcNow - started < timeout)
            {
                if (StaThreadRunner.ActiveWorkerCount == expected)
                    return true;

                Thread.Sleep(10);
            }

            return StaThreadRunner.ActiveWorkerCount == expected;
        }
    }
}
