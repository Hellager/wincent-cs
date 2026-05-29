using System;
using System.Runtime.InteropServices;

namespace Wincent
{
    internal enum ComInitStatus
    {
        Success,
        AlreadyInitialized,
        ApartmentMismatch,
        Failed
    }

    internal sealed class ComGuard : IDisposable
    {
        private const int SFalse = 1;
        private const int RpcEChangedMode = unchecked((int)0x80010106);

        private readonly INativeMethods _nativeMethods;
        private bool _disposed;

        private ComGuard(INativeMethods nativeMethods)
        {
            _nativeMethods = nativeMethods;
        }

        public static ComGuard InitializeSta(INativeMethods nativeMethods, bool disableOle1Dde = false)
        {
            if (nativeMethods == null)
                throw new ArgumentNullException(nameof(nativeMethods));

            uint coInit = NativeMethods.COINIT_APARTMENTTHREADED;
            if (disableOle1Dde)
                coInit |= NativeMethods.COINIT_DISABLE_OLE1DDE;

            int hr = nativeMethods.CoInitializeEx(IntPtr.Zero, coInit);
            ComInitStatus status = ClassifyCoInitializeResult(hr);

            if (status == ComInitStatus.ApartmentMismatch)
                throw new ComApartmentMismatchException(hr);

            if (status == ComInitStatus.Failed)
                throw new COMException($"COM initialization failed with HRESULT 0x{hr:X8}.", hr);

            return new ComGuard(nativeMethods);
        }

        public static ComInitStatus ClassifyCoInitializeResult(int hResult)
        {
            if (hResult == 0)
                return ComInitStatus.Success;

            if (hResult == SFalse)
                return ComInitStatus.AlreadyInitialized;

            if (hResult == RpcEChangedMode)
                return ComInitStatus.ApartmentMismatch;

            return hResult < 0 ? ComInitStatus.Failed : ComInitStatus.Success;
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            _nativeMethods.CoUninitialize();
            _disposed = true;
        }
    }
}
