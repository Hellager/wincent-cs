using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace Wincent
{
    internal sealed class NativeFileHandle : IDisposable
    {
        private const uint GenericRead = 0x80000000;
        private const uint FileShareRead = 0x00000001;
        private const uint OpenExisting = 3;
        private const uint FileAttributeNormal = 0x00000080;
        private const int ErrorFileNotFound = 2;
        private const int ErrorPathNotFound = 3;
        private const int ErrorAccessDenied = 5;

        private NativeFileHandle(SafeFileHandle handle)
        {
            Handle = handle;
        }

        public SafeFileHandle Handle { get; }

        public static NativeFileHandle OpenExistingForBackingFileLock(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Path cannot be empty.", nameof(path));

            SafeFileHandle handle = CreateFileW(
                path,
                GenericRead,
                FileShareRead,
                IntPtr.Zero,
                OpenExisting,
                FileAttributeNormal,
                IntPtr.Zero);

            if (handle.IsInvalid)
            {
                int errorCode = Marshal.GetLastWin32Error();
                throw CreateOpenException(path, errorCode);
            }

            return new NativeFileHandle(handle);
        }

        public void Dispose()
        {
            Handle.Dispose();
        }

        private static Exception CreateOpenException(string path, int errorCode)
        {
            if (errorCode == ErrorFileNotFound || errorCode == ErrorPathNotFound)
                return new FileNotFoundException($"Backing file not found: {path}", path);

            if (errorCode == ErrorAccessDenied)
                return new UnauthorizedAccessException($"Access denied opening backing file: {path}");

            return new IOException($"Failed to open backing file: {path}", new Win32Exception(errorCode));
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFileW(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);
    }
}
