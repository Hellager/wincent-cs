using System;
using System.IO;
using System.Runtime.InteropServices;

namespace Wincent
{
    internal interface IWindowsRecentFolder
    {
        string GetPath();
    }

    internal sealed class WindowsRecentFolder : IWindowsRecentFolder
    {
        private readonly INativeMethods _nativeMethods;
        private readonly IFileSystemOperations _fileSystem;

        public WindowsRecentFolder(INativeMethods nativeMethods, IFileSystemOperations fileSystem = null)
        {
            _nativeMethods = nativeMethods ?? throw new ArgumentNullException(nameof(nativeMethods));
            _fileSystem = fileSystem;
        }

        public string GetPath()
        {
            string path = TryGetKnownFolderPath();
            if (string.IsNullOrWhiteSpace(path))
                path = Environment.GetFolderPath(Environment.SpecialFolder.Recent);

            if (string.IsNullOrWhiteSpace(path))
                throw new InvalidPathException("Windows Recent Folder", "Unable to retrieve Windows Recent folder path.");

            if (!DirectoryExists(path))
                throw new InvalidPathException(path, "Windows Recent folder does not exist.");

            return path;
        }

        private string TryGetKnownFolderPath()
        {
            IntPtr pPath = IntPtr.Zero;
            try
            {
                int hr = _nativeMethods.SHGetKnownFolderPath(NativeMethods.FOLDERID_Recent, 0, IntPtr.Zero, out pPath);
                if (hr < 0 || pPath == IntPtr.Zero)
                    return null;

                return Marshal.PtrToStringUni(pPath);
            }
            finally
            {
                if (pPath != IntPtr.Zero)
                    _nativeMethods.CoTaskMemFree(pPath);
            }
        }

        private bool DirectoryExists(string path)
        {
            return _fileSystem == null ? Directory.Exists(path) : _fileSystem.DirectoryExists(path);
        }
    }
}
