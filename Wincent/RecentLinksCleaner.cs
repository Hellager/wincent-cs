using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace Wincent
{
    internal interface IRecentLinksCleaner
    {
        IReadOnlyList<string> DeleteForTarget(string targetPath, TimeSpan timeout);
    }

    internal interface IShortcutTargetResolver
    {
        string ResolveTarget(string shortcutPath, TimeSpan timeout);
    }

    internal interface IRecentLinkFileSystem
    {
        IEnumerable<string> EnumerateFiles(string directory);

        void DeleteFile(string path);
    }

    internal sealed class RecentLinksCleaner : IRecentLinksCleaner
    {
        private readonly IWindowsRecentFolder _recentFolder;
        private readonly IShortcutTargetResolver _resolver;
        private readonly IRecentLinkFileSystem _fileSystem;

        public RecentLinksCleaner(
            IWindowsRecentFolder recentFolder,
            IShortcutTargetResolver resolver,
            IRecentLinkFileSystem fileSystem)
        {
            _recentFolder = recentFolder ?? throw new ArgumentNullException(nameof(recentFolder));
            _resolver = resolver ?? throw new ArgumentNullException(nameof(resolver));
            _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));
        }

        public IReadOnlyList<string> DeleteForTarget(string targetPath, TimeSpan timeout)
        {
            if (targetPath == null)
                throw new ArgumentNullException(nameof(targetPath));
            if (timeout <= TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(timeout), "Timeout must be positive.");

            var deleted = new List<string>();
            string recentFolder = _recentFolder.GetPath();
            var stopwatch = Stopwatch.StartNew();

            foreach (var path in _fileSystem.EnumerateFiles(recentFolder))
            {
                if (!IsShortcutFile(path))
                    continue;

                TimeSpan remaining = timeout - stopwatch.Elapsed;
                if (remaining <= TimeSpan.Zero)
                    throw new TimeoutException($"Recent links cleanup timed out after {timeout.TotalSeconds:0.###} seconds.");

                string resolvedTarget = _resolver.ResolveTarget(path, remaining);
                if (string.IsNullOrWhiteSpace(resolvedTarget))
                    continue;

                if (!WindowsPathComparer.Equals(resolvedTarget, targetPath))
                    continue;

                _fileSystem.DeleteFile(path);
                deleted.Add(path);
            }

            return deleted.AsReadOnly();
        }

        internal static bool IsShortcutFile(string path)
        {
            return string.Equals(Path.GetExtension(path), ".lnk", StringComparison.OrdinalIgnoreCase);
        }
    }

    internal sealed class DefaultRecentLinkFileSystem : IRecentLinkFileSystem
    {
        public IEnumerable<string> EnumerateFiles(string directory)
        {
            return Directory.EnumerateFiles(directory);
        }

        public void DeleteFile(string path)
        {
            File.Delete(path);
        }
    }

    internal sealed class ShellLinkTargetResolver : IShortcutTargetResolver
    {
        private const uint STGM_READ = 0x00000000;
        private const uint SLGP_RAWPATH = 0x00000004;
        private const uint SLR_NO_UI = 0x00000001;
        private const uint SLR_NOUPDATE = 0x00000008;

        private readonly INativeMethods _nativeMethods;

        public ShellLinkTargetResolver(INativeMethods nativeMethods)
        {
            _nativeMethods = nativeMethods ?? throw new ArgumentNullException(nameof(nativeMethods));
        }

        public string ResolveTarget(string shortcutPath, TimeSpan timeout)
        {
            if (string.IsNullOrWhiteSpace(shortcutPath))
                return null;

            try
            {
                return StaThreadRunner.Run(() => ResolveTargetOnSta(shortcutPath), timeout, _nativeMethods);
            }
            catch (TimeoutException)
            {
                throw;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string ResolveTargetOnSta(string shortcutPath)
        {
            object shellLinkObject = new ShellLink();
            var shellLink = (IShellLinkW)shellLinkObject;
            var persistFile = (IPersistFile)shellLinkObject;

            persistFile.Load(shortcutPath, STGM_READ);
            shellLink.Resolve(IntPtr.Zero, SLR_NO_UI | SLR_NOUPDATE);

            var path = new StringBuilder(32768);
            shellLink.GetPath(path, path.Capacity, IntPtr.Zero, SLGP_RAWPATH);

            string target = path.ToString();
            return string.IsNullOrWhiteSpace(target) ? null : target;
        }

        [ComImport]
        [Guid("00021401-0000-0000-C000-000000000046")]
        private sealed class ShellLink
        {
        }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("000214F9-0000-0000-C000-000000000046")]
        private interface IShellLinkW
        {
            void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, IntPtr pfd, uint fFlags);

            void GetIDList(out IntPtr ppidl);

            void SetIDList(IntPtr pidl);

            void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cchMaxName);

            void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);

            void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);

            void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);

            void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);

            void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);

            void GetHotkey(out short pwHotkey);

            void SetHotkey(short wHotkey);

            void GetShowCmd(out int piShowCmd);

            void SetShowCmd(int iShowCmd);

            void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cchIconPath, out int piIcon);

            void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);

            void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);

            void Resolve(IntPtr hwnd, uint fFlags);

            void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
        }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("0000010b-0000-0000-C000-000000000046")]
        private interface IPersistFile
        {
            void GetClassID(out Guid pClassID);

            void IsDirty();

            void Load([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, uint dwMode);

            void Save([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, bool fRemember);

            void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string pszFileName);

            void GetCurFile([MarshalAs(UnmanagedType.LPWStr)] out string ppszFileName);
        }
    }

    internal sealed class NoOpRecentLinksCleaner : IRecentLinksCleaner
    {
        public IReadOnlyList<string> DeleteForTarget(string targetPath, TimeSpan timeout)
        {
            return Array.Empty<string>();
        }
    }
}
