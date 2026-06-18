using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

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

    internal interface IShortcutResolutionResolver : IShortcutTargetResolver
    {
        ShortcutResolution Resolve(string shortcutPath, TimeSpan timeout);
    }

    internal interface IRecentLinkFileSystem
    {
        IEnumerable<string> EnumerateFiles(string directory);

        void DeleteFile(string path);
    }

    internal sealed class ShortcutResolution
    {
        public ShortcutResolution(string path, bool? isDirectory)
        {
            Path = path;
            IsDirectory = isDirectory;
        }

        public string Path { get; }

        public bool? IsDirectory { get; }
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

    internal sealed class ShellLinkTargetResolver : IShortcutResolutionResolver
    {
        private const uint FILE_ATTRIBUTE_DIRECTORY = 0x00000010;
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
            var resolution = Resolve(shortcutPath, timeout);
            return resolution?.Path;
        }

        public ShortcutResolution Resolve(string shortcutPath, TimeSpan timeout)
        {
            if (string.IsNullOrWhiteSpace(shortcutPath))
                return null;

            ShortcutResolution parsed = ShellLinkParser.ResolveFile(shortcutPath, timeout);
            if (parsed != null)
                return parsed;

            try
            {
                string target = StaThreadRunner.Run(() => ResolveTargetOnSta(shortcutPath), timeout, _nativeMethods);
                return string.IsNullOrWhiteSpace(target)
                    ? null
                    : new ShortcutResolution(target, DetermineTargetIsDirectory(target, 0, false, timeout));
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

        internal static bool? DetermineTargetIsDirectory(string path, uint attributes, bool targetIsNetwork, TimeSpan timeout)
        {
            if (attributes != 0)
                return (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;

            if (ShellLinkParser.ShouldTimeoutProtectMetadata(path, targetIsNetwork))
                return MetadataIsDirectoryWithTimeout(path, timeout);

            try
            {
                if (Directory.Exists(path))
                    return true;
                if (File.Exists(path))
                    return false;
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }

            return null;
        }

        private static bool? MetadataIsDirectoryWithTimeout(string path, TimeSpan timeout)
        {
            if (timeout <= TimeSpan.Zero)
                return null;

            bool? result = null;
            Exception error = null;
            var thread = new Thread(() =>
            {
                try
                {
                    if (Directory.Exists(path))
                        result = true;
                    else if (File.Exists(path))
                        result = false;
                }
                catch (Exception ex)
                {
                    error = ex;
                }
            });
            thread.IsBackground = true;
            thread.Start();

            if (!thread.Join(timeout))
                return null;

            return error == null ? result : null;
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

    internal static class ShellLinkParser
    {
        private const uint FILE_ATTRIBUTE_DIRECTORY = 0x00000010;
        private const uint SHELL_LINK_HEADER_SIZE = 0x4c;
        private const uint HAS_LINK_TARGET_ID_LIST = 0x00000001;
        private const uint HAS_LINK_INFO = 0x00000002;
        private const uint HAS_NAME = 0x00000004;
        private const uint HAS_RELATIVE_PATH = 0x00000008;
        private const uint HAS_WORKING_DIR = 0x00000010;
        private const uint HAS_ARGUMENTS = 0x00000020;
        private const uint HAS_ICON_LOCATION = 0x00000040;
        private const uint IS_UNICODE = 0x00000080;
        private const uint FORCE_NO_LINK_INFO = 0x00000200;
        private const uint VOLUME_ID_AND_LOCAL_BASE_PATH = 0x00000001;
        private const uint COMMON_NETWORK_RELATIVE_LINK_AND_PATH_SUFFIX = 0x00000002;

        private static readonly byte[] LinkClsid =
        {
            0x01, 0x14, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
            0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46
        };

        public static ShortcutResolution ResolveFile(string shortcutPath, TimeSpan timeout)
        {
            byte[] data;
            try
            {
                data = File.ReadAllBytes(shortcutPath);
            }
            catch (FileNotFoundException)
            {
                return null;
            }
            catch (DirectoryNotFoundException)
            {
                return null;
            }
            catch (IOException)
            {
                return null;
            }
            catch (UnauthorizedAccessException)
            {
                return null;
            }

            return ResolveBytes(data, timeout);
        }

        internal static ShortcutResolution ResolveBytes(byte[] data, TimeSpan timeout)
        {
            ShellLinkSummary summary = ParseShellLinkSummary(data);
            if (summary == null)
                return null;

            string path = summary.TargetPath ?? summary.RelativePath;
            if (string.IsNullOrWhiteSpace(path))
                return null;

            bool? isDirectory = ShellLinkTargetResolver.DetermineTargetIsDirectory(
                path,
                summary.FileAttributes,
                summary.TargetIsNetwork,
                timeout);
            return new ShortcutResolution(path, isDirectory);
        }

        internal static ShellLinkSummary ParseShellLinkSummary(byte[] data)
        {
            if (!LooksLikeLnk(data))
                return null;

            uint flags = ReadUInt32(data, 0x14);
            uint fileAttributes = data.Length >= 0x1c ? ReadUInt32(data, 0x18) : 0;
            int offset = checked((int)SHELL_LINK_HEADER_SIZE);

            if ((flags & HAS_LINK_TARGET_ID_LIST) != 0)
            {
                if (offset + 2 > data.Length)
                    return null;

                int idListSize = ReadUInt16(data, offset);
                offset = checked(offset + 2 + idListSize);
                if (offset > data.Length)
                    return null;
            }

            LinkInfoSummary linkInfo = null;
            if ((flags & HAS_LINK_INFO) != 0 &&
                (flags & FORCE_NO_LINK_INFO) == 0 &&
                offset + 28 <= data.Length)
            {
                linkInfo = ParseLinkInfo(data, offset);
                if (linkInfo != null)
                    offset = Math.Min(data.Length, offset + linkInfo.Size);
            }

            string relativePath = null;
            if ((flags & HAS_NAME) != 0)
            {
                ReadLnkString(data, offset, flags, out _, out offset);
            }
            if ((flags & HAS_RELATIVE_PATH) != 0)
            {
                ReadLnkString(data, offset, flags, out relativePath, out offset);
            }
            if ((flags & HAS_WORKING_DIR) != 0)
            {
                ReadLnkString(data, offset, flags, out _, out offset);
            }
            if ((flags & HAS_ARGUMENTS) != 0)
            {
                ReadLnkString(data, offset, flags, out _, out offset);
            }
            if ((flags & HAS_ICON_LOCATION) != 0)
            {
                ReadLnkString(data, offset, flags, out _, out _);
            }

            string targetPath =
                linkInfo?.ResolvedPath ??
                linkInfo?.LocalBasePath ??
                linkInfo?.NetworkPath ??
                linkInfo?.CommonPathSuffix;
            bool targetIsNetwork =
                (linkInfo != null && !string.IsNullOrEmpty(linkInfo.NetworkPath)) ||
                LooksLikeUncPath(targetPath) ||
                LooksLikeUncPath(relativePath);

            return new ShellLinkSummary(targetPath, fileAttributes, relativePath, targetIsNetwork);
        }

        internal static bool ShouldTimeoutProtectMetadata(string path, bool targetIsNetwork)
        {
            return targetIsNetwork || LooksLikeUncPath(path);
        }

        private static bool LooksLikeLnk(byte[] data)
        {
            if (data == null || data.Length < SHELL_LINK_HEADER_SIZE || ReadUInt32(data, 0) != SHELL_LINK_HEADER_SIZE)
                return false;

            for (int i = 0; i < LinkClsid.Length; i++)
            {
                if (data[4 + i] != LinkClsid[i])
                    return false;
            }

            return true;
        }

        private static LinkInfoSummary ParseLinkInfo(byte[] data, int linkInfoStart)
        {
            uint linkInfoSizeUInt = ReadUInt32(data, linkInfoStart);
            uint linkInfoHeaderSizeUInt = ReadUInt32(data, linkInfoStart + 4);
            if (linkInfoSizeUInt > int.MaxValue || linkInfoHeaderSizeUInt > int.MaxValue)
                return null;

            int linkInfoSize = (int)linkInfoSizeUInt;
            int linkInfoHeaderSize = (int)linkInfoHeaderSizeUInt;
            int linkInfoEnd = checked(linkInfoStart + linkInfoSize);
            if (linkInfoSize < 28 ||
                linkInfoHeaderSize < 28 ||
                linkInfoHeaderSize > linkInfoSize ||
                linkInfoEnd > data.Length)
            {
                return null;
            }

            uint linkInfoFlags = ReadUInt32(data, linkInfoStart + 8);
            int localBaseOffset = checked((int)ReadUInt32(data, linkInfoStart + 16));
            int networkOffset = checked((int)ReadUInt32(data, linkInfoStart + 20));
            int commonSuffixOffset = checked((int)ReadUInt32(data, linkInfoStart + 24));
            int localBaseUnicodeOffset = linkInfoHeaderSize >= 0x24 && linkInfoStart + 32 <= data.Length
                ? checked((int)ReadUInt32(data, linkInfoStart + 28))
                : 0;
            int commonSuffixUnicodeOffset = linkInfoHeaderSize >= 0x24 && linkInfoStart + 36 <= data.Length
                ? checked((int)ReadUInt32(data, linkInfoStart + 32))
                : 0;

            string localBasePath = null;
            if ((linkInfoFlags & VOLUME_ID_AND_LOCAL_BASE_PATH) != 0)
            {
                localBasePath =
                    ReadUtf16ZStringInLinkInfo(data, linkInfoStart, linkInfoSize, localBaseUnicodeOffset) ??
                    ReadCStringInLinkInfo(data, linkInfoStart, linkInfoSize, localBaseOffset);
            }

            string commonPathSuffix =
                ReadUtf16ZStringInLinkInfo(data, linkInfoStart, linkInfoSize, commonSuffixUnicodeOffset) ??
                ReadCStringInLinkInfo(data, linkInfoStart, linkInfoSize, commonSuffixOffset);

            NetworkLinkSummary network = (linkInfoFlags & COMMON_NETWORK_RELATIVE_LINK_AND_PATH_SUFFIX) != 0
                ? ParseCommonNetworkRelativeLink(data, linkInfoStart, linkInfoSize, networkOffset)
                : null;
            string networkPath = network?.NetName;

            string localResolvedPath = ResolveLocalPath(localBasePath, commonPathSuffix);
            string networkResolvedPath = ResolveNetworkPath(networkPath, commonPathSuffix);

            return new LinkInfoSummary(
                linkInfoSize,
                localBasePath,
                commonPathSuffix,
                networkPath,
                localResolvedPath ?? networkResolvedPath);
        }

        private static NetworkLinkSummary ParseCommonNetworkRelativeLink(
            byte[] data,
            int linkInfoStart,
            int linkInfoSize,
            int relativeOffset)
        {
            if (relativeOffset == 0 || relativeOffset + 20 > linkInfoSize)
                return null;

            int start = checked(linkInfoStart + relativeOffset);
            int size = checked((int)ReadUInt32(data, start));
            if (size < 20 || relativeOffset + size > linkInfoSize)
                return null;

            int netNameOffset = checked((int)ReadUInt32(data, start + 8));
            int netNameUnicodeOffset = netNameOffset > 0x14
                ? checked((int)ReadUInt32(data, start + 20))
                : 0;
            string netName =
                ReadUtf16ZStringInLinkInfo(data, start, size, netNameUnicodeOffset) ??
                ReadCStringInLinkInfo(data, start, size, netNameOffset);

            return new NetworkLinkSummary(netName);
        }

        private static string ResolveLocalPath(string localBasePath, string commonPathSuffix)
        {
            if (!string.IsNullOrEmpty(localBasePath) && !string.IsNullOrEmpty(commonPathSuffix))
                return JoinWindowsPath(localBasePath, commonPathSuffix);

            if (LooksLikeWindowsPath(localBasePath))
                return localBasePath;

            return LooksLikeWindowsPath(commonPathSuffix) ? commonPathSuffix : null;
        }

        private static string ResolveNetworkPath(string networkPath, string commonPathSuffix)
        {
            if (!string.IsNullOrEmpty(networkPath) && !string.IsNullOrEmpty(commonPathSuffix))
                return JoinWindowsPath(networkPath, commonPathSuffix);

            return LooksLikeUncPath(networkPath) ? networkPath : null;
        }

        private static string ReadCStringInLinkInfo(byte[] data, int linkInfoStart, int linkInfoSize, int relativeOffset)
        {
            if (relativeOffset == 0 || relativeOffset >= linkInfoSize)
                return null;

            int absoluteOffset = checked(linkInfoStart + relativeOffset);
            int linkInfoEnd = checked(linkInfoStart + linkInfoSize);
            return ReadCString(data, absoluteOffset, linkInfoEnd);
        }

        private static string ReadUtf16ZStringInLinkInfo(byte[] data, int linkInfoStart, int linkInfoSize, int relativeOffset)
        {
            if (relativeOffset == 0 || relativeOffset >= linkInfoSize)
                return null;

            int absoluteOffset = checked(linkInfoStart + relativeOffset);
            int linkInfoEnd = checked(linkInfoStart + linkInfoSize);
            return ReadUtf16ZString(data, absoluteOffset, linkInfoEnd);
        }

        private static void ReadLnkString(byte[] data, int offset, uint flags, out string value, out int nextOffset)
        {
            value = null;
            nextOffset = data.Length;
            if (offset + 2 > data.Length)
                return;

            int chars = ReadUInt16(data, offset);
            int stringStart = offset + 2;
            int byteCount = (flags & IS_UNICODE) != 0 ? chars * 2 : chars;
            int stringEnd = stringStart + byteCount;
            if (stringEnd > data.Length)
                return;

            value = (flags & IS_UNICODE) != 0
                ? DecodeUtf16Lossy(data, stringStart, byteCount)
                : Encoding.UTF8.GetString(data, stringStart, byteCount);
            nextOffset = stringEnd;
        }

        private static string ReadCString(byte[] data, int offset, int limit)
        {
            if (offset >= data.Length || offset >= limit)
                return null;

            int end = offset;
            while (end < data.Length && end < limit && data[end] != 0)
                end++;

            if (end >= data.Length || end >= limit)
                return null;

            return Encoding.UTF8.GetString(data, offset, end - offset);
        }

        private static string ReadUtf16ZString(byte[] data, int offset, int limit)
        {
            if (offset + 1 >= data.Length || offset + 1 >= limit)
                return null;

            int end = offset;
            while (end + 1 < data.Length && end + 1 < limit)
            {
                if (data[end] == 0 && data[end + 1] == 0)
                    break;
                end += 2;
            }

            if (end + 1 >= data.Length || end + 1 >= limit)
                return null;

            return DecodeUtf16Lossy(data, offset, end - offset);
        }

        private static string DecodeUtf16Lossy(byte[] data, int offset, int length)
        {
            if (data == null || length <= 0)
                return string.Empty;

            int usableLength = length - length % 2;
            return usableLength <= 0 ? string.Empty : Encoding.Unicode.GetString(data, offset, usableLength);
        }

        private static string JoinWindowsPath(string basePath, string suffixPath)
        {
            if (LooksLikeWindowsPath(suffixPath) || LooksLikeUncPath(suffixPath))
                return suffixPath;

            if (basePath.EndsWith("\\", StringComparison.Ordinal) ||
                basePath.EndsWith("/", StringComparison.Ordinal) ||
                suffixPath.Length == 0)
            {
                return basePath + suffixPath;
            }

            return basePath + "\\" + suffixPath;
        }

        private static bool LooksLikeWindowsPath(string value)
        {
            return !string.IsNullOrEmpty(value) &&
                   value.Length >= 3 &&
                   value[1] == ':' &&
                   (value[2] == '\\' || value[2] == '/');
        }

        private static bool LooksLikeUncPath(string value)
        {
            return !string.IsNullOrEmpty(value) &&
                   (value.StartsWith(@"\\", StringComparison.Ordinal) ||
                    value.StartsWith("//", StringComparison.Ordinal));
        }

        private static ushort ReadUInt16(byte[] data, int offset)
        {
            return BitConverter.ToUInt16(data, offset);
        }

        private static uint ReadUInt32(byte[] data, int offset)
        {
            return BitConverter.ToUInt32(data, offset);
        }

        internal sealed class ShellLinkSummary
        {
            public ShellLinkSummary(string targetPath, uint fileAttributes, string relativePath, bool targetIsNetwork)
            {
                TargetPath = targetPath;
                FileAttributes = fileAttributes;
                RelativePath = relativePath;
                TargetIsNetwork = targetIsNetwork;
            }

            public string TargetPath { get; }

            public uint FileAttributes { get; }

            public string RelativePath { get; }

            public bool TargetIsNetwork { get; }
        }

        private sealed class LinkInfoSummary
        {
            public LinkInfoSummary(
                int size,
                string localBasePath,
                string commonPathSuffix,
                string networkPath,
                string resolvedPath)
            {
                Size = size;
                LocalBasePath = localBasePath;
                CommonPathSuffix = commonPathSuffix;
                NetworkPath = networkPath;
                ResolvedPath = resolvedPath;
            }

            public int Size { get; }

            public string LocalBasePath { get; }

            public string CommonPathSuffix { get; }

            public string NetworkPath { get; }

            public string ResolvedPath { get; }
        }

        private sealed class NetworkLinkSummary
        {
            public NetworkLinkSummary(string netName)
            {
                NetName = netName;
            }

            public string NetName { get; }
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
