using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace Wincent
{
    internal interface IQuickAccessNativeMutation
    {
        void RemoveRecentFile(string path, TimeSpan timeout);

        void PinFrequentFolder(string path, TimeSpan timeout);

        void UnpinFrequentFolder(string path, TimeSpan timeout);
    }

    internal sealed class ShellQuickAccessNativeMutation : IQuickAccessNativeMutation
    {
        private const int VerificationPollIntervalMilliseconds = 100;
        private static readonly TimeSpan DefaultVerificationTimeout = TimeSpan.FromSeconds(1);

        private readonly INativeMethods _nativeMethods;
        private readonly IShellApplicationFactory _shellApplicationFactory;
        private readonly Func<bool> _isWindows11OrLater;
        private readonly TimeSpan _verificationTimeout;

        public ShellQuickAccessNativeMutation(INativeMethods nativeMethods)
            : this(nativeMethods, new DefaultShellApplicationFactory())
        {
        }

        internal ShellQuickAccessNativeMutation(
            INativeMethods nativeMethods,
            IShellApplicationFactory shellApplicationFactory)
            : this(nativeMethods, shellApplicationFactory, IsWindows11OrLater)
        {
        }

        internal ShellQuickAccessNativeMutation(
            INativeMethods nativeMethods,
            IShellApplicationFactory shellApplicationFactory,
            Func<bool> isWindows11OrLater)
            : this(nativeMethods, shellApplicationFactory, isWindows11OrLater, DefaultVerificationTimeout)
        {
        }

        internal ShellQuickAccessNativeMutation(
            INativeMethods nativeMethods,
            IShellApplicationFactory shellApplicationFactory,
            Func<bool> isWindows11OrLater,
            TimeSpan verificationTimeout)
        {
            _nativeMethods = nativeMethods ?? throw new ArgumentNullException(nameof(nativeMethods));
            _shellApplicationFactory = shellApplicationFactory ?? throw new ArgumentNullException(nameof(shellApplicationFactory));
            _isWindows11OrLater = isWindows11OrLater ?? throw new ArgumentNullException(nameof(isWindows11OrLater));
            if (verificationTimeout <= TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(verificationTimeout), "Verification timeout must be positive.");
            _verificationTimeout = verificationTimeout;
        }

        public void RemoveRecentFile(string path, TimeSpan timeout)
        {
            StaThreadRunner.Run(
                () =>
                {
                    bool found = false;
                    ForEachItemInNamespace(ShellNamespaces.QuickAccess, (dynamic item) =>
                    {
                        if (Convert.ToBoolean(item.IsFolder))
                            return false;

                        if (!WindowsPathComparer.Equals(item.Path as string, path))
                            return false;

                        item.InvokeVerb("remove");
                        found = true;
                        return true;
                    });

                    if (!found)
                        throw new QuickAccessItemNotFoundException(path, QuickAccess.RecentFiles);
                },
                timeout,
                _nativeMethods);
        }

        public void PinFrequentFolder(string path, TimeSpan timeout)
        {
            StaThreadRunner.Run(
                () =>
                {
                    bool exists = false;
                    ForEachItemInNamespace(ShellNamespaces.FrequentFolders, (dynamic item) =>
                    {
                        if (WindowsPathComparer.Equals(item.Path as string, path))
                        {
                            exists = true;
                            return true;
                        }

                        return false;
                    });

                    if (exists)
                        throw new QuickAccessItemAlreadyExistsException(path, QuickAccess.FrequentFolders);

                    InvokeVerbOnSelfFolder(path, "pintohome");
                },
                timeout,
                _nativeMethods);
        }

        public void UnpinFrequentFolder(string path, TimeSpan timeout)
        {
            StaThreadRunner.Run(
                () =>
                {
                    if (!ContainsFrequentFolder(path))
                        throw new QuickAccessItemNotFoundException(path, QuickAccess.FrequentFolders);

                    TryInvokeVerbOnFrequentFolder(path, "unpinfromhome");
                    if (WaitForFrequentFolderPresence(path, false))
                        return;

                    // Unpinned frequent folders ignore unpinfromhome. Pin first, then apply the platform-specific
                    // unpin verb so Explorer removes the item from the Frequent Folders namespace.
                    InvokeVerbOnSelfFolder(path, "pintohome");
                    if (WaitForFrequentFolderPresence(path, false))
                        return;

                    if (_isWindows11OrLater())
                        InvokeVerbOnSelfFolder(path, "pintohome");
                    else
                        TryInvokeVerbOnFrequentFolder(path, "unpinfromhome");

                    if (WaitForFrequentFolderPresence(path, false))
                        return;

                    throw new InvalidOperationException($"Failed to remove frequent folder: {path}");
                },
                timeout,
                _nativeMethods);
        }

        private static bool IsWindows11OrLater()
        {
            return Environment.OSVersion.Version.Build >= 22000;
        }

        private bool ContainsFrequentFolder(string path)
        {
            bool exists = false;
            ForEachItemInNamespace(ShellNamespaces.FrequentFolders, (dynamic item) =>
            {
                if (WindowsPathComparer.Equals(item.Path as string, path))
                {
                    exists = true;
                    return true;
                }

                return false;
            });

            return exists;
        }

        private bool WaitForFrequentFolderPresence(string path, bool expected)
        {
            var stopwatch = Stopwatch.StartNew();
            while (stopwatch.Elapsed < _verificationTimeout)
            {
                if (ContainsFrequentFolder(path) == expected)
                    return true;

                int remainingMilliseconds = (int)(_verificationTimeout - stopwatch.Elapsed).TotalMilliseconds;
                if (remainingMilliseconds <= 0)
                    break;

                Thread.Sleep(Math.Min(VerificationPollIntervalMilliseconds, remainingMilliseconds));
            }

            return ContainsFrequentFolder(path) == expected;
        }

        private bool TryInvokeVerbOnFrequentFolder(string path, string verb)
        {
            try
            {
                bool applied = false;
                ForEachItemInNamespace(ShellNamespaces.FrequentFolders, (dynamic item) =>
                {
                    if (WindowsPathComparer.Equals(item.Path as string, path))
                    {
                        item.InvokeVerb(verb);
                        applied = true;
                        return true;
                    }

                    return false;
                });

                return applied;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private void ForEachItemInNamespace(string @namespace, Func<dynamic, bool> action)
        {
            object shellApplication = null;
            object folder = null;
            object items = null;

            try
            {
                shellApplication = _shellApplicationFactory.CreateShellApplication();
                dynamic shell = shellApplication;
                folder = shell.Namespace(@namespace);
                if (folder == null)
                    throw new InvalidOperationException($"Failed to open shell namespace: {@namespace}");

                dynamic folderObj = folder;
                items = folderObj.Items();
                if (items == null)
                    throw new InvalidOperationException("Failed to enumerate shell folder items.");

                int count = Convert.ToInt32(((dynamic)items).Count);
                for (int index = 0; index < count; index++)
                {
                    object item = null;
                    try
                    {
                        item = ((dynamic)items).Item(index);
                    }
                    catch (COMException)
                    {
                        continue;
                    }

                    if (item == null)
                        continue;

                    try
                    {
                        if (action(item))
                            return;
                    }
                    finally
                    {
                        if (item != null && Marshal.IsComObject(item))
                            Marshal.FinalReleaseComObject(item);
                    }
                }
            }
            finally
            {
                if (items != null && Marshal.IsComObject(items))
                    Marshal.FinalReleaseComObject(items);
                if (folder != null && Marshal.IsComObject(folder))
                    Marshal.FinalReleaseComObject(folder);
                if (shellApplication != null && Marshal.IsComObject(shellApplication))
                    Marshal.FinalReleaseComObject(shellApplication);
            }
        }

        private void InvokeVerbOnSelfFolder(string path, string verb)
        {
            object shellApplication = null;
            object folder = null;
            object self = null;

            try
            {
                shellApplication = _shellApplicationFactory.CreateShellApplication();
                dynamic shell = shellApplication;
                folder = shell.Namespace(path);
                if (folder == null)
                    throw new InvalidOperationException($"Failed to open shell folder: {path}");

                dynamic folderObj = folder;
                self = folderObj.Self;
                if (self == null)
                    throw new InvalidOperationException($"Failed to get shell folder self item: {path}");

                ((dynamic)self).InvokeVerb(verb);
            }
            finally
            {
                if (self != null && Marshal.IsComObject(self))
                    Marshal.FinalReleaseComObject(self);
                if (folder != null && Marshal.IsComObject(folder))
                    Marshal.FinalReleaseComObject(folder);
                if (shellApplication != null && Marshal.IsComObject(shellApplication))
                    Marshal.FinalReleaseComObject(shellApplication);
            }
        }
    }

    internal sealed class PowerShellFallbackNativeMutation : IQuickAccessNativeMutation
    {
        public void RemoveRecentFile(string path, TimeSpan timeout)
        {
            throw new InvalidOperationException("Native Quick Access mutation is disabled for this instance.");
        }

        public void PinFrequentFolder(string path, TimeSpan timeout)
        {
            throw new InvalidOperationException("Native Quick Access mutation is disabled for this instance.");
        }

        public void UnpinFrequentFolder(string path, TimeSpan timeout)
        {
            throw new InvalidOperationException("Native Quick Access mutation is disabled for this instance.");
        }
    }
}
