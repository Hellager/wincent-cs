using System;
using System.Runtime.InteropServices;

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
        private readonly INativeMethods _nativeMethods;
        private readonly IShellApplicationFactory _shellApplicationFactory;

        public ShellQuickAccessNativeMutation(INativeMethods nativeMethods)
            : this(nativeMethods, new DefaultShellApplicationFactory())
        {
        }

        internal ShellQuickAccessNativeMutation(
            INativeMethods nativeMethods,
            IShellApplicationFactory shellApplicationFactory)
        {
            _nativeMethods = nativeMethods ?? throw new ArgumentNullException(nameof(nativeMethods));
            _shellApplicationFactory = shellApplicationFactory ?? throw new ArgumentNullException(nameof(shellApplicationFactory));
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
                        return;

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

                    if (!exists)
                        throw new QuickAccessItemNotFoundException(path, QuickAccess.FrequentFolders);

                    try
                    {
                        bool applied = false;
                        ForEachItemInNamespace(ShellNamespaces.FrequentFolders, (dynamic item) =>
                        {
                            if (WindowsPathComparer.Equals(item.Path as string, path))
                            {
                                item.InvokeVerb("unpinfromhome");
                                applied = true;
                                return true;
                            }

                            return false;
                        });

                        if (applied)
                            return;
                    }
                    catch (Exception)
                    {
                    }

                    InvokeVerbOnSelfFolder(path, "pintohome");
                },
                timeout,
                _nativeMethods);
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
