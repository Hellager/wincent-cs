using System;
using System.Collections.Generic;
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
                    var folder = OpenNamespace(ShellNamespaces.QuickAccess);
                    foreach (dynamic item in EnumerateItems(folder))
                    {
                        if (Convert.ToBoolean(item.IsFolder))
                            continue;

                        string itemPath = item.Path as string;
                        if (!WindowsPathComparer.Equals(itemPath, path))
                            continue;

                        item.InvokeVerb("remove");
                        return;
                    }

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
                    if (IsInFrequentFolders(path))
                        return;

                    InvokeVerbOnSelf(path, "pintohome");
                },
                timeout,
                _nativeMethods);
        }

        public void UnpinFrequentFolder(string path, TimeSpan timeout)
        {
            StaThreadRunner.Run(
                () =>
                {
                    if (!IsInFrequentFolders(path))
                        throw new QuickAccessItemNotFoundException(path, QuickAccess.FrequentFolders);

                    try
                    {
                        FindAndInvokeVerb(path, "unpinfromhome");
                        return;
                    }
                    catch (QuickAccessItemNotFoundException)
                    {
                        throw;
                    }
                    catch (Exception)
                    {
                        InvokeVerbOnSelf(path, "pintohome");
                    }
                },
                timeout,
                _nativeMethods);
        }

        private dynamic OpenNamespace(string @namespace)
        {
            dynamic shellApplication = _shellApplicationFactory.CreateShellApplication();
            dynamic folder = shellApplication.Namespace(@namespace);
            if (folder == null)
                throw new InvalidOperationException($"Failed to open shell namespace: {@namespace}");

            return folder;
        }

        private dynamic OpenFolder(string path)
        {
            dynamic shellApplication = _shellApplicationFactory.CreateShellApplication();
            dynamic folder = shellApplication.Namespace(path);
            if (folder == null)
                throw new InvalidOperationException($"Failed to open shell folder: {path}");

            return folder;
        }

        private IEnumerable<dynamic> EnumerateItems(dynamic folder)
        {
            dynamic items = folder.Items();
            if (items == null)
                throw new InvalidOperationException("Failed to enumerate shell folder items.");

            int count = Convert.ToInt32(items.Count);
            for (int index = 0; index < count; index++)
            {
                dynamic item;
                try
                {
                    item = items.Item(index);
                }
                catch (COMException)
                {
                    continue;
                }

                if (item != null)
                    yield return item;
            }
        }

        private bool IsInFrequentFolders(string path)
        {
            var folder = OpenNamespace(ShellNamespaces.FrequentFolders);
            foreach (dynamic item in EnumerateItems(folder))
            {
                string itemPath = item.Path as string;
                if (WindowsPathComparer.Equals(itemPath, path))
                    return true;
            }

            return false;
        }

        private void FindAndInvokeVerb(string path, string verb)
        {
            var folder = OpenNamespace(ShellNamespaces.FrequentFolders);
            foreach (dynamic item in EnumerateItems(folder))
            {
                string itemPath = item.Path as string;
                if (!WindowsPathComparer.Equals(itemPath, path))
                    continue;

                item.InvokeVerb(verb);
                return;
            }

            throw new QuickAccessItemNotFoundException(path, QuickAccess.FrequentFolders);
        }

        private void InvokeVerbOnSelf(string path, string verb)
        {
            dynamic folder = OpenFolder(path);
            dynamic self = folder.Self;
            if (self == null)
                throw new InvalidOperationException($"Failed to get shell folder self item: {path}");

            self.InvokeVerb(verb);
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
