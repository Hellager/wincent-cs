using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Wincent
{
    internal interface IQuickAccessNativeQuery
    {
        IReadOnlyList<string> GetItems(QuickAccess target, TimeSpan timeout);
    }

    internal interface IShellApplicationFactory
    {
        object CreateShellApplication();
    }

    internal sealed class DefaultShellApplicationFactory : IShellApplicationFactory
    {
        public object CreateShellApplication()
        {
            Type shellType = Type.GetTypeFromProgID("Shell.Application");
            if (shellType == null)
                throw new InvalidOperationException("Shell.Application COM object is not available.");

            return Activator.CreateInstance(shellType);
        }
    }

    internal sealed class ShellQuickAccessNativeQuery : IQuickAccessNativeQuery
    {
        private readonly INativeMethods _nativeMethods;
        private readonly IShellApplicationFactory _shellApplicationFactory;

        public ShellQuickAccessNativeQuery(INativeMethods nativeMethods)
            : this(nativeMethods, new DefaultShellApplicationFactory())
        {
        }

        internal ShellQuickAccessNativeQuery(INativeMethods nativeMethods, IShellApplicationFactory shellApplicationFactory)
        {
            _nativeMethods = nativeMethods ?? throw new ArgumentNullException(nameof(nativeMethods));
            _shellApplicationFactory = shellApplicationFactory ?? throw new ArgumentNullException(nameof(shellApplicationFactory));
        }

        public IReadOnlyList<string> GetItems(QuickAccess target, TimeSpan timeout)
        {
            var query = QuickAccessNativeQueryMapping.ForTarget(target);
            return StaThreadRunner.Run(() => EnumerateItemsOnSta(query), timeout, _nativeMethods);
        }

        private IReadOnlyList<string> EnumerateItemsOnSta(QuickAccessNativeQuerySpec query)
        {
            object shellApplication = null;
            object folder = null;
            object items = null;
            try
            {
                shellApplication = _shellApplicationFactory.CreateShellApplication();
                dynamic shellApp = shellApplication;
                folder = shellApp.Namespace(query.Namespace);
                if (folder == null)
                    throw new InvalidOperationException($"Failed to open shell namespace: {query.Namespace}");

                dynamic folderObj = folder;
                items = folderObj.Items();
                if (items == null)
                    throw new InvalidOperationException($"Failed to enumerate shell namespace: {query.Namespace}");

                dynamic itemsObj = items;
                int count = Convert.ToInt32(itemsObj.Count);
                var paths = new List<string>(Math.Max(0, count));

                for (int index = 0; index < count; index++)
                {
                    TryAddItemPath(itemsObj, index, query.Filter, paths);
                }

                return paths.AsReadOnly();
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

        private static void TryAddItemPath(dynamic items, int index, QuickAccessNativeQueryFilter filter, ICollection<string> paths)
        {
            object item = null;
            try
            {
                item = items.Item(index);
                if (item == null)
                    return;

                dynamic dItem = item;
                if (!QuickAccessNativeQueryMapping.ShouldKeep(dItem, filter))
                    return;

                string path = dItem.Path as string;
                if (!string.IsNullOrWhiteSpace(path))
                    paths.Add(path);
            }
            catch (COMException)
            {
                // A single corrupt or inaccessible shell item should not fail the whole query.
            }
            catch (Exception)
            {
                // Item-level failures are skipped to match the Rust native query behavior.
            }
            finally
            {
                if (item != null && Marshal.IsComObject(item))
                    Marshal.FinalReleaseComObject(item);
            }
        }
    }

    internal enum QuickAccessNativeQueryFilter
    {
        All,
        FilesOnly,
        FoldersOnly
    }

    internal sealed class QuickAccessNativeQuerySpec
    {
        public QuickAccessNativeQuerySpec(string @namespace, QuickAccessNativeQueryFilter filter)
        {
            Namespace = @namespace;
            Filter = filter;
        }

        public string Namespace { get; }

        public QuickAccessNativeQueryFilter Filter { get; }
    }

    internal static class QuickAccessNativeQueryMapping
    {
        public static QuickAccessNativeQuerySpec ForTarget(QuickAccess target)
        {
            switch (target)
            {
                case QuickAccess.All:
                    return new QuickAccessNativeQuerySpec(ShellNamespaces.QuickAccess, QuickAccessNativeQueryFilter.All);
                case QuickAccess.RecentFiles:
                    return new QuickAccessNativeQuerySpec(ShellNamespaces.QuickAccess, QuickAccessNativeQueryFilter.FilesOnly);
                case QuickAccess.FrequentFolders:
                    return new QuickAccessNativeQuerySpec(ShellNamespaces.FrequentFolders, QuickAccessNativeQueryFilter.FoldersOnly);
                default:
                    throw new ArgumentOutOfRangeException(nameof(target), target, "Unsupported Quick Access target.");
            }
        }

        public static bool ShouldKeep(dynamic item, QuickAccessNativeQueryFilter filter)
        {
            switch (filter)
            {
                case QuickAccessNativeQueryFilter.All:
                    return true;
                case QuickAccessNativeQueryFilter.FilesOnly:
                    return !Convert.ToBoolean(item.IsFolder);
                case QuickAccessNativeQueryFilter.FoldersOnly:
                    return Convert.ToBoolean(item.IsFolder);
                default:
                    throw new ArgumentOutOfRangeException(nameof(filter), filter, "Unsupported query filter.");
            }
        }
    }

    internal sealed class PowerShellFallbackNativeQuery : IQuickAccessNativeQuery
    {
        public IReadOnlyList<string> GetItems(QuickAccess target, TimeSpan timeout)
        {
            throw new InvalidOperationException("Native Quick Access query is disabled for this instance.");
        }
    }
}
