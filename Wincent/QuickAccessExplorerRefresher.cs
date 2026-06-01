using System;
using System.Runtime.InteropServices;

namespace Wincent
{
    internal interface IExplorerRefresher
    {
        void Refresh(TimeSpan timeout);
    }

    internal sealed class ShellExplorerRefresher : IExplorerRefresher
    {
        private const string QuickAccessGuid = "679f85cb-0220-4080-b29b-5540cc05aab6";
        private const string HomeGuid = "f874310e-b6b7-47dc-bc84-b9e6b38f5903";

        private readonly INativeMethods _nativeMethods;
        private readonly IShellApplicationFactory _shellApplicationFactory;

        public ShellExplorerRefresher(INativeMethods nativeMethods)
            : this(nativeMethods, new DefaultShellApplicationFactory())
        {
        }

        internal ShellExplorerRefresher(
            INativeMethods nativeMethods,
            IShellApplicationFactory shellApplicationFactory)
        {
            _nativeMethods = nativeMethods ?? throw new ArgumentNullException(nameof(nativeMethods));
            _shellApplicationFactory = shellApplicationFactory ?? throw new ArgumentNullException(nameof(shellApplicationFactory));
        }

        public void Refresh(TimeSpan timeout)
        {
            StaThreadRunner.Run(RefreshOnStaThread, timeout, _nativeMethods);
        }

        private void RefreshOnStaThread()
        {
            object shellApplication = null;
            object windows = null;
            try
            {
                shellApplication = _shellApplicationFactory.CreateShellApplication();
                dynamic shellApp = shellApplication;
                windows = shellApp.Windows();
                if (windows == null)
                    return;

                dynamic windowsObj = windows;
                int count = Convert.ToInt32(windowsObj.Count);

                var recentResult = RefreshMatchingWindows(windows, count, IsRecentAccessLocation);
                if (recentResult.Matched > 0)
                {
                    if (recentResult.Refreshed > 0)
                        return;

                    throw new QuickAccessOperationException(
                        "RefreshExplorer",
                        QuickAccess.All,
                        null,
                        new InvalidOperationException($"Found {recentResult.Matched} Quick Access or Home Explorer windows but refreshed none."));
                }

                var explorerResult = RefreshMatchingWindows(windows, count, IsProbableExplorerLocation);
                if (explorerResult.Matched > 0 && explorerResult.Refreshed == 0)
                {
                    throw new QuickAccessOperationException(
                        "RefreshExplorer",
                        QuickAccess.All,
                        null,
                        new InvalidOperationException($"Found {explorerResult.Matched} Explorer windows but refreshed none."));
                }
            }
            finally
            {
                if (windows != null && Marshal.IsComObject(windows))
                    Marshal.FinalReleaseComObject(windows);
                if (shellApplication != null && Marshal.IsComObject(shellApplication))
                    Marshal.FinalReleaseComObject(shellApplication);
            }
        }

        private static ExplorerRefreshResult RefreshMatchingWindows(
            object windows,
            int count,
            Func<ExplorerLocation, bool> predicate)
        {
            var result = new ExplorerRefreshResult();

            for (int index = 0; index < count; index++)
            {
                object window = null;
                try
                {
                    window = ((dynamic)windows).Item(index);
                    if (window == null)
                        continue;

                    dynamic w = window;
                    var location = new ExplorerLocation(
                        ReadStringProperty(() => w.LocationName),
                        ReadStringProperty(() => w.LocationURL));

                    if (!predicate(location))
                        continue;

                    result.Matched++;
                    if (TryRefreshWindow(w))
                        result.Refreshed++;
                }
                catch (COMException)
                {
                    continue;
                }
                catch (Exception)
                {
                    continue;
                }
                finally
                {
                    if (window != null && Marshal.IsComObject(window))
                        Marshal.FinalReleaseComObject(window);
                }
            }

            return result;
        }

        private static bool TryRefreshWindow(dynamic window)
        {
            try
            {
                window.Refresh();
                return true;
            }
            catch (Exception)
            {
            }

            object document = null;
            try
            {
                document = window.Document;
                dynamic doc = document;
                doc.Refresh();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (document != null && Marshal.IsComObject(document))
                    Marshal.FinalReleaseComObject(document);
            }
        }

        private static string ReadStringProperty(Func<object> read)
        {
            try
            {
                return Convert.ToString(read()) ?? string.Empty;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        internal static bool IsRecentAccessLocation(ExplorerLocation location)
        {
            if (location == null)
                return false;

            var url = location.LocationUrl.ToLowerInvariant();
            return url.Contains(QuickAccessGuid) || url.Contains(HomeGuid);
        }

        internal static bool IsProbableExplorerLocation(ExplorerLocation location)
        {
            if (location == null)
                return false;

            var url = location.LocationUrl.ToLowerInvariant();
            return url.StartsWith("file:///", StringComparison.OrdinalIgnoreCase) ||
                   url.StartsWith("shell:::", StringComparison.OrdinalIgnoreCase) ||
                   url.StartsWith("::{", StringComparison.OrdinalIgnoreCase) ||
                   url.StartsWith("ms-shell", StringComparison.OrdinalIgnoreCase) ||
                   string.IsNullOrEmpty(location.LocationUrl) && !string.IsNullOrEmpty(location.LocationName);
        }

        private sealed class ExplorerRefreshResult
        {
            public int Matched { get; set; }

            public int Refreshed { get; set; }
        }
    }

    internal sealed class ExplorerLocation
    {
        public ExplorerLocation(string locationName, string locationUrl)
        {
            LocationName = locationName ?? string.Empty;
            LocationUrl = locationUrl ?? string.Empty;
        }

        public string LocationName { get; }

        public string LocationUrl { get; }
    }

    internal sealed class PowerShellFallbackExplorerRefresher : IExplorerRefresher
    {
        public void Refresh(TimeSpan timeout)
        {
            throw new InvalidOperationException("Native Explorer refresh is disabled for this instance.");
        }
    }
}
