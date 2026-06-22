using System;
using System.Runtime.InteropServices;
using System.Threading;

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
        private const string QuickAccessNamespace = "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}";
        private const string HomeNamespace = "shell:::{f874310e-b6b7-47dc-bc84-b9e6b38f5903}";
        private const int BrowserReadyStateComplete = 4;
        private const int BrowserReadyPollIntervalMilliseconds = 200;

        private readonly INativeMethods _nativeMethods;
        private readonly IShellApplicationFactory _shellApplicationFactory;
        private readonly Func<string> _desktopPathProvider;

        public ShellExplorerRefresher(INativeMethods nativeMethods)
            : this(nativeMethods, new DefaultShellApplicationFactory())
        {
        }

        internal ShellExplorerRefresher(
            INativeMethods nativeMethods,
            IShellApplicationFactory shellApplicationFactory)
            : this(
                  nativeMethods,
                  shellApplicationFactory,
                  () => Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory))
        {
        }

        internal ShellExplorerRefresher(
            INativeMethods nativeMethods,
            IShellApplicationFactory shellApplicationFactory,
            Func<string> desktopPathProvider)
        {
            _nativeMethods = nativeMethods ?? throw new ArgumentNullException(nameof(nativeMethods));
            _shellApplicationFactory = shellApplicationFactory ?? throw new ArgumentNullException(nameof(shellApplicationFactory));
            _desktopPathProvider = desktopPathProvider ?? throw new ArgumentNullException(nameof(desktopPathProvider));
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
                    {
                        NavigateRecentAccessWindowToDesktopAndBack(windows, count);
                        return;
                    }

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

        private void NavigateRecentAccessWindowToDesktopAndBack(object windows, int count)
        {
            object window = null;
            try
            {
                window = FindRecentAccessWindow(windows, count);
                if (window == null)
                    return;

                dynamic webBrowser = window;
                var original = ReadWindowLocation(webBrowser);
                NavigateToUrl(webBrowser, FileUrlFromPath(_desktopPathProvider()));
                NavigateBackToLocation(webBrowser, original);
            }
            catch (QuickAccessOperationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new QuickAccessOperationException("NavigateExplorerWindow", QuickAccess.All, null, ex);
            }
            finally
            {
                if (window != null && Marshal.IsComObject(window))
                    Marshal.FinalReleaseComObject(window);
            }
        }

        private static object FindRecentAccessWindow(object windows, int count)
        {
            for (int index = 0; index < count; index++)
            {
                object window = null;
                try
                {
                    window = ((dynamic)windows).Item(index);
                    if (window == null)
                        continue;

                    dynamic webBrowser = window;
                    if (IsRecentAccessLocation(ReadWindowLocation(webBrowser)))
                        return window;
                }
                catch (COMException)
                {
                }
                catch (Exception)
                {
                }

                if (window != null && Marshal.IsComObject(window))
                    Marshal.FinalReleaseComObject(window);
            }

            return null;
        }

        private static ExplorerLocation ReadWindowLocation(dynamic window)
        {
            return new ExplorerLocation(
                ReadStringProperty(() => window.LocationName),
                ReadStringProperty(() => window.LocationURL));
        }

        private static void NavigateBackToLocation(dynamic window, ExplorerLocation original)
        {
            if (!string.IsNullOrEmpty(original.LocationUrl))
            {
                NavigateToUrl(window, original.LocationUrl);
                return;
            }

            foreach (var candidate in new[] { QuickAccessNamespace, HomeNamespace })
            {
                NavigateToUrl(window, candidate);
                if (IsRecentAccessLocation(ReadWindowLocation(window)))
                    return;
            }

            throw new InvalidOperationException("Failed to navigate Explorer back to Quick Access or Home.");
        }

        private static void NavigateToUrl(dynamic window, string url)
        {
            window.Navigate2(url, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            WaitForBrowserReady(window, TimeSpan.FromSeconds(5));
        }

        private static void WaitForBrowserReady(dynamic window, TimeSpan timeout)
        {
            var deadline = DateTime.UtcNow + timeout;
            while (DateTime.UtcNow < deadline)
            {
                if (!ReadBrowserBusy(window) && ReadBrowserReadyState(window) == BrowserReadyStateComplete)
                    return;

                Thread.Sleep(BrowserReadyPollIntervalMilliseconds);
            }
        }

        private static bool ReadBrowserBusy(dynamic window)
        {
            try
            {
                return Convert.ToBoolean(window.Busy);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static int ReadBrowserReadyState(dynamic window)
        {
            try
            {
                return Convert.ToInt32(window.ReadyState);
            }
            catch (Exception)
            {
                return BrowserReadyStateComplete;
            }
        }

        internal static string FileUrlFromPath(string path)
        {
            string value = (path ?? string.Empty).Replace('\\', '/');
            if (value.Length >= 2 && value[1] == ':')
                value = "/" + value;

            return "file://" + value;
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
