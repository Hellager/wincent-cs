using System;
using System.Globalization;
using Microsoft.Win32;

namespace Wincent
{
    internal interface IQuickAccessVisibility
    {
        bool IsVisible(QuickAccess target);

        void SetVisible(QuickAccess target, bool visible);

        bool IsStartRecommendedSectionVisible();

        void SetStartRecommendedSectionVisible(bool visible);
    }

    internal interface IExplorerVisibilityRegistry
    {
        object GetValue(string name);

        RegistryValueKind? GetValueKind(string name);

        void SetDwordValue(string name, int value);

        object GetAdvancedValue(string name);

        RegistryValueKind? GetAdvancedValueKind(string name);

        void SetAdvancedDwordValue(string name, int value);
    }

    /// <summary>
    /// Registry-backed visibility controls for Explorer Quick Access and Start Recommended items.
    /// </summary>
    /// <remarks>
    /// This implementation writes only the current user's registry values. It does not open Explorer's Folder Options UI
    /// and does not deliberately clear Quick Access history. Observed Windows behavior differs between direct registry
    /// writes and UI changes: hiding Frequent Folders here hides unpinned frequent folders while pinned folders remain
    /// visible, but changing the same option in the Explorer UI can clear unpinned frequent folders and later restore
    /// default system pins. Explorer UI changes to Recent Files visibility can also clear recent file entries.
    /// </remarks>
    internal sealed class RegistryQuickAccessVisibility : IQuickAccessVisibility
    {
        // Explorer stores the Quick Access visibility toggles under the current user's Explorer key.
        // ShowFrequent controls automatically shown, unpinned frequent folders; pinned folders are separate. Start
        // Recommended recent-file visibility is controlled by Start_TrackDocs under Explorer\Advanced.
        private const string ShowRecentValueName = "ShowRecent";
        private const string ShowFrequentValueName = "ShowFrequent";
        private const string StartTrackDocsValueName = "Start_TrackDocs";

        private readonly IExplorerVisibilityRegistry _registry;

        public RegistryQuickAccessVisibility(IExplorerVisibilityRegistry registry)
        {
            _registry = registry ?? throw new ArgumentNullException(nameof(registry));
        }

        public bool IsVisible(QuickAccess target)
        {
            switch (target)
            {
                case QuickAccess.All:
                    return IsValueVisible(ShowRecentValueName, target) && IsValueVisible(ShowFrequentValueName, target);
                case QuickAccess.RecentFiles:
                    return IsValueVisible(ShowRecentValueName, target);
                case QuickAccess.FrequentFolders:
                    return IsValueVisible(ShowFrequentValueName, target);
                default:
                    throw new ArgumentOutOfRangeException(nameof(target), target, "Unsupported Quick Access target.");
            }
        }

        public void SetVisible(QuickAccess target, bool visible)
        {
            switch (target)
            {
                case QuickAccess.All:
                    SetValueVisible(ShowRecentValueName, visible, target);
                    SetValueVisible(ShowFrequentValueName, visible, target);
                    break;
                case QuickAccess.RecentFiles:
                    SetValueVisible(ShowRecentValueName, visible, target);
                    break;
                case QuickAccess.FrequentFolders:
                    SetValueVisible(ShowFrequentValueName, visible, target);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(target), target, "Unsupported Quick Access target.");
            }
        }

        public bool IsStartRecommendedSectionVisible()
        {
            return IsValueVisible(
                StartTrackDocsValueName,
                QuickAccess.RecentFiles,
                _registry.GetAdvancedValue,
                _registry.GetAdvancedValueKind,
                "ReadStartRecommendedVisibility");
        }

        public void SetStartRecommendedSectionVisible(bool visible)
        {
            SetValueVisible(
                StartTrackDocsValueName,
                visible,
                QuickAccess.RecentFiles,
                _registry.SetAdvancedDwordValue,
                "WriteStartRecommendedVisibility");
        }

        private bool IsValueVisible(string valueName, QuickAccess target)
        {
            return IsValueVisible(valueName, target, _registry.GetValue, _registry.GetValueKind, "ReadVisibility");
        }

        private bool IsValueVisible(
            string valueName,
            QuickAccess target,
            Func<string, object> getValue,
            Func<string, RegistryValueKind?> getValueKind,
            string operation)
        {
            try
            {
                object value = getValue(valueName);
                if (value == null)
                    return true;

                RegistryValueKind? valueKind = getValueKind(valueName);
                switch (valueKind ?? InferValueKind(value))
                {
                    case RegistryValueKind.DWord:
                        return Convert.ToInt32(value, CultureInfo.InvariantCulture) != 0;
                    case RegistryValueKind.QWord:
                        return Convert.ToInt64(value, CultureInfo.InvariantCulture) != 0;
                    case RegistryValueKind.String:
                    case RegistryValueKind.ExpandString:
                        int parsed;
                        if (int.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed))
                            return parsed != 0;

                        throw new InvalidOperationException("Explorer visibility registry value has an unsupported string payload.");
                    default:
                        throw new InvalidOperationException("Explorer visibility registry value has an unsupported type.");
                }
            }
            catch (QuickAccessVisibilityException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new QuickAccessVisibilityException(operation, target, valueName, ex);
            }
        }

        private static RegistryValueKind InferValueKind(object value)
        {
            if (value is byte || value is short || value is int || value is long ||
                value is sbyte || value is ushort || value is uint || value is ulong)
            {
                return value is long || value is ulong ? RegistryValueKind.QWord : RegistryValueKind.DWord;
            }

            if (value is string)
                return RegistryValueKind.String;

            throw new InvalidOperationException("Explorer visibility registry value has an unsupported type.");
        }

        private void SetValueVisible(string valueName, bool visible, QuickAccess target)
        {
            SetValueVisible(valueName, visible, target, _registry.SetDwordValue, "WriteVisibility");
        }

        private void SetValueVisible(
            string valueName,
            bool visible,
            QuickAccess target,
            Action<string, int> setDwordValue,
            string operation)
        {
            try
            {
                setDwordValue(valueName, visible ? 1 : 0);
            }
            catch (QuickAccessVisibilityException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new QuickAccessVisibilityException(operation, target, valueName, ex);
            }
        }
    }

    internal sealed class CurrentUserExplorerVisibilityRegistry : IExplorerVisibilityRegistry
    {
        private const string ExplorerKeyPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer";
        private const string ExplorerAdvancedKeyPath = ExplorerKeyPath + @"\Advanced";

        public object GetValue(string name)
        {
            return GetValue(ExplorerKeyPath, name);
        }

        public RegistryValueKind? GetValueKind(string name)
        {
            return GetValueKind(ExplorerKeyPath, name);
        }

        public void SetDwordValue(string name, int value)
        {
            SetDwordValue(ExplorerKeyPath, name, value);
        }

        public object GetAdvancedValue(string name)
        {
            return GetValue(ExplorerAdvancedKeyPath, name);
        }

        public RegistryValueKind? GetAdvancedValueKind(string name)
        {
            return GetValueKind(ExplorerAdvancedKeyPath, name);
        }

        public void SetAdvancedDwordValue(string name, int value)
        {
            SetDwordValue(ExplorerAdvancedKeyPath, name, value);
        }

        private static object GetValue(string keyPath, string name)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath, writable: false))
            {
                return key?.GetValue(name);
            }
        }

        private static RegistryValueKind? GetValueKind(string keyPath, string name)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath, writable: false))
            {
                if (key == null || Array.IndexOf(key.GetValueNames(), name) < 0)
                    return null;

                return key.GetValueKind(name);
            }
        }

        private static void SetDwordValue(string keyPath, string name, int value)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(keyPath))
            {
                if (key == null)
                    throw new InvalidOperationException("Unable to open Explorer registry key.");

                key.SetValue(name, value, RegistryValueKind.DWord);
            }
        }
    }

    internal sealed class NoOpQuickAccessVisibility : IQuickAccessVisibility
    {
        public bool IsVisible(QuickAccess target)
        {
            return true;
        }

        public void SetVisible(QuickAccess target, bool visible)
        {
        }

        public bool IsStartRecommendedSectionVisible()
        {
            return true;
        }

        public void SetStartRecommendedSectionVisible(bool visible)
        {
        }
    }
}
