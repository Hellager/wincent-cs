using System;
using System.Globalization;
using Microsoft.Win32;

namespace Wincent
{
    internal interface IQuickAccessVisibility
    {
        bool IsVisible(QuickAccess target);

        void SetVisible(QuickAccess target, bool visible);
    }

    internal interface IExplorerVisibilityRegistry
    {
        object GetValue(string name);

        RegistryValueKind? GetValueKind(string name);

        void SetDwordValue(string name, int value);
    }

    internal sealed class RegistryQuickAccessVisibility : IQuickAccessVisibility
    {
        // Explorer stores the Quick Access visibility toggles under the current user's Explorer key.
        // ShowFrequent controls automatically shown, unpinned frequent folders; pinned folders are separate.
        private const string ShowRecentValueName = "ShowRecent";
        private const string ShowFrequentValueName = "ShowFrequent";

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

        private bool IsValueVisible(string valueName, QuickAccess target)
        {
            try
            {
                object value = _registry.GetValue(valueName);
                if (value == null)
                    return true;

                RegistryValueKind? valueKind = _registry.GetValueKind(valueName);
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
                throw new QuickAccessVisibilityException("ReadVisibility", target, valueName, ex);
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
            try
            {
                _registry.SetDwordValue(valueName, visible ? 1 : 0);
            }
            catch (QuickAccessVisibilityException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new QuickAccessVisibilityException("WriteVisibility", target, valueName, ex);
            }
        }
    }

    internal sealed class CurrentUserExplorerVisibilityRegistry : IExplorerVisibilityRegistry
    {
        private const string ExplorerKeyPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer";

        public object GetValue(string name)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(ExplorerKeyPath, writable: false))
            {
                return key?.GetValue(name);
            }
        }

        public RegistryValueKind? GetValueKind(string name)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(ExplorerKeyPath, writable: false))
            {
                if (key == null || Array.IndexOf(key.GetValueNames(), name) < 0)
                    return null;

                return key.GetValueKind(name);
            }
        }

        public void SetDwordValue(string name, int value)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(ExplorerKeyPath))
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
    }
}
