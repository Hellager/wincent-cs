using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using Microsoft.Win32;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class QuickAccessVisibilityTests
    {
        [TestMethod]
        public void IsVisible_MissingValues_DefaultsToVisible()
        {
            var visibility = new RegistryQuickAccessVisibility(new StubRegistry());

            Assert.IsTrue(visibility.IsVisible(QuickAccess.RecentFiles));
            Assert.IsTrue(visibility.IsVisible(QuickAccess.FrequentFolders));
            Assert.IsTrue(visibility.IsVisible(QuickAccess.All));
        }

        [TestMethod]
        public void IsVisible_ZeroValue_IsHidden()
        {
            var registry = new StubRegistry();
            registry.Values["ShowRecent"] = 0;
            registry.Values["ShowFrequent"] = 0;
            var visibility = new RegistryQuickAccessVisibility(registry);

            Assert.IsFalse(visibility.IsVisible(QuickAccess.RecentFiles));
            Assert.IsFalse(visibility.IsVisible(QuickAccess.FrequentFolders));
            Assert.IsFalse(visibility.IsVisible(QuickAccess.All));
        }

        [TestMethod]
        public void IsVisible_NonZeroValue_IsVisible()
        {
            var registry = new StubRegistry();
            registry.Values["ShowRecent"] = 1;
            registry.ValueKinds["ShowRecent"] = RegistryValueKind.DWord;
            registry.Values["ShowFrequent"] = 2L;
            registry.ValueKinds["ShowFrequent"] = RegistryValueKind.QWord;
            var visibility = new RegistryQuickAccessVisibility(registry);

            Assert.IsTrue(visibility.IsVisible(QuickAccess.RecentFiles));
            Assert.IsTrue(visibility.IsVisible(QuickAccess.FrequentFolders));
            Assert.IsTrue(visibility.IsVisible(QuickAccess.All));
        }

        [TestMethod]
        public void IsVisible_UnsupportedValueKind_ThrowsInvalidOperationException()
        {
            var registry = new StubRegistry();
            registry.Values["ShowRecent"] = new byte[] { 1 };
            registry.ValueKinds["ShowRecent"] = RegistryValueKind.Binary;
            var visibility = new RegistryQuickAccessVisibility(registry);

            Assert.ThrowsException<InvalidOperationException>(() => visibility.IsVisible(QuickAccess.RecentFiles));
        }

        [TestMethod]
        public void IsVisible_AllRequiresBothSectionsVisible()
        {
            var registry = new StubRegistry();
            registry.Values["ShowRecent"] = 1;
            registry.Values["ShowFrequent"] = 0;
            var visibility = new RegistryQuickAccessVisibility(registry);

            Assert.IsFalse(visibility.IsVisible(QuickAccess.All));
        }

        [TestMethod]
        public void SetVisible_RecentFiles_WritesShowRecent()
        {
            var registry = new StubRegistry();
            var visibility = new RegistryQuickAccessVisibility(registry);

            visibility.SetVisible(QuickAccess.RecentFiles, false);

            Assert.AreEqual(0, registry.Values["ShowRecent"]);
            Assert.IsFalse(registry.Values.ContainsKey("ShowFrequent"));
        }

        [TestMethod]
        public void SetVisible_FrequentFolders_WritesShowFrequent()
        {
            var registry = new StubRegistry();
            var visibility = new RegistryQuickAccessVisibility(registry);

            visibility.SetVisible(QuickAccess.FrequentFolders, true);

            Assert.AreEqual(1, registry.Values["ShowFrequent"]);
            Assert.IsFalse(registry.Values.ContainsKey("ShowRecent"));
        }

        [TestMethod]
        public void SetVisible_All_WritesBothValues()
        {
            var registry = new StubRegistry();
            var visibility = new RegistryQuickAccessVisibility(registry);

            visibility.SetVisible(QuickAccess.All, false);

            Assert.AreEqual(0, registry.Values["ShowRecent"]);
            Assert.AreEqual(0, registry.Values["ShowFrequent"]);
        }

        [TestMethod]
        public void InvalidTarget_ThrowsArgumentOutOfRangeException()
        {
            var visibility = new RegistryQuickAccessVisibility(new StubRegistry());

            Assert.ThrowsException<ArgumentOutOfRangeException>(() => visibility.IsVisible((QuickAccess)999));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => visibility.SetVisible((QuickAccess)999, true));
        }

        private sealed class StubRegistry : IExplorerVisibilityRegistry
        {
            public Dictionary<string, object> Values { get; } = new Dictionary<string, object>();

            public Dictionary<string, RegistryValueKind> ValueKinds { get; } = new Dictionary<string, RegistryValueKind>();

            public object GetValue(string name)
            {
                object value;
                return Values.TryGetValue(name, out value) ? value : null;
            }

            public RegistryValueKind? GetValueKind(string name)
            {
                RegistryValueKind kind;
                return ValueKinds.TryGetValue(name, out kind) ? (RegistryValueKind?)kind : null;
            }

            public void SetDwordValue(string name, int value)
            {
                Values[name] = value;
            }
        }
    }
}
