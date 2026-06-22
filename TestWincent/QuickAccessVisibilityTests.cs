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
        public void IsStartRecommendedSectionVisible_MissingValue_DefaultsToVisible()
        {
            var visibility = new RegistryQuickAccessVisibility(new StubRegistry());

            Assert.IsTrue(visibility.IsStartRecommendedSectionVisible());
        }

        [TestMethod]
        public void IsStartRecommendedSectionVisible_ZeroValue_IsHidden()
        {
            var registry = new StubRegistry();
            registry.AdvancedValues["Start_TrackDocs"] = 0;
            var visibility = new RegistryQuickAccessVisibility(registry);

            Assert.IsFalse(visibility.IsStartRecommendedSectionVisible());
        }

        [TestMethod]
        public void SetStartRecommendedSectionVisible_WritesStartTrackDocs()
        {
            var registry = new StubRegistry();
            var visibility = new RegistryQuickAccessVisibility(registry);

            visibility.SetStartRecommendedSectionVisible(true);

            Assert.AreEqual(1, registry.AdvancedValues["Start_TrackDocs"]);
            Assert.IsFalse(registry.Values.ContainsKey("Start_TrackDocs"));
        }

        [TestMethod]
        public void IsVisible_UnsupportedValueKind_ThrowsVisibilityException()
        {
            var registry = new StubRegistry();
            registry.Values["ShowRecent"] = new byte[] { 1 };
            registry.ValueKinds["ShowRecent"] = RegistryValueKind.Binary;
            var visibility = new RegistryQuickAccessVisibility(registry);

            var ex = Assert.ThrowsException<QuickAccessVisibilityException>(
                () => visibility.IsVisible(QuickAccess.RecentFiles));

            Assert.AreEqual("ReadVisibility", ex.Operation);
            Assert.AreEqual(QuickAccess.RecentFiles, ex.Target);
            Assert.AreEqual("ShowRecent", ex.ValueName);
            Assert.IsInstanceOfType(ex.InnerException, typeof(InvalidOperationException));
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
        public void IsVisible_RegistryReadFailure_ThrowsVisibilityException()
        {
            var registry = new StubRegistry
            {
                GetValueException = new UnauthorizedAccessException("read denied")
            };
            var visibility = new RegistryQuickAccessVisibility(registry);

            var ex = Assert.ThrowsException<QuickAccessVisibilityException>(
                () => visibility.IsVisible(QuickAccess.FrequentFolders));

            Assert.AreEqual("ReadVisibility", ex.Operation);
            Assert.AreEqual(QuickAccess.FrequentFolders, ex.Target);
            Assert.AreEqual("ShowFrequent", ex.ValueName);
            Assert.AreSame(registry.GetValueException, ex.InnerException);
        }

        [TestMethod]
        public void SetVisible_RegistryWriteFailure_ThrowsVisibilityException()
        {
            var registry = new StubRegistry
            {
                SetValueException = new UnauthorizedAccessException("write denied")
            };
            var visibility = new RegistryQuickAccessVisibility(registry);

            var ex = Assert.ThrowsException<QuickAccessVisibilityException>(
                () => visibility.SetVisible(QuickAccess.RecentFiles, true));

            Assert.AreEqual("WriteVisibility", ex.Operation);
            Assert.AreEqual(QuickAccess.RecentFiles, ex.Target);
            Assert.AreEqual("ShowRecent", ex.ValueName);
            Assert.AreSame(registry.SetValueException, ex.InnerException);
        }

        [TestMethod]
        public void IsStartRecommendedSectionVisible_RegistryReadFailure_ThrowsVisibilityException()
        {
            var registry = new StubRegistry
            {
                GetAdvancedValueException = new UnauthorizedAccessException("read denied")
            };
            var visibility = new RegistryQuickAccessVisibility(registry);

            var ex = Assert.ThrowsException<QuickAccessVisibilityException>(
                () => visibility.IsStartRecommendedSectionVisible());

            Assert.AreEqual("ReadStartRecommendedVisibility", ex.Operation);
            Assert.AreEqual(QuickAccess.RecentFiles, ex.Target);
            Assert.AreEqual("Start_TrackDocs", ex.ValueName);
            Assert.AreSame(registry.GetAdvancedValueException, ex.InnerException);
        }

        [TestMethod]
        public void SetStartRecommendedSectionVisible_RegistryWriteFailure_ThrowsVisibilityException()
        {
            var registry = new StubRegistry
            {
                SetAdvancedValueException = new UnauthorizedAccessException("write denied")
            };
            var visibility = new RegistryQuickAccessVisibility(registry);

            var ex = Assert.ThrowsException<QuickAccessVisibilityException>(
                () => visibility.SetStartRecommendedSectionVisible(false));

            Assert.AreEqual("WriteStartRecommendedVisibility", ex.Operation);
            Assert.AreEqual(QuickAccess.RecentFiles, ex.Target);
            Assert.AreEqual("Start_TrackDocs", ex.ValueName);
            Assert.AreSame(registry.SetAdvancedValueException, ex.InnerException);
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

            public Dictionary<string, object> AdvancedValues { get; } = new Dictionary<string, object>();

            public Dictionary<string, RegistryValueKind> AdvancedValueKinds { get; } = new Dictionary<string, RegistryValueKind>();

            public Exception GetValueException { get; set; }

            public Exception SetValueException { get; set; }

            public Exception GetAdvancedValueException { get; set; }

            public Exception SetAdvancedValueException { get; set; }

            public object GetValue(string name)
            {
                if (GetValueException != null)
                    throw GetValueException;

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
                if (SetValueException != null)
                    throw SetValueException;

                Values[name] = value;
            }

            public object GetAdvancedValue(string name)
            {
                if (GetAdvancedValueException != null)
                    throw GetAdvancedValueException;

                object value;
                return AdvancedValues.TryGetValue(name, out value) ? value : null;
            }

            public RegistryValueKind? GetAdvancedValueKind(string name)
            {
                RegistryValueKind kind;
                return AdvancedValueKinds.TryGetValue(name, out kind) ? (RegistryValueKind?)kind : null;
            }

            public void SetAdvancedDwordValue(string name, int value)
            {
                if (SetAdvancedValueException != null)
                    throw SetAdvancedValueException;

                AdvancedValues[name] = value;
            }
        }
    }
}
