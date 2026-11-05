using DataClean.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DataClean.Tests {
    [TestClass]
    public class ThemeModeTests {
        [DataTestMethod]
        [DataRow("System")]
        [DataRow("system")]
        [DataRow("")]
        [DataRow(null)]
        [DataRow("Unbekannt")]
        public void IsDark_WithoutFixedChoice_FollowsWindows(string themeOverride) {
            Assert.IsTrue(ThemeMode.IsDark(themeOverride, systemIsDark: true));
            Assert.IsFalse(ThemeMode.IsDark(themeOverride, systemIsDark: false));
        }

        [DataTestMethod]
        [DataRow("Light")]
        [DataRow("LIGHT")]
        public void IsDark_Light_IgnoresWindows(string themeOverride) {
            Assert.IsFalse(ThemeMode.IsDark(themeOverride, systemIsDark: true));
            Assert.IsFalse(ThemeMode.IsDark(themeOverride, systemIsDark: false));
        }

        [DataTestMethod]
        [DataRow("Dark")]
        [DataRow("dark")]
        public void IsDark_Dark_IgnoresWindows(string themeOverride) {
            Assert.IsTrue(ThemeMode.IsDark(themeOverride, systemIsDark: true));
            Assert.IsTrue(ThemeMode.IsDark(themeOverride, systemIsDark: false));
        }
    }
}
