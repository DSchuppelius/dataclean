using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DataClean.Services;

namespace DataClean.Tests {
    [TestClass]
    public class SpecialFolderResolverTests {
        [TestMethod]
        public void Resolve_UnknownToken_ReturnsInputUnchanged() {
            Assert.AreEqual(@"C:\Some\Custom\Path", SpecialFolderResolver.Resolve(@"C:\Some\Custom\Path"));
        }

        [TestMethod]
        public void Resolve_TempToken_ReturnsTempPath() {
            Assert.AreEqual(Path.GetTempPath(), SpecialFolderResolver.Resolve("%TEMP%"));
        }

        [TestMethod]
        public void Resolve_WindowsToken_ReturnsWindowsFolder() {
            string expected = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            Assert.AreEqual(expected, SpecialFolderResolver.Resolve("%WINDOWS%"));
        }

        [TestMethod]
        public void Resolve_PrefetchToken_EndsWithPrefetch() {
            string result = SpecialFolderResolver.Resolve("%PREFETCH%");
            StringAssert.EndsWith(result, "Prefetch");
        }

        [TestMethod]
        public void Resolve_EmptyString_ReturnsEmptyString() {
            Assert.AreEqual(string.Empty, SpecialFolderResolver.Resolve(string.Empty));
        }
    }
}
