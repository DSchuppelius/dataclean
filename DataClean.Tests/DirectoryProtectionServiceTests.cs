using Microsoft.VisualStudio.TestTools.UnitTesting;
using DataClean.Services;

namespace DataClean.Tests {
    [TestClass]
    public class DirectoryProtectionServiceTests {
        private static readonly string[] Protected = { @"C:\Windows", @"C:\Program Files" };
        private static readonly string[] Allowed = { @"C:\Windows\Temp" };

        [TestMethod]
        public void IsProtected_DirectoryUnderProtectedRoot_ReturnsTrue() {
            Assert.IsTrue(DirectoryProtectionService.IsProtected(@"C:\Windows\System32", Protected, Allowed));
        }

        [TestMethod]
        public void IsProtected_AllowedSubdirectory_ReturnsFalse() {
            Assert.IsFalse(DirectoryProtectionService.IsProtected(@"C:\Windows\Temp\abc", Protected, Allowed));
        }

        [TestMethod]
        public void IsProtected_UnrelatedDirectory_ReturnsFalse() {
            Assert.IsFalse(DirectoryProtectionService.IsProtected(@"D:\Data\Stuff", Protected, Allowed));
        }

        [TestMethod]
        public void IsProtected_CaseInsensitiveMatch_ReturnsTrue() {
            Assert.IsTrue(DirectoryProtectionService.IsProtected(@"c:\windows\system32", Protected, Allowed));
        }

        [TestMethod]
        public void IsProtected_EmptyDirectory_ReturnsFalse() {
            Assert.IsFalse(DirectoryProtectionService.IsProtected("", Protected, Allowed));
        }

        [TestMethod]
        public void IsProtected_NullLists_ReturnsFalse() {
            Assert.IsFalse(DirectoryProtectionService.IsProtected(@"C:\Windows", null, null));
        }
    }
}
