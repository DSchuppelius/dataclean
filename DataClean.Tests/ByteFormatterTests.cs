using Microsoft.VisualStudio.TestTools.UnitTesting;
using DataClean.Services;

namespace DataClean.Tests {
    [TestClass]
    public class ByteFormatterTests {
        [TestMethod]
        public void Format_Zero_ReturnsBytes() {
            Assert.AreEqual("0 B", ByteFormatter.Format(0));
        }

        [TestMethod]
        public void Format_BelowKilobyte_ReturnsBytes() {
            Assert.AreEqual("512 B", ByteFormatter.Format(512));
        }

        [TestMethod]
        public void Format_ExactlyOneKilobyte_ReturnsKb() {
            Assert.AreEqual("1 KB", ByteFormatter.Format(1024));
        }

        [TestMethod]
        public void Format_OneAndHalfMegabyte_ReturnsRoundedMb() {
            // Nachkommatrenner ist kulturabhängig – Erwartung entsprechend aufbauen.
            string expected = (1.5).ToString("0.##") + " MB";
            Assert.AreEqual(expected, ByteFormatter.Format(1024L * 1024 + 512 * 1024));
        }

        [TestMethod]
        public void Format_Gigabyte_ReturnsGb() {
            Assert.AreEqual("2 GB", ByteFormatter.Format(2L * 1024 * 1024 * 1024));
        }

        [TestMethod]
        public void Format_Terabyte_ReturnsTb() {
            Assert.AreEqual("1 TB", ByteFormatter.Format(1024L * 1024 * 1024 * 1024));
        }
    }
}
