using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DataClean.Services;

namespace DataClean.Tests {
    [TestClass]
    public class FileAgeFilterTests {
        private static readonly DateTime Now = new DateTime(2024, 1, 31, 12, 0, 0);

        [TestMethod]
        public void ShouldSkip_WhenMinAgeDisabled_ReturnsFalse() {
            Assert.IsFalse(FileAgeFilter.ShouldSkip(Now, Now, 0, Now));
        }

        [TestMethod]
        public void ShouldSkip_WhenFileIsYoungerThanThreshold_ReturnsTrue() {
            // Beide Zeitstempel neuer als (Now - 7 Tage) => überspringen.
            var recent = Now.AddDays(-2);
            Assert.IsTrue(FileAgeFilter.ShouldSkip(recent, recent, 7, Now));
        }

        [TestMethod]
        public void ShouldSkip_WhenFileIsOlderThanThreshold_ReturnsFalse() {
            var old = Now.AddDays(-30);
            Assert.IsFalse(FileAgeFilter.ShouldSkip(old, old, 7, Now));
        }

        [TestMethod]
        public void ShouldSkip_WhenOnlyAccessTimeIsRecent_ReturnsFalse() {
            // Schreibzeit alt, Zugriffszeit neu => nicht überspringen (beide müssen jung sein).
            Assert.IsFalse(FileAgeFilter.ShouldSkip(Now.AddDays(-30), Now.AddDays(-1), 7, Now));
        }

        [TestMethod]
        public void ShouldSkip_WhenExactlyAtThreshold_ReturnsFalse() {
            var threshold = Now.AddDays(-7);
            Assert.IsFalse(FileAgeFilter.ShouldSkip(threshold, threshold, 7, Now));
        }
    }
}
