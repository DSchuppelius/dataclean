using Microsoft.VisualStudio.TestTools.UnitTesting;
using DataClean;
using DataClean.Services;

namespace DataClean.Tests {
    [TestClass]
    public class CleanupProfileTests {
        private static CleanupOptions SampleOptions() {
            return new CleanupOptions {
                DryRun = true,
                CreateRestorePoint = true,
                DeleteToRecycleBin = true,
                CleanBrowserCaches = true,
                DismComponentCleanup = true,
                CleanMinidumps = true,
                CleanPrefetch = true,
                CleanDeliveryOptimization = true,
                CleanOldLogs = true,
                OldLogAgeDays = 42,
                EmptyAllRecycleBins = true,
                ConfirmBeforeDelete = false,
                GenerateReport = true,
                ReportPath = @"C:\Reports",
                ReportFormat = "CSV",
                EmailReport = true,
                ExcludePatterns = "*.keep",
                SkipIfEnoughFreeSpace = true,
                SchedulePendingDeletes = true
            };
        }

        [TestMethod]
        public void FromOptions_SetsName() {
            var profile = CleanupProfile.FromOptions("Mein Profil", SampleOptions());
            Assert.AreEqual("Mein Profil", profile.Name);
        }

        [TestMethod]
        public void RoundTrip_PreservesAllValues() {
            var original = SampleOptions();
            var profile = CleanupProfile.FromOptions("P", original);

            var target = new CleanupOptions();
            profile.ApplyTo(target);

            Assert.AreEqual(original.DryRun, target.DryRun);
            Assert.AreEqual(original.CreateRestorePoint, target.CreateRestorePoint);
            Assert.AreEqual(original.DeleteToRecycleBin, target.DeleteToRecycleBin);
            Assert.AreEqual(original.CleanBrowserCaches, target.CleanBrowserCaches);
            Assert.AreEqual(original.DismComponentCleanup, target.DismComponentCleanup);
            Assert.AreEqual(original.CleanMinidumps, target.CleanMinidumps);
            Assert.AreEqual(original.CleanPrefetch, target.CleanPrefetch);
            Assert.AreEqual(original.CleanDeliveryOptimization, target.CleanDeliveryOptimization);
            Assert.AreEqual(original.CleanOldLogs, target.CleanOldLogs);
            Assert.AreEqual(original.OldLogAgeDays, target.OldLogAgeDays);
            Assert.AreEqual(original.EmptyAllRecycleBins, target.EmptyAllRecycleBins);
            Assert.AreEqual(original.ConfirmBeforeDelete, target.ConfirmBeforeDelete);
            Assert.AreEqual(original.GenerateReport, target.GenerateReport);
            Assert.AreEqual(original.ReportPath, target.ReportPath);
            Assert.AreEqual(original.ReportFormat, target.ReportFormat);
            Assert.AreEqual(original.EmailReport, target.EmailReport);
            Assert.AreEqual(original.ExcludePatterns, target.ExcludePatterns);
            Assert.AreEqual(original.SkipIfEnoughFreeSpace, target.SkipIfEnoughFreeSpace);
            Assert.AreEqual(original.SchedulePendingDeletes, target.SchedulePendingDeletes);
        }
    }
}
