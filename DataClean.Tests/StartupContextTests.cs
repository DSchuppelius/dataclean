using DataClean.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DataClean.Tests {
    [TestClass]
    public class StartupContextTests {
        [DataTestMethod]
        [DataRow("svchost")]
        [DataRow("SVCHOST")]
        [DataRow("taskeng")]
        [DataRow("taskhostw")]
        public void IsSchedulerProcess_SchedulerNames_ReturnTrue(string name) {
            Assert.IsTrue(StartupContext.IsSchedulerProcess(name));
        }

        [DataTestMethod]
        [DataRow("explorer")]
        [DataRow("cmd")]
        [DataRow("devenv")]
        [DataRow("DataClean")]
        [DataRow("")]
        [DataRow(null)]
        public void IsSchedulerProcess_OtherNames_ReturnFalse(string name) {
            Assert.IsFalse(StartupContext.IsSchedulerProcess(name));
        }

        [TestMethod]
        public void IsStartedByTaskScheduler_InTestRunner_ReturnsFalse() {
            Assert.IsFalse(StartupContext.IsStartedByTaskScheduler());
        }
    }
}
