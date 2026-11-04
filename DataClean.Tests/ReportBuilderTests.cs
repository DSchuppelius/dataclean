using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DataClean.Services;

namespace DataClean.Tests {
    [TestClass]
    public class ReportBuilderTests {
        private static ReportData SampleData() {
            return new ReportData {
                RunStartTime = new DateTime(2024, 1, 31, 8, 15, 30),
                MachineName = "TEST-PC",
                DryRun = false,
                DeletedItemCount = 42,
                FreedBytes = 1024L * 1024,
                SkippedItemCount = 3,
                FailedItemCount = 1,
                Duration = TimeSpan.FromSeconds(75),
                Tasks = "Browser-Caches, Prefetch"
            };
        }

        [TestMethod]
        public void BuildFileName_Html_HasHtmlExtensionAndTimestamp() {
            string name = ReportBuilder.BuildFileName(new DateTime(2024, 1, 31, 8, 15, 30), isHtml: true);
            Assert.AreEqual("DataClean-Report_20240131_081530.html", name);
        }

        [TestMethod]
        public void BuildFileName_Csv_HasCsvExtension() {
            string name = ReportBuilder.BuildFileName(new DateTime(2024, 1, 31, 8, 15, 30), isHtml: false);
            Assert.AreEqual("DataClean-Report_20240131_081530.csv", name);
        }

        [TestMethod]
        public void BuildHtml_ContainsKeyFiguresAndEncodesValues() {
            string html = ReportBuilder.BuildHtml(SampleData());
            StringAssert.Contains(html, "<h1>DataClean Report</h1>");
            StringAssert.Contains(html, "TEST-PC");
            StringAssert.Contains(html, "42");
            StringAssert.Contains(html, "Löschen");
            StringAssert.Contains(html, "00:01:15");
        }

        [TestMethod]
        public void BuildHtml_DryRun_ShowsPreviewMode() {
            var data = SampleData();
            data.DryRun = true;
            string html = ReportBuilder.BuildHtml(data);
            StringAssert.Contains(html, "Vorschau (Dry-Run)");
        }

        [TestMethod]
        public void BuildCsv_HasHeaderAndRows() {
            string csv = ReportBuilder.BuildCsv(SampleData());
            StringAssert.Contains(csv, "Feld;Wert");
            StringAssert.Contains(csv, "Computer;TEST-PC");
            StringAssert.Contains(csv, "Gelöschte Elemente;42");
            StringAssert.Contains(csv, "Dauer;00:01:15");
        }
    }
}
