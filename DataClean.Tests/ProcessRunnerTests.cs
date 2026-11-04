using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using DataClean.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DataClean.Tests {
    [TestClass]
    public class ProcessRunnerTests {
        private static readonly TimeSpan Generous = TimeSpan.FromSeconds(60);

        [TestMethod]
        public void Run_MoreOutputThanThePipeBuffer_DoesNotBlock() {
            // Rund 200 KB: Ein ungelesener Puffer (4 KB) würde das Programm hier anhalten.
            int lines = 0;
            ProcessResult result = ProcessRunner.Run("cmd.exe",
                "/c for /L %i in (1,1,4000) do @echo Zeile %i mit etwas Text, damit der Puffer sicher vollaeuft",
                line => lines++, Generous);

            Assert.IsFalse(result.TimedOut);
            Assert.AreEqual(0, result.ExitCode);
            Assert.AreEqual(4000, lines);
        }

        [TestMethod]
        public void Run_ReturnsExitCodeAndErrorOutput() {
            ProcessResult result = ProcessRunner.Run("cmd.exe", "/c echo kaputt 1>&2 & exit /b 3", null, Generous);

            Assert.AreEqual(3, result.ExitCode);
            StringAssert.Contains(result.ErrorOutput, "kaputt");
        }

        [TestMethod]
        public void Run_StillRunningAfterTimeout_ReportsTimedOutAndReturns() {
            var clock = Stopwatch.StartNew();
            ProcessResult result = ProcessRunner.Run("ping.exe", "-n 4 127.0.0.1", null, TimeSpan.FromMilliseconds(300));

            Assert.IsTrue(result.TimedOut);
            Assert.IsTrue(clock.Elapsed < TimeSpan.FromSeconds(2), "Dauer: " + clock.Elapsed);
        }

        [TestMethod]
        public void Run_MissingProgram_Throws() {
            try {
                ProcessRunner.Run("gibt-es-nicht-4711.exe", string.Empty, null, Generous);
                Assert.Fail("Win32Exception erwartet");
            } catch (Win32Exception) {
                // Erwartet: Der Aufrufer protokolliert den Startfehler.
            }
        }

        [TestMethod]
        public void ReadLines_DeliversProgressBarWithoutLineEnd() {
            var lines = new List<string>();
            ProcessRunner.ReadLines(new StringReader("[= 10.0% =]\r[== 20.0% ==]"), lines.Add);
            CollectionAssert.AreEqual(new[] { "[= 10.0% =]", "[== 20.0% ==]" }, lines);
        }

        [TestMethod]
        public void ReadLines_SplitsAtCarriageReturnAndNewline() {
            var lines = new List<string>();
            ProcessRunner.ReadLines(new StringReader("10%\r20%\r\n\r\nfertig"), lines.Add);
            CollectionAssert.AreEqual(new[] { "10%", "20%", "fertig" }, lines);
        }
    }

    [TestClass]
    public class DismProgressParserTests {
        [DataTestMethod]
        [DataRow("[==========================100.0%==========================]", 100d)]
        [DataRow("[=====                      10.0%                          ]", 10d)]
        [DataRow("[===   5,7% ]", 5.7d)]
        [DataRow("[                           0.0%                           ]", 0d)]
        public void TryParsePercent_ProgressBar_ReturnsValue(string line, double expected) {
            double percent;
            Assert.IsTrue(DismProgressParser.TryParsePercent(line, out percent));
            Assert.AreEqual(expected, percent, 0.001);
        }

        [DataTestMethod]
        [DataRow("Tool zur Imageverwaltung für die Bereitstellung")]
        [DataRow("Version: 10.0.26100.1")]
        [DataRow("Speicher: 42%")]
        [DataRow("[==== 250.0% ====]")]
        [DataRow("")]
        [DataRow(null)]
        public void TryParsePercent_OtherText_ReturnsFalse(string line) {
            double percent;
            Assert.IsFalse(DismProgressParser.TryParsePercent(line, out percent));
        }
    }
}
