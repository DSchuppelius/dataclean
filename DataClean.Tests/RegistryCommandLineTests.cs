using DataClean.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DataClean.Tests {
    [TestClass]
    public class RegistryCommandLineTests {
        [DataTestMethod]
        [DataRow("\"C:\\Program Files\\App\\app.exe\" /silent", "C:\\Program Files\\App\\app.exe")]
        [DataRow("\"C:\\App\\app.exe\"", "C:\\App\\app.exe")]
        [DataRow("C:\\App\\app.exe", "C:\\App\\app.exe")]
        [DataRow("C:\\Program Files\\My App\\app.exe /background", "C:\\Program Files\\My App\\app.exe")]
        [DataRow("C:\\App\\app.exe,0", "C:\\App\\app.exe")]
        [DataRow("%ProgramFiles%\\App\\app.exe -x", "%ProgramFiles%\\App\\app.exe")]
        [DataRow("C:\\App\\lib.dll\\2", "C:\\App\\lib.dll")]
        public void ExtractPath_FindsTheStartedFile(string commandLine, string expected) {
            Assert.AreEqual(expected, RegistryCommandLine.ExtractPath(commandLine));
        }

        [DataTestMethod]
        [DataRow(null)]
        [DataRow("")]
        [DataRow("   ")]
        [DataRow("MsiExec.exe /X{12345678-1234-1234-1234-123456789012}")]
        [DataRow("rundll32.exe C:\\App\\lib.dll,Start")]
        [DataRow("notepad.exe")]
        [DataRow("\\\\server\\share\\app.exe")]
        [DataRow("C:\\App\\ohne-endung")]
        [DataRow("\"C:\\App\\ohne-ende.exe")]
        public void ExtractPath_WithoutCertainty_ReturnsNull(string commandLine) {
            Assert.IsNull(RegistryCommandLine.ExtractPath(commandLine));
        }

        [TestMethod]
        public void ExtractPath_SecondProgramInArguments_ReturnsNull() {
            // Ohne Anführungszeichen könnte auch "C:\Program Files\foo.exe dir\bar.exe" gemeint sein.
            Assert.IsNull(RegistryCommandLine.ExtractPath("C:\\Program Files\\foo.exe dir\\bar.exe"));
        }

        [DataTestMethod]
        [DataRow("C:\\App\\app.exe.FriendlyAppName", "C:\\App\\app.exe")]
        [DataRow("C:\\App\\app.exe.ApplicationCompany", "C:\\App\\app.exe")]
        [DataRow("@C:\\App\\res.dll,-101", "C:\\App\\res.dll")]
        [DataRow("@%SystemRoot%\\system32\\shell32.dll,-8506", "%SystemRoot%\\system32\\shell32.dll")]
        public void ExtractMuiPath_FindsTheFile(string valueName, string expected) {
            Assert.AreEqual(expected, RegistryCommandLine.ExtractMuiPath(valueName));
        }

        [TestMethod]
        public void StripResourceIndex_RemovesOnlyTheNumber() {
            Assert.AreEqual("C:\\App\\lib.dll", RegistryCommandLine.StripResourceIndex("C:\\App\\lib.dll\\2"));
            Assert.AreEqual("C:\\App\\lib.dll", RegistryCommandLine.StripResourceIndex("C:\\App\\lib.dll"));
        }
    }
}
