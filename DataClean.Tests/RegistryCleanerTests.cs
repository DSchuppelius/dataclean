using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using DataClean.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Win32;

namespace DataClean.Tests {
    /// <summary>
    /// Prüft Suche, Sicherung und Entfernen an einem eigenen Schlüssel unter HKCU. Die echten
    /// Einträge werden über <see cref="RegistryScope.Sandbox"/> nie berührt.
    /// </summary>
    [TestClass]
    public class RegistryCleanerTests {
        private const string RunKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        private const string UninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
        private const string SharedDllsKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDLLs";
        private const string ClsidKey = @"SOFTWARE\Classes\CLSID";

        private string sandboxKey;
        private string folder;
        private string existing;
        private string missing;
        private RegistryScope scope;
        private RegistryScanner scanner;
        private RegistryCleaner cleaner;

        [TestInitialize]
        public void SetUp() {
            sandboxKey = @"Software\DataClean.Tests\" + Guid.NewGuid().ToString("N");
            folder = Path.Combine(Path.GetTempPath(), "DataClean.Tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            existing = Path.Combine(folder, "vorhanden.exe");
            File.WriteAllText(existing, string.Empty);
            missing = Path.Combine(folder, "fehlt.exe");

            scope = RegistryScope.Sandbox(sandboxKey);
            var probe = new RegistryPathProbe();
            scanner = new RegistryScanner(scope, probe);
            cleaner = new RegistryCleaner(scope, probe, Path.Combine(folder, "Sicherungen"));
        }

        [TestCleanup]
        public void TearDown() {
            Registry.CurrentUser.DeleteSubKeyTree(sandboxKey, false);
            Directory.Delete(folder, true);
        }

        private RegistryKey Create(string keyPath) {
            return Registry.CurrentUser.CreateSubKey(sandboxKey + @"\" + RegistryHive.LocalMachine + @"\" + keyPath);
        }

        private List<RegistryIssue> Scan(params RegistryCategory[] categories) {
            return scanner.Scan(categories, null, CancellationToken.None);
        }

        [TestMethod]
        public void Scan_Startup_ReportsOnlyTheMissingFile() {
            using (RegistryKey run = Create(RunKey)) {
                run.SetValue("Vorhanden", "\"" + existing + "\" /x");
                run.SetValue("Verwaist", "\"" + missing + "\" /x");
                run.SetValue("Unklar", "rundll32.exe irgendwas.dll,Start");
                run.SetValue("Netz", @"\\server\freigabe\app.exe");
            }

            RegistryIssue issue = Scan(RegistryCategory.StartupEntries).Single();

            Assert.AreEqual("Verwaist", issue.ValueName);
            Assert.AreEqual(missing, issue.MissingPath);
        }

        [TestMethod]
        public void Scan_Uninstall_SkipsSystemComponentsAndUnclearCommands() {
            using (RegistryKey key = Create(UninstallKey + @"\Verwaist")) {
                key.SetValue("DisplayName", "Altes Programm");
                key.SetValue("UninstallString", "\"" + missing + "\"");
            }
            using (RegistryKey key = Create(UninstallKey + @"\System")) {
                key.SetValue("SystemComponent", 1);
                key.SetValue("UninstallString", "\"" + missing + "\"");
            }
            using (RegistryKey key = Create(UninstallKey + @"\Msi")) {
                key.SetValue("UninstallString", "MsiExec.exe /X{12345678-1234-1234-1234-123456789012}");
            }

            RegistryIssue issue = Scan(RegistryCategory.ObsoleteSoftware).Single();

            Assert.AreEqual("Altes Programm", issue.Name);
            Assert.IsTrue(issue.IsWholeKey);
        }

        [TestMethod]
        public void Scan_ComClass_KeepsAdvertisedAndPartlyValidClasses() {
            string verwaist = ClsidKey + @"\{00000000-0000-0000-0000-000000000001}";
            using (RegistryKey key = Create(verwaist + @"\InprocServer32")) {
                key.SetValue(string.Empty, missing);
            }
            // Über den Windows Installer angeboten: Die Datei kommt bei Bedarf zurück.
            using (RegistryKey key = Create(ClsidKey + @"\{00000000-0000-0000-0000-000000000002}\InprocServer32")) {
                key.SetValue(string.Empty, missing);
                key.SetValue("InprocServer32", new[] { "beschreibung" }, RegistryValueKind.MultiString);
            }
            // Einer von zwei Servern ist noch da.
            using (RegistryKey key = Create(ClsidKey + @"\{00000000-0000-0000-0000-000000000003}\InprocServer32")) {
                key.SetValue(string.Empty, missing);
            }
            using (RegistryKey key = Create(ClsidKey + @"\{00000000-0000-0000-0000-000000000003}\LocalServer32")) {
                key.SetValue(string.Empty, existing);
            }

            RegistryIssue issue = Scan(RegistryCategory.ComClasses).Single();

            Assert.AreEqual(verwaist, issue.KeyPath);
        }

        [TestMethod]
        public void Scan_ApplicationPaths_ReportsWholeKeyOfMissingProgram() {
            string appPaths = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths";
            using (RegistryKey key = Create(appPaths + @"\alt.exe")) {
                key.SetValue(string.Empty, "\"" + missing + "\"");
            }
            using (RegistryKey key = Create(appPaths + @"\gut.exe")) {
                key.SetValue(string.Empty, existing);
            }

            RegistryIssue issue = Scan(RegistryCategory.ApplicationPaths).Single();

            Assert.AreEqual(appPaths + @"\alt.exe", issue.KeyPath);
            Assert.IsTrue(issue.IsWholeKey);
        }

        [TestMethod]
        public void Scan_Applications_NeedsAllCommandsToBeMissing() {
            string applications = @"SOFTWARE\Classes\Applications";
            using (RegistryKey key = Create(applications + @"\alt.exe\shell\open\command")) {
                key.SetValue(string.Empty, "\"" + missing + "\" \"%1\"");
            }
            using (RegistryKey key = Create(applications + @"\gemischt.exe\shell\open\command")) {
                key.SetValue(string.Empty, "\"" + missing + "\"");
            }
            using (RegistryKey key = Create(applications + @"\gemischt.exe\shell\edit\command")) {
                key.SetValue(string.Empty, "\"" + existing + "\"");
            }

            RegistryIssue issue = Scan(RegistryCategory.Applications).Single();

            Assert.AreEqual("alt.exe", issue.Name);
        }

        [TestMethod]
        public void Scan_SharedDlls_ReadsThePathFromTheValueName() {
            using (RegistryKey key = Create(SharedDllsKey)) {
                key.SetValue(missing, 1);
                key.SetValue(existing, 1);
            }

            RegistryIssue issue = Scan(RegistryCategory.SharedDlls).Single();

            Assert.AreEqual(missing, issue.ValueName);
            Assert.AreEqual("fehlt.exe", issue.Name);
        }

        [TestMethod]
        public void Scan_Fonts_ResolvesFileNamesAgainstTheFontsFolder() {
            using (RegistryKey key = Create(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")) {
                key.SetValue("Verwaist (TrueType)", "gibt-es-nicht-" + Guid.NewGuid().ToString("N") + ".ttf");
                key.SetValue("Vorhanden (TrueType)", existing);
            }

            RegistryIssue issue = Scan(RegistryCategory.Fonts).Single();

            Assert.AreEqual("Verwaist (TrueType)", issue.ValueName);
            StringAssert.StartsWith(issue.MissingPath, Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts"));
        }

        [TestMethod]
        public void Scan_HelpFiles_CombinesNameAndFolder() {
            File.WriteAllText(Path.Combine(folder, "gut.hlp"), string.Empty);
            using (RegistryKey key = Create(@"SOFTWARE\Microsoft\Windows\Help")) {
                key.SetValue("alt.hlp", folder);
                key.SetValue("gut.hlp", folder);
            }

            RegistryIssue issue = Scan(RegistryCategory.HelpFiles).Single();

            Assert.AreEqual(Path.Combine(folder, "alt.hlp"), issue.MissingPath);
        }

        [TestMethod]
        public void Scan_Sounds_ReportsTheDefaultValueOfTheEvent() {
            string keyPath = @"AppEvents\Schemes\Apps\App\Ereignis\.Current";
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(sandboxKey + @"\" + RegistryHive.CurrentUser + @"\" + keyPath)) {
                key.SetValue(string.Empty, Path.Combine(folder, "fehlt.wav"));
            }

            RegistryIssue issue = Scan(RegistryCategory.SoundEvents).Single();

            Assert.AreEqual(keyPath, issue.KeyPath);
            Assert.AreEqual(string.Empty, issue.ValueName);
            Assert.AreEqual("Ereignis", issue.Name);
        }

        [TestMethod]
        public void Scan_MuiCache_NamesTheTwoValuesOfAProgramDifferently() {
            string keyPath = @"Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache";
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(sandboxKey + @"\" + RegistryHive.CurrentUser + @"\" + keyPath)) {
                key.SetValue(missing + ".FriendlyAppName", "Programm");
                key.SetValue(missing + ".ApplicationCompany", "Firma");
                key.SetValue(existing + ".FriendlyAppName", "Bleibt");
                key.SetValue("LangID", new byte[] { 7, 4 }, RegistryValueKind.Binary);
            }

            List<string> names = Scan(RegistryCategory.MuiCache).Select(issue => issue.Name).OrderBy(name => name).ToList();

            CollectionAssert.AreEqual(new[] { "fehlt.exe (Firma)", "fehlt.exe (Name)" }, names);
        }

        [TestMethod]
        public void Scan_TypeLibrary_ReportsTheVersionAndIgnoresTheResourceIndex() {
            string library = @"SOFTWARE\Classes\TypeLib\{00000000-0000-0000-0000-000000000009}";
            using (RegistryKey key = Create(library + @"\1.0\0\win32")) {
                key.SetValue(string.Empty, missing + @"\2");
            }
            using (RegistryKey key = Create(library + @"\2.0\0\win32")) {
                key.SetValue(string.Empty, existing + @"\2");
            }

            RegistryIssue issue = Scan(RegistryCategory.TypeLibraries).Single();

            Assert.AreEqual(library + @"\1.0", issue.KeyPath);
            Assert.AreEqual(missing, issue.MissingPath);
        }

        [TestMethod]
        public void Check_WindowsFolder_OnlyCountsWhenAllowed() {
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "gibt-es-nicht.dll");
            var probe = new RegistryPathProbe();

            Assert.AreEqual(RegistryPathState.Unknown, probe.Check(path, false));
            Assert.AreEqual(RegistryPathState.Missing, probe.Check(path, true));
        }

        [DataTestMethod]
        [DataRow(@"\\server\freigabe\app.exe")]
        [DataRow("app.exe")]
        [DataRow(@"%GIBT_ES_NICHT%\app.exe")]
        [DataRow(@"C:\Program Files\WindowsApps\paket\app.exe")]
        public void Check_Undecidable_IsUnknown(string path) {
            Assert.AreEqual(RegistryPathState.Unknown, new RegistryPathProbe().Check(path, true));
        }

        [TestMethod]
        public void Clean_RemovesOnlyTheSelectedEntryAndWritesBackupFirst() {
            using (RegistryKey run = Create(RunKey)) {
                run.SetValue("Gueltig", "\"" + existing + "\"");
                run.SetValue("Verwaist", "\"" + missing + "\"");
            }

            RegistryCleanResult result = cleaner.Clean(Scan(RegistryCategory.StartupEntries));

            Assert.AreEqual(1, result.Removed.Count);
            Assert.IsTrue(File.Exists(result.BackupPath));
            using (RegistryKey run = Create(RunKey)) {
                CollectionAssert.AreEqual(new[] { "Gueltig" }, run.GetValueNames());
            }
        }

        [TestMethod]
        public void Clean_FileIsBackMeanwhile_KeepsTheEntry() {
            using (RegistryKey run = Create(RunKey)) {
                run.SetValue("Verwaist", "\"" + missing + "\"");
            }
            List<RegistryIssue> issues = Scan(RegistryCategory.StartupEntries);
            File.WriteAllText(missing, string.Empty);

            RegistryCleanResult result = cleaner.Clean(issues);

            Assert.AreEqual(0, result.Removed.Count);
            Assert.AreEqual(1, result.Skipped.Count);
            Assert.IsNull(result.BackupPath);
            using (RegistryKey run = Create(RunKey)) {
                CollectionAssert.Contains(run.GetValueNames(), "Verwaist");
            }
        }

        [TestMethod]
        public void Clean_OutsideOfCheckedLocations_IsRefused() {
            string fremd = @"SOFTWARE\Fremd\Hersteller";
            using (Create(fremd)) {
            }
            var issue = new RegistryIssue(RegistryCategory.ObsoleteSoftware, RegistryHive.LocalMachine, fremd, null, "Fremd", missing);

            RegistryCleanResult result = cleaner.Clean(new[] { issue });

            Assert.AreEqual(1, result.Failed.Count);
            using (RegistryKey key = scope.Open(RegistryHive.LocalMachine, fremd)) {
                Assert.IsNotNull(key);
            }
        }

        [TestMethod]
        public void IsAllowed_TheCheckedLocationItselfCanNeverBeRemoved() {
            var wholeLocation = new RegistryIssue(RegistryCategory.ObsoleteSoftware, RegistryHive.LocalMachine, UninstallKey, null, "Alles", missing);
            var entryBelow = new RegistryIssue(RegistryCategory.ObsoleteSoftware, RegistryHive.LocalMachine, UninstallKey + @"\App", null, "App", missing);
            var valueInside = new RegistryIssue(RegistryCategory.SharedDlls, RegistryHive.LocalMachine, SharedDllsKey, missing, "App", missing);
            var otherHive = new RegistryIssue(RegistryCategory.SharedDlls, RegistryHive.CurrentUser, SharedDllsKey, missing, "App", missing);

            Assert.IsFalse(RegistryLocations.IsAllowed(wholeLocation));
            Assert.IsTrue(RegistryLocations.IsAllowed(entryBelow));
            Assert.IsTrue(RegistryLocations.IsAllowed(valueInside));
            Assert.IsFalse(RegistryLocations.IsAllowed(otherHive));
        }

        [TestMethod]
        public void Restore_BringsBackAllValueTypesAndSubKeys() {
            string entry = UninstallKey + @"\Verwaist";
            byte[] binary = { 1, 2, 250 };
            string[] multi = { "eins", "zwei \"mit\" Zeichen" };
            using (RegistryKey key = Create(entry)) {
                key.SetValue("UninstallString", "\"" + missing + "\"");
                key.SetValue("Text", @"C:\Pfad ""mit"" Zeichen");
                key.SetValue("Variable", @"%TEMP%\x", RegistryValueKind.ExpandString);
                key.SetValue("Zahl", unchecked((int)0xF0000001), RegistryValueKind.DWord);
                key.SetValue("Gross", 0x1122334455667788L, RegistryValueKind.QWord);
                key.SetValue("Binaer", binary, RegistryValueKind.Binary);
                key.SetValue("Liste", multi, RegistryValueKind.MultiString);
                key.SetValue(string.Empty, "Standardwert");
            }
            using (RegistryKey sub = Create(entry + @"\Unter")) {
                sub.SetValue("Wert", "drin");
            }

            RegistryCleanResult result = cleaner.Clean(Scan(RegistryCategory.ObsoleteSoftware));
            using (RegistryKey gone = scope.Open(RegistryHive.LocalMachine, entry)) {
                Assert.IsNull(gone);
            }

            ProcessResult restore = RegistryBackup.Restore(result.BackupPath);
            Assert.AreEqual(0, restore.ExitCode, restore.ErrorOutput);

            using (RegistryKey key = scope.Open(RegistryHive.LocalMachine, entry)) {
                Assert.AreEqual(@"C:\Pfad ""mit"" Zeichen", key.GetValue("Text"));
                Assert.AreEqual(@"%TEMP%\x", key.GetValue("Variable", null, RegistryValueOptions.DoNotExpandEnvironmentNames));
                Assert.AreEqual(RegistryValueKind.ExpandString, key.GetValueKind("Variable"));
                Assert.AreEqual(unchecked((int)0xF0000001), key.GetValue("Zahl"));
                Assert.AreEqual(0x1122334455667788L, key.GetValue("Gross"));
                CollectionAssert.AreEqual(binary, (byte[])key.GetValue("Binaer"));
                CollectionAssert.AreEqual(multi, (string[])key.GetValue("Liste"));
                Assert.AreEqual("Standardwert", key.GetValue(string.Empty));
            }
            using (RegistryKey sub = scope.Open(RegistryHive.LocalMachine, entry + @"\Unter")) {
                Assert.AreEqual("drin", sub.GetValue("Wert"));
            }
        }

        [TestMethod]
        public void Restore_ValueEntry_LeavesLaterChangesOfOtherValuesAlone() {
            using (RegistryKey run = Create(RunKey)) {
                run.SetValue("Verwaist", "\"" + missing + "\"");
                run.SetValue("Anderer", "alt");
            }
            RegistryCleanResult result = cleaner.Clean(Scan(RegistryCategory.StartupEntries));
            using (RegistryKey run = Create(RunKey)) {
                run.SetValue("Anderer", "neu");
            }

            Assert.AreEqual(0, RegistryBackup.Restore(result.BackupPath).ExitCode);

            using (RegistryKey run = Create(RunKey)) {
                Assert.AreEqual("neu", run.GetValue("Anderer"));
                CollectionAssert.Contains(run.GetValueNames(), "Verwaist");
            }
        }

        [TestMethod]
        public void FormatValue_UnknownKind_RefusesInsteadOfWritingAWrongBackup() {
            Assert.ThrowsExactly<NotSupportedException>(() => RegistryBackup.FormatValue("x", RegistryValueKind.Unknown, new byte[0]));
        }

        [TestMethod]
        public void FindLatestBackup_ReturnsTheNewestFile() {
            string directory = Path.Combine(folder, "Liste");
            Assert.IsNull(RegistryCleaner.FindLatestBackup(directory));

            Directory.CreateDirectory(directory);
            File.WriteAllText(Path.Combine(directory, "registry-20260101-000000-000.reg"), string.Empty);
            string newest = Path.Combine(directory, "registry-20261105-120000-000.reg");
            File.WriteAllText(newest, string.Empty);

            Assert.AreEqual(newest, RegistryCleaner.FindLatestBackup(directory));
        }
    }
}
