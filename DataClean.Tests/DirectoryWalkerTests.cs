using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Principal;
using DataClean.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DataClean.Tests {
    [TestClass]
    public class DirectoryWalkerTests {
        private string root;

        [TestInitialize]
        public void Setup() {
            root = Path.Combine(Path.GetTempPath(), "DataCleanWalkerTest_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
        }

        [TestCleanup]
        public void Cleanup() {
            if (Directory.Exists(root)) {
                Directory.Delete(root, true);
            }
        }

        [TestMethod]
        public void Walk_ReturnsRootAndAllSubdirectories() {
            Directory.CreateDirectory(Path.Combine(root, "a", "b"));
            Directory.CreateDirectory(Path.Combine(root, "c"));

            var names = DirectoryWalker.Walk(new DirectoryInfo(root))
                .Select(d => d.FullName.Substring(root.Length).TrimStart('\\'))
                .OrderBy(n => n)
                .ToList();

            CollectionAssert.AreEqual(new[] { "", "a", "a\\b", "c" }, names);
        }

        [TestMethod]
        public void Walk_DoesNotFollowJunctions() {
            string target = Path.Combine(root, "ziel");
            string area = Path.Combine(root, "bereich");
            Directory.CreateDirectory(Path.Combine(target, "innen"));
            Directory.CreateDirectory(area);
            string link = Path.Combine(area, "verknuepfung");
            RunCmd($"mklink /J \"{link}\" \"{target}\"");
            Assert.IsTrue(Directory.Exists(link), "Junction konnte nicht angelegt werden.");

            try {
                var found = DirectoryWalker.Walk(new DirectoryInfo(area)).Select(d => d.FullName).ToList();
                CollectionAssert.AreEqual(new[] { area }, found);
            } finally {
                Directory.Delete(link);
            }
        }

        [TestMethod]
        public void Walk_SkipsUnreadableDirectoryAndContinues() {
            string denied = Path.Combine(root, "gesperrt");
            string readable = Path.Combine(root, "lesbar");
            Directory.CreateDirectory(denied);
            Directory.CreateDirectory(readable);

            var rule = new FileSystemAccessRule(WindowsIdentity.GetCurrent().User,
                FileSystemRights.ListDirectory, AccessControlType.Deny);
            var info = new DirectoryInfo(denied);
            DirectorySecurity security = info.GetAccessControl();
            security.AddAccessRule(rule);
            info.SetAccessControl(security);

            try {
                string unreadable = null;
                var found = DirectoryWalker.Walk(new DirectoryInfo(root), d => unreadable = d.FullName)
                    .Select(d => d.FullName).ToList();

                Assert.AreEqual(denied, unreadable);
                CollectionAssert.Contains(found, readable);
                CollectionAssert.DoesNotContain(found, denied);
            } finally {
                security.RemoveAccessRule(rule);
                info.SetAccessControl(security);
            }
        }

        [TestMethod]
        public void EnumerateFiles_ReturnsOnlyDirectFiles() {
            File.WriteAllText(Path.Combine(root, "oben.txt"), "x");
            Directory.CreateDirectory(Path.Combine(root, "unter"));
            File.WriteAllText(Path.Combine(root, "unter", "unten.txt"), "x");

            var names = DirectoryWalker.EnumerateFiles(new DirectoryInfo(root)).Select(f => f.Name).ToList();

            CollectionAssert.AreEqual(new[] { "oben.txt" }, names);
        }

        [TestMethod]
        public void EnumerateFiles_MissingDirectory_ReturnsEmpty() {
            var files = DirectoryWalker.EnumerateFiles(new DirectoryInfo(Path.Combine(root, "fehlt")));
            Assert.AreEqual(0, files.Count());
        }

        private static void RunCmd(string arguments) {
            var psi = new ProcessStartInfo("cmd.exe", "/c " + arguments) {
                CreateNoWindow = true,
                UseShellExecute = false
            };
            using (Process process = Process.Start(psi)) {
                process.WaitForExit();
            }
        }
    }
}
