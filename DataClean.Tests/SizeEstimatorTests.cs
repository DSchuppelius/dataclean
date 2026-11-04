using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DataClean.Services;

namespace DataClean.Tests {
    [TestClass]
    public class SizeEstimatorTests {
        private string tempRoot;

        [TestInitialize]
        public void Setup() {
            tempRoot = Path.Combine(Path.GetTempPath(), "DataCleanTest_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempRoot);
        }

        [TestCleanup]
        public void Cleanup() {
            try {
                if (Directory.Exists(tempRoot)) {
                    Directory.Delete(tempRoot, true);
                }
            } catch {
                // Aufräumfehler im Test ignorieren.
            }
        }

        [TestMethod]
        public void EstimateSize_NullList_ReturnsZero() {
            Assert.AreEqual(0, SizeEstimator.EstimateSize(null));
        }

        [TestMethod]
        public void EstimateSize_EmptyDirectory_ReturnsZero() {
            var dirs = new List<DirectoryInfo> { new DirectoryInfo(tempRoot) };
            Assert.AreEqual(0, SizeEstimator.EstimateSize(dirs));
        }

        [TestMethod]
        public void EstimateSize_NonExistentDirectory_ReturnsZero() {
            var dirs = new List<DirectoryInfo> {
                new DirectoryInfo(Path.Combine(tempRoot, "nichtvorhanden"))
            };
            Assert.AreEqual(0, SizeEstimator.EstimateSize(dirs));
        }

        [TestMethod]
        public void EstimateSize_SingleFile_ReturnsFileSize() {
            string file = Path.Combine(tempRoot, "a.txt");
            File.WriteAllBytes(file, new byte[1234]);

            var dirs = new List<DirectoryInfo> { new DirectoryInfo(tempRoot) };
            Assert.AreEqual(1234, SizeEstimator.EstimateSize(dirs));
        }

        [TestMethod]
        public void EstimateSize_NestedFiles_SumsRecursively() {
            File.WriteAllBytes(Path.Combine(tempRoot, "a.txt"), new byte[100]);
            string sub = Path.Combine(tempRoot, "sub");
            Directory.CreateDirectory(sub);
            File.WriteAllBytes(Path.Combine(sub, "b.txt"), new byte[200]);
            string sub2 = Path.Combine(sub, "deep");
            Directory.CreateDirectory(sub2);
            File.WriteAllBytes(Path.Combine(sub2, "c.txt"), new byte[300]);

            var dirs = new List<DirectoryInfo> { new DirectoryInfo(tempRoot) };
            Assert.AreEqual(600, SizeEstimator.EstimateSize(dirs));
        }
    }
}
