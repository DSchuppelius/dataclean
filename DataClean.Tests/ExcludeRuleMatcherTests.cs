using Microsoft.VisualStudio.TestTools.UnitTesting;
using DataClean.Services;

namespace DataClean.Tests {
    [TestClass]
    public class ExcludeRuleMatcherTests {
        [TestMethod]
        public void NoPatternsMeansNoExclusion() {
            var matcher = new ExcludeRuleMatcher(null);
            Assert.IsFalse(matcher.HasRules);
            Assert.IsFalse(matcher.IsExcluded(@"C:\Temp\file.txt"));
        }

        [TestMethod]
        public void StarMatchesExtensionOnFileName() {
            var matcher = new ExcludeRuleMatcher(new[] { "*.keep" });
            Assert.IsTrue(matcher.IsExcluded(@"C:\Temp\wichtig.keep"));
            Assert.IsFalse(matcher.IsExcluded(@"C:\Temp\wichtig.txt"));
        }

        [TestMethod]
        public void StarDoesNotCrossDirectorySeparator() {
            var matcher = new ExcludeRuleMatcher(new[] { @"C:\Temp\*.log" });
            Assert.IsTrue(matcher.IsExcluded(@"C:\Temp\run.log"));
            Assert.IsFalse(matcher.IsExcluded(@"C:\Temp\sub\run.log"));
        }

        [TestMethod]
        public void DoubleStarCrossesDirectorySeparator() {
            var matcher = new ExcludeRuleMatcher(new[] { @"C:\Temp\**\run.log" });
            Assert.IsTrue(matcher.IsExcluded(@"C:\Temp\sub\deep\run.log"));
        }

        [TestMethod]
        public void QuestionMarkMatchesSingleCharacter() {
            var matcher = new ExcludeRuleMatcher(new[] { "file?.txt" });
            Assert.IsTrue(matcher.IsExcluded(@"C:\a\file1.txt"));
            Assert.IsFalse(matcher.IsExcluded(@"C:\a\file12.txt"));
        }

        [TestMethod]
        public void MatchingIsCaseInsensitive() {
            var matcher = new ExcludeRuleMatcher(new[] { "*.KEEP" });
            Assert.IsTrue(matcher.IsExcluded(@"C:\a\data.keep"));
        }

        [TestMethod]
        public void InvalidPatternsAreIgnoredGracefully() {
            var matcher = new ExcludeRuleMatcher(new[] { "", "   ", "*.ok" });
            Assert.IsTrue(matcher.HasRules);
            Assert.IsTrue(matcher.IsExcluded(@"C:\a\x.ok"));
        }
    }
}
