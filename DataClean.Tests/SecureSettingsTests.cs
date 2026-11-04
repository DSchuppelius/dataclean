using Microsoft.VisualStudio.TestTools.UnitTesting;
using DataClean.Services;

namespace DataClean.Tests {
    [TestClass]
    public class SecureSettingsTests {
        [TestMethod]
        public void Encrypt_NullOrEmpty_ReturnsEmpty() {
            Assert.AreEqual(string.Empty, SecureSettings.Encrypt(null));
            Assert.AreEqual(string.Empty, SecureSettings.Encrypt(string.Empty));
        }

        [TestMethod]
        public void Decrypt_NullOrEmpty_ReturnsEmpty() {
            Assert.AreEqual(string.Empty, SecureSettings.Decrypt(null));
            Assert.AreEqual(string.Empty, SecureSettings.Decrypt(string.Empty));
        }

        [TestMethod]
        public void EncryptThenDecrypt_RoundTrips() {
            const string secret = "P@ssw0rd!_äöü_123";
            string encrypted = SecureSettings.Encrypt(secret);

            Assert.IsFalse(string.IsNullOrEmpty(encrypted));
            Assert.AreNotEqual(secret, encrypted);
            Assert.AreEqual(secret, SecureSettings.Decrypt(encrypted));
        }

        [TestMethod]
        public void Encrypt_ProducesBase64() {
            string encrypted = SecureSettings.Encrypt("hello");
            // Base64-dekodierbar ohne Ausnahme.
            var bytes = System.Convert.FromBase64String(encrypted);
            Assert.IsTrue(bytes.Length > 0);
        }

        [TestMethod]
        public void Decrypt_InvalidData_ReturnsEmpty() {
            // Der absichtlich provozierte Fehler soll nicht im echten Anwendungsprotokoll landen.
            LogLevel previous = Logger.MinimumLevel;
            Logger.MinimumLevel = LogLevel.None;
            try {
                Assert.AreEqual(string.Empty, SecureSettings.Decrypt("kein-gültiges-base64!!!"));
            } finally {
                Logger.MinimumLevel = previous;
            }
        }
    }
}
