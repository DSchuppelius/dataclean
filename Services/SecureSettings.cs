using System;
using System.Security.Cryptography;
using System.Text;
namespace DataClean.Services {
    /// <summary>
    /// Verschlüsselt und entschlüsselt sensible Zeichenfolgen (z. B. das SMTP-Passwort)
    /// mithilfe der Windows Data Protection API (DPAPI). Die Daten werden an den aktuellen
    /// Benutzer gebunden (<see cref="DataProtectionScope.CurrentUser"/>), sodass nur dieser
    /// sie wieder entschlüsseln kann. Die verschlüsselten Werte werden Base64-kodiert, um
    /// sie als String in den Anwendungseinstellungen ablegen zu können.
    /// </summary>
    public static class SecureSettings {
        // Zusätzlicher Entropie-Wert, der die Verschlüsselung an diese Anwendung bindet.
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("DataClean.SecureSettings.v1");

        /// <summary>
        /// Verschlüsselt den Klartext und gibt ihn Base64-kodiert zurück.
        /// Ein leerer oder null-Wert ergibt eine leere Zeichenfolge.
        /// </summary>
        public static string Encrypt(string plainText) {
            if (string.IsNullOrEmpty(plainText)) {
                return string.Empty;
            }

            try {
                byte[] data = Encoding.UTF8.GetBytes(plainText);
                byte[] encrypted = ProtectedData.Protect(data, Entropy, DataProtectionScope.CurrentUser);
                return Convert.ToBase64String(encrypted);
            } catch (Exception ex) {
                Logger.Error(ex, "SecureSettings.Encrypt");
                return string.Empty;
            }
        }

        /// <summary>
        /// Entschlüsselt einen zuvor mit <see cref="Encrypt"/> erzeugten Base64-Wert.
        /// Bei ungültigen Daten oder leerem Wert wird eine leere Zeichenfolge zurückgegeben.
        /// </summary>
        public static string Decrypt(string encryptedBase64) {
            if (string.IsNullOrEmpty(encryptedBase64)) {
                return string.Empty;
            }

            try {
                byte[] encrypted = Convert.FromBase64String(encryptedBase64);
                byte[] data = ProtectedData.Unprotect(encrypted, Entropy, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(data);
            } catch (Exception ex) {
                Logger.Error(ex, "SecureSettings.Decrypt");
                return string.Empty;
            }
        }
    }
}
