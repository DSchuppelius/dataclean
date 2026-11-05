using Microsoft.Win32;

namespace DataClean.Services {
    /// <summary>
    /// Bestimmt, welche Registry bearbeitet wird. Tests leiten alle Zugriffe in einen eigenen Schlüssel
    /// unter HKCU um und berühren die echten Einträge nie.
    /// </summary>
    public sealed class RegistryScope {
        public static readonly RegistryScope System = new RegistryScope(null);

        private readonly string sandboxKey;

        private RegistryScope(string sandboxKey) {
            this.sandboxKey = sandboxKey;
        }

        public static RegistryScope Sandbox(string keyBelowCurrentUser) {
            return new RegistryScope(keyBelowCurrentUser);
        }

        /// <summary>Öffnet einen Schlüssel oder liefert null, wenn es ihn nicht gibt.</summary>
        public RegistryKey Open(RegistryHive hive, string keyPath, bool writable = false) {
            // Die 64-Bit-Sicht zeigt alle Einträge an ihrem echten Ort, also auch WOW6432Node.
            using (RegistryKey root = RegistryKey.OpenBaseKey(PhysicalHive(hive), RegistryView.Registry64)) {
                return root.OpenSubKey(PhysicalPath(hive, keyPath), writable);
            }
        }

        /// <summary>Vollständiger Name des Schlüssels, wie er in einer .reg-Datei steht.</summary>
        public string FullName(RegistryHive hive, string keyPath) {
            string hiveName = PhysicalHive(hive) == RegistryHive.LocalMachine ? "HKEY_LOCAL_MACHINE" : "HKEY_CURRENT_USER";
            return hiveName + @"\" + PhysicalPath(hive, keyPath);
        }

        private RegistryHive PhysicalHive(RegistryHive hive) {
            return sandboxKey == null ? hive : RegistryHive.CurrentUser;
        }

        private string PhysicalPath(RegistryHive hive, string keyPath) {
            return sandboxKey == null ? keyPath : sandboxKey + @"\" + hive + @"\" + keyPath;
        }
    }
}
