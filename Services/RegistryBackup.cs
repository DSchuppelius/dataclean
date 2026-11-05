using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace DataClean.Services {
    /// <summary>
    /// Schreibt die betroffenen Einträge vor dem Entfernen in eine .reg-Datei. Es werden nur die Einträge
    /// gesichert, die auch entfernt werden: Würde der ganze Schlüssel exportiert, setzte das Einspielen
    /// spätere Änderungen an seinen übrigen Werten zurück.
    /// </summary>
    public static class RegistryBackup {
        private const string Header = "Windows Registry Editor Version 5.00";

        public static void Write(string filePath, IEnumerable<RegistryIssue> issues, RegistryScope scope) {
            var text = new StringBuilder();
            text.Append(Header).Append("\r\n");

            foreach (RegistryIssue issue in issues) {
                using (RegistryKey key = scope.Open(issue.Hive, issue.KeyPath)) {
                    if (key == null) {
                        continue;
                    }
                    string name = scope.FullName(issue.Hive, issue.KeyPath);
                    if (issue.IsWholeKey) {
                        AppendKey(text, key, name);
                    } else if (key.GetValueNames().Contains(issue.ValueName, StringComparer.OrdinalIgnoreCase)) {
                        text.Append("\r\n[").Append(name).Append("]\r\n");
                        AppendValue(text, key, issue.ValueName);
                    }
                }
            }

            Directory.CreateDirectory(Path.GetDirectoryName(filePath));
            // Regedit erwartet UTF-16 mit BOM.
            File.WriteAllText(filePath, text.ToString(), Encoding.Unicode);
        }

        /// <summary>Spielt eine Sicherung über reg.exe wieder ein.</summary>
        public static ProcessResult Restore(string filePath) {
            string regExe = Path.Combine(Environment.SystemDirectory, "reg.exe");
            return ProcessRunner.Run(regExe, "import \"" + filePath + "\"", line => { }, TimeSpan.FromMinutes(2));
        }

        private static void AppendKey(StringBuilder text, RegistryKey key, string fullName) {
            text.Append("\r\n[").Append(fullName).Append("]\r\n");
            foreach (string valueName in key.GetValueNames()) {
                AppendValue(text, key, valueName);
            }
            foreach (string subName in key.GetSubKeyNames()) {
                using (RegistryKey sub = key.OpenSubKey(subName)) {
                    if (sub != null) {
                        AppendKey(text, sub, fullName + @"\" + subName);
                    }
                }
            }
        }

        private static void AppendValue(StringBuilder text, RegistryKey key, string valueName) {
            RegistryValueKind kind = key.GetValueKind(valueName);
            object data = key.GetValue(valueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
            text.Append(FormatValue(valueName, kind, data)).Append("\r\n");
        }

        /// <summary>Eine Zeile im .reg-Format. Unbekannte Werttypen lassen sich nicht sicher wiederherstellen.</summary>
        public static string FormatValue(string name, RegistryValueKind kind, object data) {
            string left = name.Length == 0 ? "@" : "\"" + Escape(name) + "\"";
            switch (kind) {
                case RegistryValueKind.String:
                    string value = (string)data;
                    return value.Any(char.IsControl)
                        ? left + "=hex(1):" + Hex(Encoding.Unicode.GetBytes(value + "\0"))
                        : left + "=\"" + Escape(value) + "\"";
                case RegistryValueKind.ExpandString:
                    return left + "=hex(2):" + Hex(Encoding.Unicode.GetBytes((string)data + "\0"));
                case RegistryValueKind.DWord:
                    return left + "=dword:" + unchecked((uint)(int)data).ToString("x8");
                case RegistryValueKind.QWord:
                    return left + "=hex(b):" + Hex(BitConverter.GetBytes((long)data));
                case RegistryValueKind.MultiString:
                    string joined = string.Concat(((string[])data).Select(part => part + "\0")) + "\0";
                    return left + "=hex(7):" + Hex(Encoding.Unicode.GetBytes(joined));
                case RegistryValueKind.Binary:
                    return left + "=hex:" + Hex((byte[])data);
                case RegistryValueKind.None:
                    return left + "=hex(0):" + Hex(data as byte[] ?? new byte[0]);
                default:
                    throw new NotSupportedException($"Werttyp {kind} von \"{name}\" lässt sich nicht sichern.");
            }
        }

        private static string Escape(string text) {
            return text.Replace(@"\", @"\\").Replace("\"", "\\\"");
        }

        private static string Hex(byte[] bytes) {
            return string.Join(",", bytes.Select(b => b.ToString("x2")));
        }
    }
}
