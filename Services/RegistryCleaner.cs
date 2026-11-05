using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using Microsoft.Win32;

namespace DataClean.Services {
    public sealed class RegistryCleanResult {
        public List<RegistryIssue> Removed { get; } = new List<RegistryIssue>();

        /// <summary>Einträge, die inzwischen keinen Fund mehr darstellen, etwa weil die Datei wieder da ist.</summary>
        public List<RegistryIssue> Skipped { get; } = new List<RegistryIssue>();

        public List<KeyValuePair<RegistryIssue, string>> Failed { get; } = new List<KeyValuePair<RegistryIssue, string>>();

        public string BackupPath { get; set; }
    }

    /// <summary>
    /// Entfernt ausgewählte Funde. Zuerst entsteht die Sicherung; misslingt sie, wird nichts entfernt.
    /// Wird nur nach ausdrücklicher Bestätigung aufgerufen, nie aus dem unbeaufsichtigten Lauf.
    /// </summary>
    public sealed class RegistryCleaner {
        public static readonly string DefaultBackupDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DataClean", "RegistryBackups");

        private readonly RegistryScope scope;
        private readonly RegistryPathProbe probe;
        private readonly string backupDirectory;

        public RegistryCleaner() : this(RegistryScope.System, new RegistryPathProbe(), DefaultBackupDirectory) {
        }

        public RegistryCleaner(RegistryScope scope, RegistryPathProbe probe, string backupDirectory) {
            this.scope = scope;
            this.probe = probe;
            this.backupDirectory = backupDirectory;
        }

        /// <summary>Die neueste Sicherung, oder null, wenn es keine gibt.</summary>
        public static string FindLatestBackup(string directory) {
            return Directory.Exists(directory)
                ? Directory.GetFiles(directory, "registry-*.reg").OrderByDescending(file => file, StringComparer.OrdinalIgnoreCase).FirstOrDefault()
                : null;
        }

        public RegistryCleanResult Clean(IEnumerable<RegistryIssue> issues) {
            var result = new RegistryCleanResult();
            var confirmed = new List<RegistryIssue>();

            foreach (RegistryIssue issue in issues) {
                if (!RegistryLocations.IsAllowed(issue)) {
                    result.Failed.Add(new KeyValuePair<RegistryIssue, string>(issue, "Liegt außerhalb der geprüften Bereiche."));
                } else if (probe.Check(issue.MissingPath, true) != RegistryPathState.Missing) {
                    // Zwischen Prüfung und Entfernen kann die Datei zurückgekehrt sein.
                    result.Skipped.Add(issue);
                } else {
                    confirmed.Add(issue);
                }
            }

            if (confirmed.Count == 0) {
                return result;
            }

            result.BackupPath = Path.Combine(backupDirectory, $"registry-{DateTime.Now:yyyyMMdd-HHmmss-fff}.reg");
            RegistryBackup.Write(result.BackupPath, confirmed, scope);

            foreach (RegistryIssue issue in confirmed) {
                string fullName = scope.FullName(issue.Hive, issue.KeyPath);
                try {
                    Remove(issue);
                    result.Removed.Add(issue);
                    Logger.Info($"Registry-Eintrag entfernt: {fullName} [{issue.ValueName}]", "Registry");
                } catch (Exception ex) when (ex is SecurityException || ex is UnauthorizedAccessException || ex is IOException) {
                    result.Failed.Add(new KeyValuePair<RegistryIssue, string>(issue, ex.Message));
                    Logger.Warn($"Registry-Eintrag nicht entfernt: {fullName}: {ex.Message}", "Registry");
                }
            }
            return result;
        }

        private void Remove(RegistryIssue issue) {
            if (issue.IsWholeKey) {
                int separator = issue.KeyPath.LastIndexOf('\\');
                using (RegistryKey parent = scope.Open(issue.Hive, issue.KeyPath.Substring(0, separator), writable: true)) {
                    parent?.DeleteSubKeyTree(issue.KeyPath.Substring(separator + 1), throwOnMissingSubKey: false);
                }
            } else {
                using (RegistryKey key = scope.Open(issue.Hive, issue.KeyPath, writable: true)) {
                    key?.DeleteValue(issue.ValueName, throwOnMissingValue: false);
                }
            }
        }
    }
}
