using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;

namespace DataClean {
    /// <summary>
    /// Schweregrad eines erkannten System-/Hardwareproblems.
    /// </summary>
    public enum HealthSeverity {
        Warning,
        Critical
    }

    /// <summary>
    /// Beschreibt ein erkanntes System- oder Hardwareproblem.
    /// </summary>
    public class HealthIssue {
        public HealthSeverity Severity { get; }
        public string Category { get; }
        public string Message { get; }

        public HealthIssue(HealthSeverity severity, string category, string message) {
            Severity = severity;
            Category = category;
            Message = message;
        }

        public override string ToString() {
            return $"[{Severity}] {Category}: {Message}";
        }
    }

    /// <summary>
    /// Führt schreibgeschützte Diagnosen der Systemgesundheit durch (Festplatten-SMART,
    /// Datenträgerfehler, Hardware-/Speicherfehler via WHEA, freier Speicherplatz).
    /// Alle Methoden sind defensiv implementiert und werfen keine Ausnahmen nach außen.
    /// </summary>
    public static class HealthCheck {
        /// <summary>
        /// Führt alle aktivierten Prüfungen aus und liefert die gefundenen Probleme.
        /// </summary>
        /// <param name="minFreeDiskPercent">Schwellwert für kritisch wenig freien Speicher.</param>
        public static List<HealthIssue> Run(int minFreeDiskPercent) {
            var issues = new List<HealthIssue>();

            issues.AddRange(CheckSmartStatus());
            issues.AddRange(CheckDiskPredictFailure());
            issues.AddRange(CheckDiskSpace(minFreeDiskPercent));
            issues.AddRange(CheckHardwareEventLog());

            return issues;
        }

        /// <summary>
        /// Prüft den Status aller physischen Datenträger über Win32_DiskDrive.
        /// Ein Status ungleich "OK" deutet auf einen drohenden oder vorhandenen Defekt hin.
        /// </summary>
        private static IEnumerable<HealthIssue> CheckSmartStatus() {
            var result = new List<HealthIssue>();
            try {
                using (var searcher = new ManagementObjectSearcher("SELECT Model, Status FROM Win32_DiskDrive")) {
                    foreach (ManagementObject disk in searcher.Get()) {
                        string model = disk["Model"]?.ToString() ?? "Unbekannt";
                        string status = disk["Status"]?.ToString() ?? "Unbekannt";

                        if (!string.Equals(status, "OK", StringComparison.OrdinalIgnoreCase)) {
                            result.Add(new HealthIssue(HealthSeverity.Critical, "Festplatte",
                                $"Datenträger '{model}' meldet den Status '{status}'."));
                        }
                    }
                }
            } catch (Exception ex) {
                // WMI nicht verfügbar oder keine Berechtigung – nicht als Fehler werten.
                Services.Logger.Warn($"CheckSmartStatus: {ex.Message}", "HealthCheck");
            }
            return result;
        }

        /// <summary>
        /// Wertet die SMART-Ausfallvorhersage über MSStorageDriver_FailurePredictStatus aus (Namespace root\wmi).
        /// Erfordert in der Regel Administratorrechte.
        /// </summary>
        private static IEnumerable<HealthIssue> CheckDiskPredictFailure() {
            var result = new List<HealthIssue>();
            try {
                var scope = new ManagementScope(@"\\.\root\wmi");
                var query = new ObjectQuery("SELECT InstanceName, PredictFailure FROM MSStorageDriver_FailurePredictStatus");
                using (var searcher = new ManagementObjectSearcher(scope, query)) {
                    foreach (ManagementObject obj in searcher.Get()) {
                        bool predictFailure = Convert.ToBoolean(obj["PredictFailure"]);
                        string instance = obj["InstanceName"]?.ToString() ?? "Unbekannt";
                        if (predictFailure) {
                            result.Add(new HealthIssue(HealthSeverity.Critical, "Festplatte (SMART)",
                                $"SMART meldet einen bevorstehenden Ausfall für '{instance}'."));
                        }
                    }
                }
            } catch (Exception ex) {
                Services.Logger.Warn($"CheckDiskPredictFailure: {ex.Message}", "HealthCheck");
            }
            return result;
        }

        /// <summary>
        /// Prüft den freien Speicherplatz aller lokalen Festplatten gegen einen Schwellwert.
        /// </summary>
        private static IEnumerable<HealthIssue> CheckDiskSpace(int minFreeDiskPercent) {
            var result = new List<HealthIssue>();
            if (minFreeDiskPercent <= 0) {
                return result;
            }

            try {
                foreach (var drive in DriveInfo.GetDrives()) {
                    if (!drive.IsReady || drive.DriveType != DriveType.Fixed || drive.TotalSize == 0) {
                        continue;
                    }

                    double freePercent = (double)drive.TotalFreeSpace / drive.TotalSize * 100.0;
                    if (freePercent < minFreeDiskPercent) {
                        result.Add(new HealthIssue(HealthSeverity.Critical, "Speicherplatz",
                            $"Laufwerk {drive.Name} hat nur noch {freePercent:0.#}% freien Speicher."));
                    }
                }
            } catch (Exception ex) {
                Services.Logger.Warn($"CheckDiskSpace: {ex.Message}", "HealthCheck");
            }
            return result;
        }

        /// <summary>
        /// Durchsucht das System-Event-Log der letzten 24 Stunden nach Hardware-/Speicherfehlern
        /// (WHEA-Logger) und schweren Datenträgerfehlern.
        /// </summary>
        private static IEnumerable<HealthIssue> CheckHardwareEventLog() {
            var result = new List<HealthIssue>();
            try {
                using (var log = new System.Diagnostics.EventLog("System")) {
                    DateTime since = DateTime.Now.AddHours(-24);

                    var relevantSources = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                        "Microsoft-Windows-WHEA-Logger",
                        "disk",
                        "Ntfs",
                        "volmgr",
                        "Microsoft-Windows-DiskDiagnosticDataCollector"
                    };

                    int diskErrors = 0;
                    int hardwareErrors = 0;

                    // Betroffene Datenträger erfassen (Bezeichner -> Fehleranzahl),
                    // damit im Ergebnis vermerkt ist, welche Datenträger Fehler meldeten.
                    var affectedDisks = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                    foreach (System.Diagnostics.EventLogEntry entry in log.Entries) {
                        if (entry.TimeGenerated < since) {
                            continue;
                        }

                        if (entry.EntryType != System.Diagnostics.EventLogEntryType.Error) {
                            continue;
                        }

                        if (!relevantSources.Contains(entry.Source)) {
                            continue;
                        }

                        if (entry.Source.Equals("Microsoft-Windows-WHEA-Logger", StringComparison.OrdinalIgnoreCase)) {
                            hardwareErrors++;
                        } else {
                            diskErrors++;

                            string disk = ExtractDiskIdentifier(entry);
                            if (!string.IsNullOrEmpty(disk)) {
                                affectedDisks.TryGetValue(disk, out int count);
                                affectedDisks[disk] = count + 1;
                            }
                        }
                    }

                    if (hardwareErrors > 0) {
                        result.Add(new HealthIssue(HealthSeverity.Critical, "Hardware/Arbeitsspeicher",
                            $"{hardwareErrors} Hardwarefehler (WHEA) in den letzten 24 Stunden erkannt."));
                    }

                    if (diskErrors > 0) {
                        string disksText = affectedDisks.Count > 0
                            ? " Betroffene Datenträger: " + string.Join(", ",
                                affectedDisks.OrderBy(kv => kv.Key)
                                             .Select(kv => $"{kv.Key} ({kv.Value})")) + "."
                            : " Die betroffenen Datenträger konnten nicht eindeutig ermittelt werden.";

                        result.Add(new HealthIssue(HealthSeverity.Critical, "Datenträger",
                            $"{diskErrors} schwere Datenträgerfehler in den letzten 24 Stunden erkannt.{disksText}"));
                    }
                }
            } catch (Exception ex) {
                Services.Logger.Warn($"CheckHardwareEventLog: {ex.Message}", "HealthCheck");
            }
            return result;
        }

        /// <summary>
        /// Versucht, aus einem Datenträger-Fehlereintrag den betroffenen Datenträger zu ermitteln
        /// (z. B. "\Device\Harddisk0\DR0"). Fällt auf die Quelle zurück, wenn kein Bezeichner gefunden wird.
        /// </summary>
        private static string ExtractDiskIdentifier(System.Diagnostics.EventLogEntry entry) {
            try {
                // Zuerst in den Ersetzungszeichenketten nach einem Geräte-/Datenträgerbezeichner suchen.
                if (entry.ReplacementStrings != null) {
                    foreach (var value in entry.ReplacementStrings) {
                        string match = MatchDiskIdentifier(value);
                        if (!string.IsNullOrEmpty(match)) {
                            return match;
                        }
                    }
                }

                // Andernfalls die vollständige Meldung durchsuchen.
                string fromMessage = MatchDiskIdentifier(entry.Message);
                if (!string.IsNullOrEmpty(fromMessage)) {
                    return fromMessage;
                }
            } catch {
                // Bei nicht lesbaren Einträgen keinen Bezeichner liefern.
            }

            return null;
        }

        private static string MatchDiskIdentifier(string text) {
            if (string.IsNullOrEmpty(text)) {
                return null;
            }

            // Gängige Formen: "\Device\Harddisk0\DR0", "Harddisk1", "PhysicalDrive2", "\Device\HarddiskVolume3".
            var match = System.Text.RegularExpressions.Regex.Match(
                text,
                @"(?:\\Device\\)?Harddisk(?:Volume)?\d+(?:\\DR\d+)?|PhysicalDrive\d+",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            return match.Success ? match.Value : null;
        }
    }
}
