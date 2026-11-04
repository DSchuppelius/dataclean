namespace DataClean {
    /// <summary>
    /// Bündelt die aus den Anwendungseinstellungen geladenen Bereinigungs- und Report-Optionen.
    /// Reduziert die Anzahl einzelner Felder in der Form und bietet einen zentralen Mapping-Punkt.
    /// </summary>
    public class CleanupOptions {
        public bool EmptyTrash { get; set; }
        public int MinFileAgeDays { get; set; }
        public bool RemoveWindowsOld { get; set; }
        public bool FlushDnsCache { get; set; }
        public bool ClearEventLogs { get; set; }
        public bool HealthCheckEnabled { get; set; }
        public int MinFreeDiskPercent { get; set; }

        public bool DryRun { get; set; }
        public bool CreateRestorePoint { get; set; }
        public bool DeleteToRecycleBin { get; set; }
        public bool CleanBrowserCaches { get; set; }
        public bool DismComponentCleanup { get; set; }
        public bool CleanMinidumps { get; set; }
        public bool CleanPrefetch { get; set; }
        public bool CleanDeliveryOptimization { get; set; }
        public bool CleanOldLogs { get; set; }
        public int OldLogAgeDays { get; set; }
        public bool EmptyAllRecycleBins { get; set; }
        public bool ConfirmBeforeDelete { get; set; }
        public bool GenerateReport { get; set; }
        public string ReportPath { get; set; }
        public string ReportFormat { get; set; }
        public bool EmailReport { get; set; }

        /// <summary>
        /// Zeilengetrennte Glob-Ausschlussmuster (z. B. *.keep, **\\wichtig\\*).
        /// </summary>
        public string ExcludePatterns { get; set; }

        /// <summary>
        /// Überspringt den Lauf, wenn auf dem Systemlaufwerk noch mehr als
        /// MinFreeDiskPercent freier Speicher vorhanden ist (Schwellenwert-Trigger).
        /// </summary>
        public bool SkipIfEnoughFreeSpace { get; set; }

        /// <summary>
        /// Markiert gesperrte Dateien zum Löschen beim nächsten Systemneustart (MoveFileEx).
        /// </summary>
        public bool SchedulePendingDeletes { get; set; }

        /// <summary>
        /// Lädt die Optionen aus den persistierten Anwendungseinstellungen.
        /// </summary>
        public static CleanupOptions FromSettings() {
            var s = Properties.Settings.Default;
            return new CleanupOptions {
                EmptyTrash = s.emptyTrash,
                MinFileAgeDays = s.minFileAgeDays,
                RemoveWindowsOld = s.removeWindowsOld,
                FlushDnsCache = s.flushDnsCache,
                ClearEventLogs = s.clearEventLogs,
                HealthCheckEnabled = s.healthCheckEnabled,
                MinFreeDiskPercent = s.minFreeDiskPercent,
                DryRun = s.dryRun,
                CreateRestorePoint = s.createRestorePoint,
                DeleteToRecycleBin = s.deleteToRecycleBin,
                CleanBrowserCaches = s.cleanBrowserCaches,
                DismComponentCleanup = s.dismComponentCleanup,
                CleanMinidumps = s.cleanMinidumps,
                CleanPrefetch = s.cleanPrefetch,
                CleanDeliveryOptimization = s.cleanDeliveryOptimization,
                CleanOldLogs = s.cleanOldLogs,
                OldLogAgeDays = s.oldLogAgeDays,
                EmptyAllRecycleBins = s.emptyAllRecycleBins,
                ConfirmBeforeDelete = s.confirmBeforeDelete,
                GenerateReport = s.generateReport,
                ReportPath = s.reportPath,
                ReportFormat = s.reportFormat,
                EmailReport = s.emailReport,
                ExcludePatterns = s.excludePatterns,
                SkipIfEnoughFreeSpace = s.skipIfEnoughFreeSpace,
                SchedulePendingDeletes = s.schedulePendingDeletes
            };
        }

        /// <summary>
        /// Persistiert die editierbaren (benutzerbezogenen) Optionen zurück in die Anwendungseinstellungen.
        /// Anwendungsweite Einstellungen (z. B. EmptyTrash, HealthCheckEnabled) sind zur Laufzeit
        /// schreibgeschützt und werden daher hier bewusst nicht geschrieben.
        /// </summary>
        public void Save() {
            var s = Properties.Settings.Default;

            s.dryRun = DryRun;
            s.createRestorePoint = CreateRestorePoint;
            s.deleteToRecycleBin = DeleteToRecycleBin;
            s.cleanBrowserCaches = CleanBrowserCaches;
            s.dismComponentCleanup = DismComponentCleanup;
            s.cleanMinidumps = CleanMinidumps;
            s.cleanPrefetch = CleanPrefetch;
            s.cleanDeliveryOptimization = CleanDeliveryOptimization;
            s.cleanOldLogs = CleanOldLogs;
            s.oldLogAgeDays = OldLogAgeDays;
            s.emptyAllRecycleBins = EmptyAllRecycleBins;
            s.confirmBeforeDelete = ConfirmBeforeDelete;
            s.generateReport = GenerateReport;
            s.reportPath = ReportPath ?? string.Empty;
            s.reportFormat = string.IsNullOrWhiteSpace(ReportFormat) ? "HTML" : ReportFormat;
            s.emailReport = EmailReport;
            s.excludePatterns = ExcludePatterns ?? string.Empty;
            s.skipIfEnoughFreeSpace = SkipIfEnoughFreeSpace;
            s.schedulePendingDeletes = SchedulePendingDeletes;

            s.Save();
        }
    }
}
