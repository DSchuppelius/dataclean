using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace DataClean.Services {
    /// <summary>
    /// Ein benanntes Bereinigungsprofil – eine gespeicherte Momentaufnahme der
    /// editierbaren Bereinigungs- und Report-Optionen.
    /// </summary>
    [DataContract]
    public class CleanupProfile {
        [DataMember(Order = 0)] public string Name { get; set; }
        [DataMember(Order = 1)] public bool DryRun { get; set; }
        [DataMember(Order = 2)] public bool CreateRestorePoint { get; set; }
        [DataMember(Order = 3)] public bool DeleteToRecycleBin { get; set; }
        [DataMember(Order = 4)] public bool CleanBrowserCaches { get; set; }
        [DataMember(Order = 5)] public bool DismComponentCleanup { get; set; }
        [DataMember(Order = 6)] public bool CleanMinidumps { get; set; }
        [DataMember(Order = 7)] public bool CleanPrefetch { get; set; }
        [DataMember(Order = 8)] public bool CleanDeliveryOptimization { get; set; }
        [DataMember(Order = 9)] public bool CleanOldLogs { get; set; }
        [DataMember(Order = 10)] public int OldLogAgeDays { get; set; }
        [DataMember(Order = 11)] public bool EmptyAllRecycleBins { get; set; }
        [DataMember(Order = 12)] public bool ConfirmBeforeDelete { get; set; }
        [DataMember(Order = 13)] public bool GenerateReport { get; set; }
        [DataMember(Order = 14)] public string ReportPath { get; set; }
        [DataMember(Order = 15)] public string ReportFormat { get; set; }
        [DataMember(Order = 16)] public bool EmailReport { get; set; }
        [DataMember(Order = 17)] public string ExcludePatterns { get; set; }
        [DataMember(Order = 18)] public bool SkipIfEnoughFreeSpace { get; set; }
        [DataMember(Order = 19)] public bool SchedulePendingDeletes { get; set; }

        /// <summary>
        /// Erzeugt ein Profil aus den übergebenen Optionen.
        /// </summary>
        public static CleanupProfile FromOptions(string name, CleanupOptions o) {
            return new CleanupProfile {
                Name = name,
                DryRun = o.DryRun,
                CreateRestorePoint = o.CreateRestorePoint,
                DeleteToRecycleBin = o.DeleteToRecycleBin,
                CleanBrowserCaches = o.CleanBrowserCaches,
                DismComponentCleanup = o.DismComponentCleanup,
                CleanMinidumps = o.CleanMinidumps,
                CleanPrefetch = o.CleanPrefetch,
                CleanDeliveryOptimization = o.CleanDeliveryOptimization,
                CleanOldLogs = o.CleanOldLogs,
                OldLogAgeDays = o.OldLogAgeDays,
                EmptyAllRecycleBins = o.EmptyAllRecycleBins,
                ConfirmBeforeDelete = o.ConfirmBeforeDelete,
                GenerateReport = o.GenerateReport,
                ReportPath = o.ReportPath,
                ReportFormat = o.ReportFormat,
                EmailReport = o.EmailReport,
                ExcludePatterns = o.ExcludePatterns,
                SkipIfEnoughFreeSpace = o.SkipIfEnoughFreeSpace,
                SchedulePendingDeletes = o.SchedulePendingDeletes
            };
        }

        /// <summary>
        /// Überträgt die Profilwerte in ein CleanupOptions-Objekt.
        /// </summary>
        public void ApplyTo(CleanupOptions o) {
            o.DryRun = DryRun;
            o.CreateRestorePoint = CreateRestorePoint;
            o.DeleteToRecycleBin = DeleteToRecycleBin;
            o.CleanBrowserCaches = CleanBrowserCaches;
            o.DismComponentCleanup = DismComponentCleanup;
            o.CleanMinidumps = CleanMinidumps;
            o.CleanPrefetch = CleanPrefetch;
            o.CleanDeliveryOptimization = CleanDeliveryOptimization;
            o.CleanOldLogs = CleanOldLogs;
            o.OldLogAgeDays = OldLogAgeDays;
            o.EmptyAllRecycleBins = EmptyAllRecycleBins;
            o.ConfirmBeforeDelete = ConfirmBeforeDelete;
            o.GenerateReport = GenerateReport;
            o.ReportPath = ReportPath;
            o.ReportFormat = ReportFormat;
            o.EmailReport = EmailReport;
            o.ExcludePatterns = ExcludePatterns;
            o.SkipIfEnoughFreeSpace = SkipIfEnoughFreeSpace;
            o.SchedulePendingDeletes = SchedulePendingDeletes;
        }
    }

    [CollectionDataContract(Name = "Profiles", ItemName = "Profile")]
    public class CleanupProfileList : List<CleanupProfile> {
        public CleanupProfileList() { }
        public CleanupProfileList(IEnumerable<CleanupProfile> items) : base(items) { }
    }

    /// <summary>
    /// Persistiert benannte Bereinigungsprofile als JSON unter
    /// %APPDATA%\DataClean\profiles.json. Alle Methoden sind defensiv und werfen
    /// keine Ausnahmen nach außen.
    /// </summary>
    public static class CleanupProfileStore {
        public static string StoreDirectory { get; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "DataClean");

        private static string StoreFile => Path.Combine(StoreDirectory, "profiles.json");

        /// <summary>
        /// Lädt alle gespeicherten Profile. Gibt bei Fehlern eine leere Liste zurück.
        /// </summary>
        public static List<CleanupProfile> LoadAll() {
            try {
                if (!File.Exists(StoreFile)) {
                    return new List<CleanupProfile>();
                }

                using (var stream = File.OpenRead(StoreFile)) {
                    var serializer = new DataContractJsonSerializer(typeof(CleanupProfileList));
                    var list = (CleanupProfileList)serializer.ReadObject(stream);
                    return list != null ? new List<CleanupProfile>(list) : new List<CleanupProfile>();
                }
            } catch (Exception ex) {
                Logger.Error(ex, "CleanupProfileStore.LoadAll");
                return new List<CleanupProfile>();
            }
        }

        /// <summary>
        /// Gibt die Namen aller gespeicherten Profile zurück.
        /// </summary>
        public static List<string> GetNames() {
            return LoadAll().Select(p => p.Name).OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
        }

        /// <summary>
        /// Gibt das Profil mit dem angegebenen Namen zurück oder null.
        /// </summary>
        public static CleanupProfile Get(string name) {
            if (string.IsNullOrWhiteSpace(name)) {
                return null;
            }
            return LoadAll().FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Speichert (oder überschreibt) ein Profil anhand seines Namens.
        /// </summary>
        public static bool Save(CleanupProfile profile) {
            if (profile == null || string.IsNullOrWhiteSpace(profile.Name)) {
                return false;
            }

            try {
                var all = LoadAll();
                all.RemoveAll(p => string.Equals(p.Name, profile.Name, StringComparison.OrdinalIgnoreCase));
                all.Add(profile);
                return Persist(all);
            } catch (Exception ex) {
                Logger.Error(ex, "CleanupProfileStore.Save");
                return false;
            }
        }

        /// <summary>
        /// Löscht das Profil mit dem angegebenen Namen.
        /// </summary>
        public static bool Delete(string name) {
            if (string.IsNullOrWhiteSpace(name)) {
                return false;
            }

            try {
                var all = LoadAll();
                int removed = all.RemoveAll(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
                return removed > 0 && Persist(all);
            } catch (Exception ex) {
                Logger.Error(ex, "CleanupProfileStore.Delete");
                return false;
            }
        }

        private static bool Persist(List<CleanupProfile> profiles) {
            try {
                Directory.CreateDirectory(StoreDirectory);
                using (var stream = File.Create(StoreFile)) {
                    var serializer = new DataContractJsonSerializer(typeof(CleanupProfileList));
                    serializer.WriteObject(stream, new CleanupProfileList(profiles));
                }
                return true;
            } catch (Exception ex) {
                Logger.Error(ex, "CleanupProfileStore.Persist");
                return false;
            }
        }
    }
}
