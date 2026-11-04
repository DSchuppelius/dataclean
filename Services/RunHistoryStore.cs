using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace DataClean.Services {
    /// <summary>
    /// Ein einzelner Eintrag der Lauf-Historie – fasst das Ergebnis eines
    /// Bereinigungslaufs zusammen.
    /// </summary>
    [DataContract]
    public class RunHistoryEntry {
        [DataMember(Order = 0)] public DateTime StartTime { get; set; }
        [DataMember(Order = 1)] public DateTime EndTime { get; set; }
        [DataMember(Order = 2)] public long FreedBytes { get; set; }
        [DataMember(Order = 3)] public int DeletedItems { get; set; }
        [DataMember(Order = 4)] public int SkippedItems { get; set; }
        [DataMember(Order = 5)] public int FailedItems { get; set; }
        [DataMember(Order = 6)] public bool DryRun { get; set; }
        [DataMember(Order = 7)] public bool Unattended { get; set; }
        [DataMember(Order = 8)] public bool Canceled { get; set; }
        [DataMember(Order = 9)] public string Tasks { get; set; }
    }

    [CollectionDataContract(Name = "History", ItemName = "Run")]
    public class RunHistoryList : List<RunHistoryEntry> {
        public RunHistoryList() { }
        public RunHistoryList(IEnumerable<RunHistoryEntry> items) : base(items) { }
    }

    /// <summary>
    /// Persistiert die Historie abgeschlossener Bereinigungsläufe als JSON unter
    /// %APPDATA%\DataClean\history.json. Hält maximal <see cref="MaxEntries"/> Einträge.
    /// Alle Methoden sind defensiv und werfen keine Ausnahmen nach außen.
    /// </summary>
    public static class RunHistoryStore {
        public const int MaxEntries = 100;

        public static string StoreDirectory { get; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "DataClean");

        private static string StoreFile => Path.Combine(StoreDirectory, "history.json");

        /// <summary>
        /// Lädt alle Historieneinträge, neueste zuerst.
        /// </summary>
        public static List<RunHistoryEntry> LoadAll() {
            try {
                if (!File.Exists(StoreFile)) {
                    return new List<RunHistoryEntry>();
                }

                using (var stream = File.OpenRead(StoreFile)) {
                    var serializer = new DataContractJsonSerializer(typeof(RunHistoryList));
                    var list = (RunHistoryList)serializer.ReadObject(stream);
                    if (list == null) {
                        return new List<RunHistoryEntry>();
                    }
                    return list.OrderByDescending(e => e.StartTime).ToList();
                }
            } catch (Exception ex) {
                Logger.Error(ex, "RunHistoryStore.LoadAll");
                return new List<RunHistoryEntry>();
            }
        }

        /// <summary>
        /// Fügt einen Eintrag hinzu und begrenzt die Gesamtzahl auf <see cref="MaxEntries"/>.
        /// </summary>
        public static bool Append(RunHistoryEntry entry) {
            if (entry == null) {
                return false;
            }

            try {
                var all = LoadAll();
                all.Insert(0, entry);
                if (all.Count > MaxEntries) {
                    all = all.Take(MaxEntries).ToList();
                }
                return Persist(all);
            } catch (Exception ex) {
                Logger.Error(ex, "RunHistoryStore.Append");
                return false;
            }
        }

        /// <summary>
        /// Löscht die gesamte Historie.
        /// </summary>
        public static bool Clear() {
            try {
                if (File.Exists(StoreFile)) {
                    File.Delete(StoreFile);
                }
                return true;
            } catch (Exception ex) {
                Logger.Error(ex, "RunHistoryStore.Clear");
                return false;
            }
        }

        private static bool Persist(List<RunHistoryEntry> entries) {
            try {
                Directory.CreateDirectory(StoreDirectory);
                using (var stream = File.Create(StoreFile)) {
                    var serializer = new DataContractJsonSerializer(typeof(RunHistoryList));
                    serializer.WriteObject(stream, new RunHistoryList(entries));
                }
                return true;
            } catch (Exception ex) {
                Logger.Error(ex, "RunHistoryStore.Persist");
                return false;
            }
        }
    }
}
