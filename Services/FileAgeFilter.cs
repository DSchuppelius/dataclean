using System;

namespace DataClean.Services {
    /// <summary>
    /// Entscheidet anhand des Mindestalters, ob eine Datei beim Löschen übersprungen werden soll.
    /// Eine Datei gilt als "zu jung", wenn sowohl Schreib- als auch Zugriffszeit neuer als die
    /// Schwelle sind.
    /// </summary>
    public static class FileAgeFilter {
        public static bool ShouldSkip(DateTime lastWriteTime, DateTime lastAccessTime, int minFileAgeDays, DateTime now) {
            if (minFileAgeDays <= 0) {
                return false;
            }

            DateTime threshold = now.AddDays(-minFileAgeDays);
            return lastWriteTime > threshold && lastAccessTime > threshold;
        }
    }
}
