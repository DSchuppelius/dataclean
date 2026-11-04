using System;
using System.IO;
using System.Linq;

namespace DataClean.Services {
    /// <summary>
    /// Verfügbare Protokollierungsstufen in aufsteigender Dringlichkeit.
    /// </summary>
    public enum LogLevel {
        Info = 0,
        Warn = 1,
        Error = 2,
        /// <summary>Nichts protokollieren, z. B. für absichtlich provozierte Fehler in Tests.</summary>
        None = 3
    }

    /// <summary>
    /// Einfaches, threadsicheres Datei-Logging. Schreibt in eine Tagesdatei unter
    /// %LOCALAPPDATA%\DataClean\Logs. Alle Methoden sind defensiv und werfen nie Ausnahmen
    /// nach außen – fehlgeschlagenes Logging darf die Anwendung nicht beeinträchtigen
    /// (wichtig für den unbeaufsichtigten Autostart).
    /// Unterstützt eine minimale Protokollierungsstufe sowie eine größenbasierte
    /// Dateirotation mit begrenzter Anzahl archivierter Dateien.
    /// </summary>
    public static class Logger {
        private static readonly object SyncRoot = new object();

        /// <summary>
        /// Maximale Größe einer Protokolldatei in Bytes, bevor rotiert wird (Standard 2 MB).
        /// </summary>
        public static long MaxFileSizeBytes { get; set; } = 2 * 1024 * 1024;

        /// <summary>
        /// Anzahl archivierter (rotierter) Dateien, die aufbewahrt werden.
        /// </summary>
        public static int MaxArchiveFiles { get; set; } = 5;

        /// <summary>
        /// Nachrichten unterhalb dieser Stufe werden verworfen.
        /// </summary>
        public static LogLevel MinimumLevel { get; set; } = LogLevel.Info;

        /// <summary>
        /// Verzeichnis, in das die Protokolldateien geschrieben werden.
        /// </summary>
        public static string LogDirectory { get; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DataClean", "Logs");

        private static string CurrentLogFile =>
            Path.Combine(LogDirectory, $"dataclean-{DateTime.Now:yyyyMMdd}.log");

        public static void Info(string message, string category = null) {
            Write(LogLevel.Info, "INFO", category, message);
        }

        public static void Warn(string message, string category = null) {
            Write(LogLevel.Warn, "WARN", category, message);
        }

        public static void Error(string message, string category = null) {
            Write(LogLevel.Error, "ERROR", category, message);
        }

        public static void Error(Exception ex, string category = null) {
            if (ex == null) {
                return;
            }
            Write(LogLevel.Error, "ERROR", category, $"{ex.GetType().Name}: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
        }

        private static void Write(LogLevel level, string levelText, string category, string message) {
            if (level < MinimumLevel) {
                return;
            }

            try {
                Directory.CreateDirectory(LogDirectory);

                string prefix = string.IsNullOrEmpty(category) ? levelText : $"{levelText}/{category}";
                string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{prefix}] {message}";

                lock (SyncRoot) {
                    RotateIfNeeded();
                    File.AppendAllText(CurrentLogFile, line + Environment.NewLine);
                }

                System.Diagnostics.Debug.WriteLine(line);
            } catch {
                // Logging darf niemals die Anwendung stören.
            }
        }

        // Rotiert die aktuelle Datei, wenn sie die maximale Größe überschreitet, und
        // begrenzt die Anzahl der archivierten Dateien. Aufrufer hält SyncRoot.
        private static void RotateIfNeeded() {
            try {
                string current = CurrentLogFile;
                if (!File.Exists(current)) {
                    return;
                }

                var info = new FileInfo(current);
                if (info.Length < MaxFileSizeBytes) {
                    return;
                }

                string archiveName = Path.Combine(LogDirectory,
                    $"dataclean-{DateTime.Now:yyyyMMdd}-{DateTime.Now:HHmmssfff}.log");
                File.Move(current, archiveName);

                PruneArchives();
            } catch {
                // Rotation darf das Logging nicht blockieren.
            }
        }

        private static void PruneArchives() {
            try {
                var archives = Directory.GetFiles(LogDirectory, "dataclean-*-*.log")
                    .Select(f => new FileInfo(f))
                    .OrderByDescending(f => f.LastWriteTimeUtc)
                    .Skip(Math.Max(0, MaxArchiveFiles))
                    .ToList();

                foreach (var old in archives) {
                    try {
                        old.Delete();
                    } catch {
                        // Einzelne nicht löschbare Datei überspringen.
                    }
                }
            } catch {
                // Aufräumen ist best-effort.
            }
        }
    }
}

