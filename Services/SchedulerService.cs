using System;
using System.Diagnostics;
using System.Text;

namespace DataClean.Services {
    /// <summary>
    /// Integriert DataClean in die Windows-Aufgabenplanung über das Kommandozeilen-Werkzeug
    /// <c>schtasks.exe</c>. Ermöglicht das Anlegen, Entfernen und Abfragen einer geplanten
    /// täglichen Bereinigungsaufgabe. Alle Operationen erfordern in der Regel Administratorrechte.
    /// </summary>
    public static class SchedulerService {
        public const string TaskName = "DataClean Automatische Bereinigung";

        /// <summary>
        /// Legt eine geplante Aufgabe an, die DataClean täglich zur angegebenen Uhrzeit ausführt.
        /// Eine bestehende Aufgabe gleichen Namens wird überschrieben.
        /// </summary>
        public static bool CreateDailyTask(TimeSpan timeOfDay, out string message) {
            string exePath = Process.GetCurrentProcess().MainModule.FileName;
            string time = string.Format("{0:D2}:{1:D2}", timeOfDay.Hours, timeOfDay.Minutes);

            string arguments =
                $"/Create /TN \"{TaskName}\" /TR \"\\\"{exePath}\\\" /autostart\" /SC DAILY /ST {time} /RL HIGHEST /F";

            return RunSchtasks(arguments, out message);
        }

        /// <summary>
        /// Entfernt die geplante Bereinigungsaufgabe, sofern vorhanden.
        /// </summary>
        public static bool DeleteTask(out string message) {
            return RunSchtasks($"/Delete /TN \"{TaskName}\" /F", out message);
        }

        /// <summary>
        /// Prüft, ob die geplante Bereinigungsaufgabe existiert.
        /// </summary>
        public static bool TaskExists() {
            string message;
            return RunSchtasks($"/Query /TN \"{TaskName}\"", out message);
        }

        private static bool RunSchtasks(string arguments, out string message) {
            try {
                var psi = new ProcessStartInfo("schtasks.exe", arguments) {
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };

                using (var process = Process.Start(psi)) {
                    string stdout = process.StandardOutput.ReadToEnd();
                    string stderr = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    message = string.IsNullOrWhiteSpace(stderr) ? stdout.Trim() : stderr.Trim();
                    bool success = process.ExitCode == 0;
                    if (!success) {
                        Logger.Warn($"schtasks beendet mit Code {process.ExitCode}: {message}", "SchedulerService");
                    }
                    return success;
                }
            } catch (Exception ex) {
                Logger.Error(ex, "SchedulerService");
                message = ex.Message;
                return false;
            }
        }
    }
}
