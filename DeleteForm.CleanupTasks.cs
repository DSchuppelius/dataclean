using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace DataClean {
    public partial class DeleteForm {
        // Löscht Dateien/Ordner eines Pfads über den bestehenden Lösch-Hotpath,
        // sodass Dry-Run, Papierkorb-Option, Altersfilter und Statistik konsistent greifen.
        private async Task CleanPathAsync(string path, string taskLabel) {
            if (string.IsNullOrWhiteSpace(path)) {
                return;
            }

            var dir = new DirectoryInfo(Environment.ExpandEnvironmentVariables(path));
            if (!dir.Exists) {
                LogToDetailTextBox($"{taskLabel}: {dir.FullName} nicht vorhanden – übersprungen.");
                return;
            }

            Invoke(new Action(() => {
                UpdateStatusLabel(label1, $"{taskLabel}...", themeText);
                UpdateStatusLabel(directoryValueLabel, dir.FullName);
            }));

            await ProcessDirectory(dir);
        }

        private async Task CleanBrowserCachesAsync() {
            performedTasks.Add("Browser-Caches");

            string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string roaming = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

            var cachePaths = new List<string> {
                Path.Combine(local, @"Microsoft\Edge\User Data\Default\Cache"),
                Path.Combine(local, @"Microsoft\Edge\User Data\Default\Code Cache"),
                Path.Combine(local, @"Google\Chrome\User Data\Default\Cache"),
                Path.Combine(local, @"Google\Chrome\User Data\Default\Code Cache"),
                Path.Combine(local, @"BraveSoftware\Brave-Browser\User Data\Default\Cache"),
            };

            // Firefox-Profile liegen unter %LOCALAPPDATA%\Mozilla\Firefox\Profiles\<profil>\cache2
            string firefoxProfiles = Path.Combine(local, @"Mozilla\Firefox\Profiles");
            if (Directory.Exists(firefoxProfiles)) {
                foreach (var profile in Directory.EnumerateDirectories(firefoxProfiles)) {
                    cachePaths.Add(Path.Combine(profile, "cache2"));
                }
            }

            // Edge/Chrome speichern teils auch unter Roaming – defensiv ergänzen.
            cachePaths.Add(Path.Combine(roaming, @"Opera Software\Opera Stable\Cache"));

            foreach (var cachePath in cachePaths) {
                await CleanPathAsync(cachePath, "Browser-Cache");
            }
        }

        private async Task CleanMinidumpsAsync() {
            performedTasks.Add("Absturz-Dumps");

            string windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);

            // Minidump-Verzeichnis
            await CleanPathAsync(Path.Combine(windir, "Minidump"), "Minidumps");

            // Große MEMORY.DMP separat über den Hotpath löschen
            string memoryDump = Path.Combine(windir, "MEMORY.DMP");
            var memFile = new FileInfo(memoryDump);
            if (memFile.Exists) {
                Invoke(new Action(() => UpdateStatusLabel(label1, "Entferne MEMORY.DMP...", themeText)));
                long size = GetFileLength(memFile);
                if (await TryDeleteItemAsync(memFile)) {
                    System.Threading.Interlocked.Add(ref freedBytes, size);
                    LogToDetailTextBox($"MEMORY.DMP entfernt: {FormatBytes(size)}");
                }
            }

            // Benutzerbezogene Absturz-Dumps
            string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            await CleanPathAsync(Path.Combine(local, @"CrashDumps"), "Absturz-Dumps");
        }

        private async Task CleanPrefetchAsync() {
            performedTasks.Add("Prefetch");
            string windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            await CleanPathAsync(Path.Combine(windir, "Prefetch"), "Prefetch");
        }

        private async Task CleanDeliveryOptimizationAsync() {
            performedTasks.Add("Delivery Optimization");
            string windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            await CleanPathAsync(
                Path.Combine(windir, @"SoftwareDistribution\DeliveryOptimization"),
                "Delivery Optimization");
        }

        private async Task CleanOldLogsAsync() {
            performedTasks.Add($"Alte Protokolle (> {oldLogAgeDays} Tage)");

            string windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            var logRoots = new[] {
                Path.Combine(windir, "Logs"),
                Path.Combine(windir, @"System32\LogFiles"),
                Path.Combine(windir, "Temp"),
            };

            DateTime threshold = DateTime.Now.AddDays(-Math.Max(0, oldLogAgeDays));
            var extensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".log", ".etl", ".old", ".dmp" };

            foreach (var root in logRoots) {
                if (!Directory.Exists(root)) {
                    continue;
                }

                Invoke(new Action(() => {
                    UpdateStatusLabel(label1, "Entferne alte Protokolle...", themeText);
                    UpdateStatusLabel(directoryValueLabel, root);
                }));

                List<FileInfo> oldFiles;
                try {
                    oldFiles = new DirectoryInfo(root)
                        .EnumerateFiles("*", SearchOption.AllDirectories)
                        .Where(f => {
                            try {
                                return extensions.Contains(f.Extension) && f.LastWriteTime < threshold;
                            } catch {
                                return false;
                            }
                        })
                        .ToList();
                } catch (Exception ex) when (ex is UnauthorizedAccessException || ex is IOException) {
                    LogToDetailTextBox($"Protokollordner {root} nicht vollständig lesbar: {ex.Message}");
                    continue;
                }

                foreach (var file in oldFiles) {
                    long size = GetFileLength(file);
                    if (await TryDeleteItemAsync(file)) {
                        System.Threading.Interlocked.Add(ref freedBytes, size);
                        LogToDetailTextBox($"Altes Protokoll entfernt: {file.FullName}");
                    }
                }
            }
        }

        private async Task RunDismComponentCleanupAsync() {
            performedTasks.Add("DISM Komponentenbereinigung");

            if (!IsAdministrator()) {
                LogToDetailTextBox("DISM Komponentenbereinigung übersprungen – Administratorrechte erforderlich.");
                return;
            }

            Invoke(new Action(() => UpdateStatusLabel(label1, "DISM Komponentenbereinigung läuft (kann dauern)...", themeText)));

            try {
                var psi = new System.Diagnostics.ProcessStartInfo("dism.exe",
                    "/Online /Cleanup-Image /StartComponentCleanup") {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                await Task.Run(() => {
                    using (var process = System.Diagnostics.Process.Start(psi)) {
                        string stderr = process.StandardError.ReadToEnd();
                        process.WaitForExit(600000); // bis zu 10 Minuten

                        LogToDetailTextBox(process.ExitCode == 0
                            ? "DISM Komponentenbereinigung erfolgreich abgeschlossen."
                            : $"DISM Komponentenbereinigung mit Code {process.ExitCode} beendet: {stderr.Trim()}");
                    }
                });
            } catch (Exception ex) {
                LogToDetailTextBox($"DISM Komponentenbereinigung fehlgeschlagen: {ex.Message}");
            }
        }

        // Leert den Papierkorb aller Laufwerke einzeln (statt nur global).
        private void EmptyAllRecycleBins() {
            performedTasks.Add("Papierkorb (alle Laufwerke)");

            if (dryRun) {
                LogToDetailTextBox("[Vorschau] Würde Papierkorb aller Laufwerke leeren.");
                return;
            }

            Invoke(new Action(() => UpdateStatusLabel(label1, "Leere Papierkorb aller Laufwerke...", themeText)));

            foreach (var drive in DriveInfo.GetDrives()) {
                if (!drive.IsReady) {
                    continue;
                }

                try {
                    int result = SHEmptyRecycleBin(IntPtr.Zero, drive.Name,
                        SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND);

                    // S_OK (0) oder "bereits leer" (unchecked 0x8000FFFF) als Erfolg werten.
                    if (result == 0 || result == unchecked((int)0x8000FFFF)) {
                        LogToDetailTextBox($"Papierkorb auf {drive.Name} geleert.");
                    } else {
                        LogToDetailTextBox($"Papierkorb auf {drive.Name} konnte nicht geleert werden (Code {result}).");
                    }
                } catch (Exception ex) {
                    LogToDetailTextBox($"Papierkorb auf {drive.Name} fehlgeschlagen: {ex.Message}");
                }
            }
        }
    }
}
