using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace DataClean {
    public partial class DeleteForm {
        // Löscht Dateien/Ordner eines Pfads über den bestehenden Lösch-Hotpath,
        // sodass Dry-Run, Papierkorb-Option, Altersfilter und Statistik konsistent greifen.
        private async Task CleanPathAsync(string path, string taskLabel, CancellationToken cancellationToken) {
            if (string.IsNullOrWhiteSpace(path)) {
                return;
            }

            cancellationToken.ThrowIfCancellationRequested();

            var dir = new DirectoryInfo(Environment.ExpandEnvironmentVariables(path));
            if (!dir.Exists) {
                LogToDetailTextBox($"{taskLabel}: {dir.FullName} nicht vorhanden – übersprungen.");
                return;
            }

            Invoke(new Action(() => {
                UpdateStatusLabel(label1, $"{taskLabel}...", themeStatusText);
                UpdateStatusLabel(directoryValueLabel, dir.FullName);
            }));

            await ProcessDirectory(dir, cancellationToken);
        }

        private async Task CleanBrowserCachesAsync(CancellationToken cancellationToken) {
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
                await CleanPathAsync(cachePath, "Browser-Cache", cancellationToken);
            }
        }

        private async Task CleanMinidumpsAsync(CancellationToken cancellationToken) {
            performedTasks.Add("Absturz-Dumps");

            string windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);

            // Minidump-Verzeichnis
            await CleanPathAsync(Path.Combine(windir, "Minidump"), "Minidumps", cancellationToken);

            // Große MEMORY.DMP separat über den Hotpath löschen
            string memoryDump = Path.Combine(windir, "MEMORY.DMP");
            var memFile = new FileInfo(memoryDump);
            if (memFile.Exists) {
                Invoke(new Action(() => UpdateStatusLabel(label1, "Entferne MEMORY.DMP...", themeStatusText)));
                long size = GetFileLength(memFile);
                if (await TryDeleteItemAsync(memFile)) {
                    System.Threading.Interlocked.Add(ref freedBytes, size);
                    if (!dryRun) {
                        LogToDetailTextBox($"MEMORY.DMP entfernt: {FormatBytes(size)}");
                    }
                }
            }

            // Benutzerbezogene Absturz-Dumps
            string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            await CleanPathAsync(Path.Combine(local, @"CrashDumps"), "Absturz-Dumps", cancellationToken);
        }

        private async Task CleanPrefetchAsync(CancellationToken cancellationToken) {
            performedTasks.Add("Prefetch");
            string windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            await CleanPathAsync(Path.Combine(windir, "Prefetch"), "Prefetch", cancellationToken);
        }

        private async Task CleanDeliveryOptimizationAsync(CancellationToken cancellationToken) {
            performedTasks.Add("Delivery Optimization");
            string windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            await CleanPathAsync(
                Path.Combine(windir, @"SoftwareDistribution\DeliveryOptimization"),
                "Delivery Optimization", cancellationToken);
        }

        private async Task CleanOldLogsAsync(CancellationToken cancellationToken) {
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
                    UpdateStatusLabel(label1, "Entferne alte Protokolle...", themeStatusText);
                    UpdateStatusLabel(directoryValueLabel, root);
                }));

                // Streaming-Enumeration: Dateien werden lazy geliefert, sodass nicht alle
                // FileInfo-Objekte gleichzeitig im Speicher gehalten werden. Unzugängliche
                // Unterordner werden einzeln übersprungen, statt den gesamten Root abzubrechen.
                foreach (var file in WalkDirectories(new DirectoryInfo(root)).SelectMany(Services.DirectoryWalker.EnumerateFiles)) {
                    await WaitIfPausedAsync(cancellationToken);
                    cancellationToken.ThrowIfCancellationRequested();

                    bool matches;
                    try {
                        matches = extensions.Contains(file.Extension) && file.LastWriteTime < threshold;
                    } catch {
                        matches = false;
                    }

                    if (!matches) {
                        continue;
                    }

                    long size = GetFileLength(file);
                    if (await TryDeleteItemAsync(file)) {
                        System.Threading.Interlocked.Add(ref freedBytes, size);
                        if (!dryRun) {
                            LogToDetailTextBox($"Altes Protokoll entfernt: {file.FullName}");
                        }
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

            await RunDismProcessAsync("dism.exe", "/Online /Cleanup-Image /StartComponentCleanup");
        }

        // Letzte Ausgabe von DISM. Fehler schreibt DISM in die normale Ausgabe, nicht in die Fehlerausgabe.
        private readonly Queue<string> dismOutputTail = new Queue<string>();
        private const int DismOutputTailLines = 6;
        private double lastDismPercent = -1d;

        private async Task RunDismProcessAsync(string fileName, string arguments) {
            BeginBusyStep("DISM Komponentenbereinigung läuft (kann dauern)", "Windows-Komponentenspeicher");
            try {
                Services.ProcessResult result = await Task.Run(() =>
                    Services.ProcessRunner.Run(fileName, arguments, OnDismOutput, DismTimeout));
                ReportDismResult(result);
            } catch (Exception ex) {
                ReportDismProblem($"DISM Komponentenbereinigung fehlgeschlagen: {ex.Message}");
            } finally {
                EndBusyStep();
            }
        }

        private void ReportDismResult(Services.ProcessResult result) {
            string progress = lastDismPercent >= 0d ? $" Letzter Fortschritt: {lastDismPercent:0} %." : string.Empty;

            if (result.TimedOut) {
                ReportDismProblem($"DISM Komponentenbereinigung nach {DismTimeout.TotalMinutes:0} Minuten noch nicht beendet und läuft im Hintergrund weiter.{progress}");
            } else if (result.ExitCode == 0) {
                LogToDetailTextBox("DISM Komponentenbereinigung erfolgreich abgeschlossen.");
            } else {
                string detail = result.ErrorOutput;
                if (string.IsNullOrEmpty(detail)) {
                    lock (dismOutputTail) {
                        detail = string.Join(" ", dismOutputTail);
                    }
                }
                ReportDismProblem($"DISM Komponentenbereinigung mit Code {result.ExitCode} (0x{result.ExitCode:X8}) beendet.{progress} {detail}".TrimEnd());
            }
        }

        // Das Fenster schließt sich meist von selbst; das Log bleibt als Nachweis erhalten.
        private void ReportDismProblem(string message) {
            LogToDetailTextBox(message);
            Services.Logger.Warn(message, "DISM");
        }

        private void OnDismOutput(string line) {
            double percent;
            if (Services.DismProgressParser.TryParsePercent(line, out percent)) {
                lastDismPercent = percent;
                ReportBusyProgress(percent);
            } else if (!string.IsNullOrWhiteSpace(line)) {
                lock (dismOutputTail) {
                    dismOutputTail.Enqueue(line.Trim());
                    while (dismOutputTail.Count > DismOutputTailLines) {
                        dismOutputTail.Dequeue();
                    }
                }
            }
        }

        // Leert den Papierkorb aller Laufwerke einzeln (statt nur global).
        private void EmptyAllRecycleBins() {
            performedTasks.Add("Papierkorb (alle Laufwerke)");

            if (dryRun) {
                LogToDetailTextBox("[Vorschau] Würde Papierkorb aller Laufwerke leeren.");
                return;
            }

            BeginBusyStep("Leere Papierkorb aller Laufwerke", "Papierkorb");

            foreach (var drive in DriveInfo.GetDrives()) {
                if (!drive.IsReady) {
                    continue;
                }

                try {
                    int result = NativeMethods.SHEmptyRecycleBin(IntPtr.Zero, drive.Name,
                        NativeMethods.SHERB_NOCONFIRMATION | NativeMethods.SHERB_NOPROGRESSUI | NativeMethods.SHERB_NOSOUND);

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

            EndBusyStep();
        }

        // Durchläuft einen Ordnerbaum und zählt nicht lesbare Ordner als übersprungen.
        private IEnumerable<DirectoryInfo> WalkDirectories(DirectoryInfo root) {
            return Services.DirectoryWalker.Walk(root, unreadable => {
                System.Threading.Interlocked.Increment(ref skippedItemCount);
                LogToDetailTextBox($"Übersprungen (kein Zugriff): {unreadable.FullName}");
                RecordActivity(ActivityStatus.Skipped, unreadable.FullName, 0);
            });
        }
    }
}
