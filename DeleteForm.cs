using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataClean {
    public partial class DeleteForm : UI.FuturisticForm {
        private readonly List<DirectoryInfo> protectedDirectories;
        private readonly List<DirectoryInfo> allowedSubdirectories;

        private List<DirectoryInfo> cleaningDirectories;

        private bool emptyTrash = false;
        private int minFileAgeDays = 0;
        private long freedBytes = 0;
        private bool removeWindowsOld = false;
        private bool flushDnsCache = false;
        private bool clearEventLogs = false;
        private bool healthCheckEnabled = false;
        private int minFreeDiskPercent = 5;
        private bool severeHealthIssueDetected = false;
        private Task healthCheckTask;
        private bool updateServicesStopped = false;
        private CancellationTokenSource cleanupCts;

        // Ergebnisse der Systemprüfung werden zwischengespeichert und erst am Ende
        // ausgegeben, damit sie nicht im Protokoll der gelöschten Dateien untergehen.
        private List<HealthIssue> pendingHealthIssues;

        // Sicherheits- und Erweiterungs-Optionen
        private bool dryRun = false;
        private bool createRestorePoint = false;
        private bool deleteToRecycleBin = false;
        private bool cleanBrowserCaches = false;
        private bool dismComponentCleanup = false;
        private bool cleanMinidumps = false;
        private bool cleanPrefetch = false;
        private bool cleanDeliveryOptimization = false;
        private bool cleanOldLogs = false;
        private int oldLogAgeDays = 30;
        private bool emptyAllRecycleBins = false;
        private bool confirmBeforeDelete = true;
        private bool generateReport = false;
        private string reportPath = string.Empty;
        private string reportFormat = "HTML";
        private bool emailReport = false;
        private bool reportWritten = false;
        private Services.ExcludeRuleMatcher excludeMatcher;
        private bool skipIfEnoughFreeSpace = false;
        private bool schedulePendingDeletes = false;

        // Phasenanzeige: Gesamtzahl der aktivierten Bereinigungs-Phasen und laufender Index.
        private int totalPhases = 0;
        private int currentPhase = 0;

        // Statistik für Report
        private int deletedItemCount = 0;
        private int skippedItemCount = 0;
        private int failedItemCount = 0;
        private readonly DateTime runStartTime = DateTime.Now;
        private readonly List<string> performedTasks = new List<string>();

        // Drosselung der UI-Updates im Lösch-Hotpath, um blockierende Invoke-Aufrufe
        // pro Element zu vermeiden (Hauptursache für langsames Löschen).
        private readonly System.Diagnostics.Stopwatch uiThrottle = System.Diagnostics.Stopwatch.StartNew();
        private const int UiUpdateIntervalMs = 100;
        private int pendingProgressSteps = 0;
        private readonly System.Text.StringBuilder detailLogBuffer = new System.Text.StringBuilder();
        private readonly object detailLogLock = new object();

        // Wird beim unbeaufsichtigten Autostart (geplante Aufgabe) gesetzt und überspringt
        // interaktive Rückfragen wie die Größen-Vorschau-Bestätigung.
        private readonly bool unattended;

        public DeleteForm() : this(false) {
        }

        public DeleteForm(bool unattended) {
            // Unbeaufsichtigt ist der Lauf entweder bei explizitem Autostart-Flag oder
            // wenn keine interaktive Benutzersitzung vorhanden ist (z. B. geplante Aufgabe
            // mit "Ausführen, auch wenn der Benutzer nicht angemeldet ist"). So müssen
            // bestehende Aufgaben nicht angepasst werden.
            this.unattended = unattended || !Environment.UserInteractive;

            protectedDirectories = new List<DirectoryInfo>();
            allowedSubdirectories = new List<DirectoryInfo>();

            LoadDirectoriesFromConfig(Properties.Settings.Default.protectedDirectories, protectedDirectories);
            LoadDirectoriesFromConfig(Properties.Settings.Default.allowedSubdirectories, allowedSubdirectories);

            var options = CleanupOptions.FromSettings();

            this.emptyTrash = options.EmptyTrash;
            this.minFileAgeDays = options.MinFileAgeDays;
            this.removeWindowsOld = options.RemoveWindowsOld;
            this.flushDnsCache = options.FlushDnsCache;
            this.clearEventLogs = options.ClearEventLogs;
            this.healthCheckEnabled = options.HealthCheckEnabled;
            this.minFreeDiskPercent = options.MinFreeDiskPercent;

            this.dryRun = options.DryRun;
            this.createRestorePoint = options.CreateRestorePoint;
            this.deleteToRecycleBin = options.DeleteToRecycleBin;
            this.cleanBrowserCaches = options.CleanBrowserCaches;
            this.dismComponentCleanup = options.DismComponentCleanup;
            this.cleanMinidumps = options.CleanMinidumps;
            this.cleanPrefetch = options.CleanPrefetch;
            this.cleanDeliveryOptimization = options.CleanDeliveryOptimization;
            this.cleanOldLogs = options.CleanOldLogs;
            this.oldLogAgeDays = options.OldLogAgeDays;
            this.emptyAllRecycleBins = options.EmptyAllRecycleBins;
            this.confirmBeforeDelete = options.ConfirmBeforeDelete;
            this.generateReport = options.GenerateReport;
            this.reportPath = options.ReportPath;
            this.reportFormat = options.ReportFormat;
            this.emailReport = options.EmailReport;

            this.excludeMatcher = new Services.ExcludeRuleMatcher(
                (options.ExcludePatterns ?? string.Empty)
                    .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));
            this.skipIfEnoughFreeSpace = options.SkipIfEnoughFreeSpace;
            this.schedulePendingDeletes = options.SchedulePendingDeletes;

            this.EnableSidebar = true;
            this.EnableMaximizeButton = true;
            this.AllowResize = true;
            this.ResizeMinimum = new Size(880, 640);
            this.CaptionSubtitle = "Temporäre Dateien bereinigen";

            InitializeComponent();

            InitializeDashboard();

            SetupTrayMenu();

            this.FormClosing += DeleteForm_FormClosing;
            this.Resize += DeleteForm_Resize;
        }

        // Richtet das Kontextmenü des Infobereich-Symbols ein und aktiviert
        // "In den Infobereich minimieren" per Doppelklick/Menü.
        private void SetupTrayMenu() {
            var menu = new ContextMenuStrip { Renderer = new UI.ThemedMenuRenderer() };

            var openItem = new ToolStripMenuItem("Fenster anzeigen");
            openItem.Click += (s, e) => RestoreFromTray();

            var settingsItem = new ToolStripMenuItem("Einstellungen...");
            settingsItem.Click += (s, e) => {
                RestoreFromTray();
                ShowPage(Page.Settings);
            };

            var exitItem = new ToolStripMenuItem("Beenden");
            exitItem.Click += (s, e) => ExitApplication();

            menu.Items.Add(openItem);
            menu.Items.Add(settingsItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(exitItem);

            notifyIcon1.ContextMenuStrip = menu;
            notifyIcon1.DoubleClick += (s, e) => RestoreFromTray();
        }

        private void RestoreFromTray() {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
            this.Activate();
        }

        private void DeleteForm_Resize(object sender, EventArgs e) {
            // Beim Minimieren in den Infobereich ausweichen statt in die Taskleiste.
            if (this.WindowState == FormWindowState.Minimized) {
                this.ShowInTaskbar = false;
                notifyIcon1.Visible = true;
                this.Hide();
                return;
            }

            ApplyWindowLayout();
        }

        private void DeleteForm_FormClosing(object sender, FormClosingEventArgs e) {
            if (e.CloseReason == CloseReason.UserClosing && !settingsPage.ConfirmLeave()) {
                e.Cancel = true;
                return;
            }

            // Sicherstellen, dass die Windows-Update-Dienste beim Beenden wieder laufen,
            // auch wenn die Anwendung während der Bereinigung geschlossen wird.
            if (updateServicesStopped) {
                SetWindowsUpdateServices(true);
            }
        }

        private void DeleteForm_Shown(object sender, EventArgs e) {
            LoadConfiguration();
            StartCleanup();

            if (healthCheckEnabled) {
                healthCheckTask = Task.Run(() => RunHealthCheck());
            }
        }

        private void StartCleanup() {
            cleanupCts?.Dispose();
            cleanupCts = new CancellationTokenSource();
            var token = cleanupCts.Token;
            PrepareDashboardForRun();
            Task.Run(() => CleanDirectories(token));
        }

        /// <summary>
        /// Ermittelt die geschätzte Gesamtgröße der zu bereinigenden Verzeichnisse und
        /// lässt den Benutzer den Löschvorgang bestätigen. Gibt false zurück, wenn der
        /// Benutzer abbricht.
        /// </summary>
        private async Task<bool> ConfirmSizePreview(CancellationToken cancellationToken) {
            Invoke(new Action(() => UpdateStatusLabel(label1, "Berechne Größe der zu bereinigenden Daten...", themeStatusText)));

            long totalBytes = await Task.Run(
                () => Services.SizeEstimator.EstimateSize(cleaningDirectories, cancellationToken),
                cancellationToken);

            string readable = Services.ByteFormatter.Format(totalBytes);
            LogToDetailTextBox($"Geschätzte Größe der zu bereinigenden Verzeichnisse: {readable}");

            bool confirmed = false;
            Invoke(new Action(() => {
                var result = UI.AppDialog.Show(
                    this,
                    $"Es werden ca. {readable} in {cleaningDirectories.Count} Verzeichnis(sen) bereinigt.{Environment.NewLine}{Environment.NewLine}Möchten Sie fortfahren?",
                    "Bereinigung bestätigen",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1);
                confirmed = result == DialogResult.Yes;
            }));

            return confirmed;
        }

        private void LoadConfiguration() {
            EnsureHandleCreated();

            cleaningDirectories = new List<DirectoryInfo>();

            string directoriesConfig = Properties.Settings.Default.cleanDirectories;
            if (string.IsNullOrEmpty(directoriesConfig)) {
                UI.AppDialog.Show(this, "Keine Verzeichnisse in der Konfigurationsdatei gefunden.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            foreach (var path in directoriesConfig.Split(';').Select(p => ResolveSpecialFolders(p.Trim()))) {
                if (string.IsNullOrEmpty(path) || !Directory.Exists(path)) {
                    LogToDetailTextBox($"Verzeichnis {path} existiert nicht oder ist leer.");
                } else if (IsProtectedDirectory(path)) {
                    LogToDetailTextBox($"Verzeichnis {path} wird nicht bereinigt, da es ein geschütztes Systemverzeichnis ist.");
                } else {
                    AddDirectoryToList(cleaningDirectories, path);
                }
            }
        }

        private void LoadDirectoriesFromConfig(string configValue, List<DirectoryInfo> directoryList) {
            if (string.IsNullOrEmpty(configValue)) return;

            foreach (var dir in configValue.Split(';').Select(p => ResolveSpecialFolders(p.Trim()))) {
                AddDirectoryToList(directoryList, dir);
            }
        }

        private void AddDirectoryToList(List<DirectoryInfo> list, string path) {
            if (Directory.Exists(path) && !list.Any(dir => dir.FullName == path)) {
                list.Add(new DirectoryInfo(path));
            }
        }

        private string ResolveSpecialFolders(string path) {
            return Services.SpecialFolderResolver.Resolve(path);
        }

        private string ResolveFirefoxCache() {
            return Services.SpecialFolderResolver.ResolveFirefoxCache();
        }

        private bool IsProtectedDirectory(string directory) {
            return Services.DirectoryProtectionService.IsProtected(
                directory,
                protectedDirectories.Select(d => d.FullName),
                allowedSubdirectories.Select(d => d.FullName));
        }

        private void EmptyRecycleBin() {
            if (dryRun) {
                LogToDetailTextBox("[Vorschau] Würde den Papierkorb leeren.");
                return;
            }

            try {
                // Leert den Papierkorb ohne Bestätigung, ohne UI und ohne Sound
                int result = NativeMethods.SHEmptyRecycleBin(IntPtr.Zero, null, NativeMethods.SHERB_NOCONFIRMATION | NativeMethods.SHERB_NOPROGRESSUI | NativeMethods.SHERB_NOSOUND);

                switch ((uint)result) {
                    case 0: // S_OK
                        LogToDetailTextBox("Papierkorb erfolgreich geleert.");
                        break;
                    case 0x8000FFFF: // E_UNEXPECTED – tritt auf, wenn der Papierkorb bereits leer ist
                        LogToDetailTextBox("Papierkorb ist bereits leer – nichts zu löschen.");
                        break;
                    case 0x80070005: // E_ACCESSDENIED
                        LogToDetailTextBox("Papierkorb konnte nicht geleert werden: Zugriff verweigert (Administratorrechte erforderlich).");
                        break;
                    default:
                        // HResult in lesbare Systemmeldung übersetzen.
                        string message = new System.ComponentModel.Win32Exception(result).Message;
                        LogToDetailTextBox($"Papierkorb konnte nicht geleert werden (0x{(uint)result:X8}): {message}");
                        break;
                }
            } catch (Exception ex) {
                LogError(ex);
            }
        }

        private void RunHealthCheck() {
            try {
                // Nur berechnen und zwischenspeichern – die Ausgabe erfolgt gebündelt
                // am Ende in CompleteProcess, nach den gelöschten Dateien.
                var issues = HealthCheck.Run(minFreeDiskPercent);
                pendingHealthIssues = issues;

                var critical = issues.Where(i => i.Severity == HealthSeverity.Critical).ToList();
                if (critical.Count > 0) {
                    severeHealthIssueDetected = true;
                    NotifySevereIssues(critical);
                }
            } catch (Exception ex) {
                LogError(ex);
            }
        }

        // Gibt die zwischengespeicherten Ergebnisse der Systemprüfung aus.
        // Wird am Ende aufgerufen, damit sie klar getrennt unter den gelöschten Dateien stehen.
        private void FlushHealthCheckResults() {
            if (!healthCheckEnabled) {
                return;
            }

            var issues = pendingHealthIssues;

            LogToDetailTextBox(string.Empty);
            LogToDetailTextBox("===== Systemprüfung =====");

            if (issues == null || issues.Count == 0) {
                LogToDetailTextBox("Systemprüfung: keine Auffälligkeiten.");
                return;
            }

            foreach (var issue in issues) {
                LogToDetailTextBox($"Systemprüfung: {issue}");
            }
        }

        private void NotifySevereIssues(List<HealthIssue> issues) {
            string machine = Environment.MachineName;
            string title = $"DataClean: Schwerwiegende Systemfehler auf {machine}";
            string body = "Folgende schwerwiegende Probleme wurden erkannt:" + Environment.NewLine + Environment.NewLine +
                string.Join(Environment.NewLine, issues.Select(i => $"- {i.Category}: {i.Message}")) +
                Environment.NewLine + Environment.NewLine +
                $"Erkannt am {DateTime.Now:yyyy-MM-dd HH:mm:ss}.";

            if (Properties.Settings.Default.notifyByTray) {
                ShowTrayNotification(title, body);
            }

            if (Properties.Settings.Default.notifyByEmail) {
                SendEmailNotification(title, body);
            }
        }

        private void ShowTrayNotification(string title, string body) {
            ShowTrayNotification(title, body, ToolTipIcon.Error);
        }

        private void ShowTrayNotification(string title, string body, ToolTipIcon icon) {
            try {
                Invoke(new Action(() => {
                    notifyIcon1.Visible = true;
                    notifyIcon1.BalloonTipTitle = title;
                    notifyIcon1.BalloonTipText = body;
                    notifyIcon1.BalloonTipIcon = icon;
                    notifyIcon1.ShowBalloonTip(15000);
                }));
            } catch (Exception ex) {
                LogError(ex);
            }
        }

        private void SendEmailNotification(string subject, string body) {
            try {
                var settings = Properties.Settings.Default;

                if (string.IsNullOrWhiteSpace(settings.smtpHost)
                    || string.IsNullOrWhiteSpace(settings.notifyFromAddress)
                    || string.IsNullOrWhiteSpace(settings.notifyToAddress)) {
                    LogToDetailTextBox("E-Mail-Benachrichtigung übersprungen: SMTP-Konfiguration unvollständig.");
                    return;
                }

                using (var message = new MailMessage(settings.notifyFromAddress, settings.notifyToAddress) {
                    Subject = subject,
                    Body = body
                })
                using (var client = new SmtpClient(settings.smtpHost, settings.smtpPort) {
                    EnableSsl = settings.smtpEnableSsl
                }) {
                    // Falls ein SMTP-Benutzer hinterlegt ist, das per DPAPI verschlüsselte
                    // Passwort entschlüsseln und zur Authentifizierung verwenden.
                    if (!string.IsNullOrWhiteSpace(settings.smtpUser)) {
                        string password = Services.SecureSettings.Decrypt(settings.smtpPasswordEncrypted);
                        client.Credentials = new System.Net.NetworkCredential(settings.smtpUser, password);
                    }

                    client.Send(message);
                    LogToDetailTextBox($"E-Mail-Benachrichtigung an {settings.notifyToAddress} gesendet.");
                }
            } catch (Exception ex) {
                LogToDetailTextBox($"E-Mail-Benachrichtigung fehlgeschlagen: {ex.Message}");
            }
        }

        private async Task RemoveWindowsOldAsync() {
            try {
                Invoke(new Action(() => {
                    UpdateStatusLabel(directoryValueLabel, "Windows.old");
                    UpdateStatusLabel(label1, dryRun ? "Prüfe Windows.old..." : "Entferne Windows.old...", themeStatusText);
                }));

                string systemDrive = Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.Windows));
                var windowsOld = new DirectoryInfo(Path.Combine(systemDrive, "Windows.old"));

                if (!windowsOld.Exists) {
                    LogToDetailTextBox("Windows.old nicht vorhanden – übersprungen.");
                    return;
                }

                long size = GetItemSize(windowsOld);
                bool deleted = await TryDeleteItemAsync(windowsOld);
                if (deleted) {
                    freedBytes += size;
                    if (!dryRun) {
                        LogToDetailTextBox($"Windows.old entfernt: {FormatBytes(size)}");
                    }
                } else {
                    LogToDetailTextBox("Windows.old konnte nicht vollständig entfernt werden (Administratorrechte erforderlich).");
                }
            } catch (Exception ex) {
                LogError(ex);
            }
        }

        private void FlushDnsCache() {
            if (dryRun) {
                LogToDetailTextBox("[Vorschau] Würde den DNS-Cache leeren.");
                return;
            }

            try {
                Invoke(new Action(() => UpdateStatusLabel(label1, "Leere DNS-Cache...", themeStatusText)));

                var startInfo = new System.Diagnostics.ProcessStartInfo("ipconfig", "/flushdns") {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (var process = System.Diagnostics.Process.Start(startInfo)) {
                    process.WaitForExit();
                    LogToDetailTextBox(process.ExitCode == 0
                        ? "DNS-Cache erfolgreich geleert."
                        : $"DNS-Cache konnte nicht geleert werden. Exit-Code: {process.ExitCode}");
                }
            } catch (Exception ex) {
                LogError(ex);
            }
        }

        private void ClearEventLogs() {
            if (dryRun) {
                LogToDetailTextBox("[Vorschau] Würde die Windows-Event-Logs leeren.");
                return;
            }

            try {
                Invoke(new Action(() => UpdateStatusLabel(label1, "Leere Windows-Event-Logs...", themeStatusText)));

                foreach (var logName in System.Diagnostics.EventLog.GetEventLogs().Select(l => l.Log).Distinct()) {
                    try {
                        using (var eventLog = new System.Diagnostics.EventLog(logName)) {
                            eventLog.Clear();
                        }
                        LogToDetailTextBox($"Event-Log geleert: {logName}");
                    } catch (Exception ex) when (ex is System.Security.SecurityException || ex is InvalidOperationException || ex is UnauthorizedAccessException) {
                        LogToDetailTextBox($"Event-Log {logName} konnte nicht geleert werden: {ex.Message}");
                    }
                }
            } catch (Exception ex) {
                LogError(ex);
            }
        }

        private void SetWindowsUpdateServices(bool start) {
            // wuauserv (Windows Update) und bits (BITS) sperren SoftwareDistribution\Download.
            string[] services = { "wuauserv", "bits" };
            string action = start ? "start" : "stop";

            foreach (var service in services) {
                try {
                    Invoke(new Action(() => UpdateStatusLabel(label1,
                        $"Dienst {service} wird {(start ? "gestartet" : "gestoppt")}...", themeStatusText)));

                    var psi = new System.Diagnostics.ProcessStartInfo("net", $"{action} {service}") {
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    };

                    using (var process = System.Diagnostics.Process.Start(psi)) {
                        process.WaitForExit(30000);
                    }
                } catch (Exception ex) {
                    LogToDetailTextBox($"Dienst {service} konnte nicht {(start ? "gestartet" : "gestoppt")} werden: {ex.Message}");
                }
            }

            if (!start) {
                // Kurze Pause, damit Dateihandles freigegeben werden.
                System.Threading.Thread.Sleep(1000);
            }

            // Zustand merken, damit die Dienste beim Beenden zuverlässig wiederhergestellt werden.
            updateServicesStopped = !start;
        }

        private async Task CreateRestorePointAsync() {
            try {
                Invoke(new Action(() => UpdateStatusLabel(label1, "Erstelle Systemwiederherstellungspunkt...", themeStatusText)));
                performedTasks.Add("Systemwiederherstellungspunkt");

                // Checkpoint-Computer legt einen Wiederherstellungspunkt an. Erfordert aktivierte
                // Systemwiederherstellung und Administratorrechte; Fehler werden toleriert.
                string psCommand =
                    "Checkpoint-Computer -Description 'DataClean Bereinigung' -RestorePointType 'MODIFY_SETTINGS'";

                var psi = new System.Diagnostics.ProcessStartInfo("powershell.exe",
                    $"-NoProfile -ExecutionPolicy Bypass -Command \"{psCommand}\"") {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                await Task.Run(() => {
                    using (var process = System.Diagnostics.Process.Start(psi)) {
                        string stderr = process.StandardError.ReadToEnd();
                        process.WaitForExit(120000);

                        if (process.ExitCode == 0) {
                            LogToDetailTextBox("Systemwiederherstellungspunkt erfolgreich erstellt.");
                        } else {
                            LogToDetailTextBox($"Systemwiederherstellungspunkt konnte nicht erstellt werden: {stderr.Trim()}");
                        }
                    }
                });
            } catch (Exception ex) {
                LogToDetailTextBox($"Systemwiederherstellungspunkt übersprungen: {ex.Message}");
            }
        }

        // Prüft, ob auf dem Systemlaufwerk mehr als minFreeDiskPercent freier Speicher vorhanden ist.
        private bool HasEnoughFreeSpace() {
            if (minFreeDiskPercent <= 0) {
                return false;
            }

            try {
                string systemRoot = Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.System));
                var drive = new DriveInfo(systemRoot);
                if (!drive.IsReady || drive.TotalSize <= 0) {
                    return false;
                }

                double freePercent = (double)drive.TotalFreeSpace / drive.TotalSize * 100.0;
                return freePercent >= minFreeDiskPercent;
            } catch (Exception ex) {
                Services.Logger.Warn($"Schwellenwert-Prüfung fehlgeschlagen: {ex.Message}", "DeleteForm");
                return false;
            }
        }

        private async Task CleanDirectories(CancellationToken cancellationToken) {
            try {
                totalPhases = GetEnabledPhaseNames().Count;
                currentPhase = 0;

                // Schwellenwert-Trigger: Wenn aktiviert und auf dem Systemlaufwerk noch
                // genügend freier Speicher vorhanden ist, wird der Lauf übersprungen.
                if (skipIfEnoughFreeSpace && HasEnoughFreeSpace()) {
                    LogToDetailTextBox($"=== Bereinigung übersprungen – genügend freier Speicher (> {minFreeDiskPercent} %) ===");
                    Services.Logger.Info($"Lauf übersprungen – freier Speicher über Schwellenwert ({minFreeDiskPercent} %).", "DeleteForm");
                    Invoke(new Action(() => {
                        ShowFinishedState("Bereinigung übersprungen");
                        UpdateStatusLabel(label1, $"Genügend freier Speicher (> {minFreeDiskPercent} %).", themeSuccess);
                    }));
                    if (unattended) {
                        Invoke(new Action(() => {
                            notifyIcon1.Visible = false;
                            Application.Exit();
                        }));
                    }
                    return;
                }

                if (dryRun) {
                    Invoke(new Action(() => UpdateStatusLabel(label1, "Vorschau-Modus – es wird nichts gelöscht.", themeStatusText)));
                    LogToDetailTextBox("=== VORSCHAU-MODUS (Dry-Run) – es werden keine Daten gelöscht ===");
                } else {
                    // Vor dem eigentlichen Löschen eine Größen-Vorschau anzeigen und
                    // vom Benutzer bestätigen lassen. Die Rückfrage entfällt im
                    // unbeaufsichtigten Modus oder wenn sie per Einstellung
                    // (confirmBeforeDelete) deaktiviert wurde.
                    if (!unattended && confirmBeforeDelete && !await ConfirmSizePreview(cancellationToken)) {
                        LogToDetailTextBox("=== Bereinigung durch Benutzer abgebrochen (Vorschau) ===");
                        // Der Benutzer hat das Löschen abgelehnt, bevor etwas passiert ist –
                        // die Anwendung kann direkt geschlossen werden.
                        Invoke(new Action(() => {
                            notifyIcon1.Visible = false;
                            Application.Exit();
                        }));
                        return;
                    }
                }

                // Optional einen Systemwiederherstellungspunkt anlegen, bevor größere
                // Löschvorgänge beginnen (nicht im Vorschau-Modus).
                if (createRestorePoint && !dryRun) {
                    await CreateRestorePointAsync();
                }

                // Wenn der Windows-Update-Download-Ordner bereinigt werden soll, müssen die
                // sperrenden Dienste (Windows Update, BITS) zuvor gestoppt werden. Im
                // Vorschau-Modus wird nichts gelöscht, die Dienste laufen also einfach weiter.
                bool needsUpdateServices = !dryRun && cleaningDirectories.Any(d =>
                    d.FullName.IndexOf("SoftwareDistribution", StringComparison.OrdinalIgnoreCase) >= 0);

                if (needsUpdateServices) {
                    SetWindowsUpdateServices(false);
                }

                try {
                    await BeginPhaseAsync("Verzeichnisse bereinigen", cancellationToken);
                    for (int i = 0; i < cleaningDirectories.Count; i++) {
                        var scanDir = cleaningDirectories[i];
                        int row = i;
                        await WaitIfPausedAsync(cancellationToken);
                        cancellationToken.ThrowIfCancellationRequested();
                        Invoke(new Action(() => {
                            UpdateStatusLabel(directoryValueLabel, $"{scanDir}");
                            SetRowStatus(directoriesList, row, UI.StatusKind.Running, "Läuft");
                        }));
                        await ProcessDirectory(scanDir, cancellationToken);
                        Invoke(new Action(() => {
                            directoriesDone++;
                            progressBar1.Value = 0;
                            SetRowStatus(directoriesList, row, UI.StatusKind.Success, "Fertig");
                        }));
                    }
                } finally {
                    if (needsUpdateServices) {
                        SetWindowsUpdateServices(true);
                    }
                }

                if (emptyTrash) {
                    await BeginPhaseAsync("Papierkorb leeren", cancellationToken);
                    UpdateStatusLabel(directoryValueLabel, "Papierkorb");
                    if (!dryRun) {
                        UpdateStatusLabel(label1, "Leere Papierkorb...", themeStatusText);
                    }
                    EmptyRecycleBin();
                    if (!dryRun) {
                        UpdateStatusLabel(label1, "Papierkorb geleert.", themeStatusText);
                    }
                }

                if (removeWindowsOld) {
                    await BeginPhaseAsync("Windows.old entfernen", cancellationToken);
                    await RemoveWindowsOldAsync();
                }

                if (cleanBrowserCaches) {
                    await BeginPhaseAsync("Browser-Caches", cancellationToken);
                    await CleanBrowserCachesAsync(cancellationToken);
                }

                if (cleanMinidumps) {
                    await BeginPhaseAsync("Absturz-Dumps", cancellationToken);
                    await CleanMinidumpsAsync(cancellationToken);
                }

                if (cleanPrefetch) {
                    await BeginPhaseAsync("Prefetch", cancellationToken);
                    await CleanPrefetchAsync(cancellationToken);
                }

                if (cleanDeliveryOptimization) {
                    await BeginPhaseAsync("Delivery Optimization", cancellationToken);
                    await CleanDeliveryOptimizationAsync(cancellationToken);
                }

                if (cleanOldLogs) {
                    await BeginPhaseAsync("Alte Protokolle", cancellationToken);
                    await CleanOldLogsAsync(cancellationToken);
                }

                if (emptyAllRecycleBins) {
                    await BeginPhaseAsync("Papierkorb aller Laufwerke", cancellationToken);
                    EmptyAllRecycleBins();
                }

                if (dismComponentCleanup && !dryRun) {
                    await BeginPhaseAsync("DISM Komponentenbereinigung", cancellationToken);
                    await RunDismComponentCleanupAsync();
                }

                if (flushDnsCache) {
                    await BeginPhaseAsync("DNS-Cache leeren", cancellationToken);
                    FlushDnsCache();
                }

                if (clearEventLogs) {
                    await BeginPhaseAsync("Event-Logs leeren", cancellationToken);
                    ClearEventLogs();
                }

                await CompleteProcess();
            } catch (OperationCanceledException) {
                // Auch bei Abbruch die bis dahin gesammelte Statistik als Report sichern.
                if (generateReport) {
                    WriteReport();
                }
                CompleteProcessCanceled();
            } catch (Exception ex) {
                // Bei unerwartetem Fehler Report noch schreiben, damit die Statistik nicht verloren geht.
                if (generateReport) {
                    WriteReport();
                }
                LogError(ex);
                FlushUiUpdates(force: true);
                Invoke(new Action(() => {
                    FinishRunTables(canceled: true);
                    ShowFinishedState("Bereinigung fehlgeschlagen");
                    UpdateStatusLabel(label1, $"Fehler: {ex.Message}", themeError);
                }));
            }
        }

        // Liefert die aktivierten Phasen in Ausführungsreihenfolge (Fortschritt und Phasenliste).
        private List<string> GetEnabledPhaseNames() {
            var phases = new List<string> { "Verzeichnisse bereinigen" };
            if (emptyTrash) phases.Add("Papierkorb leeren");
            if (removeWindowsOld) phases.Add("Windows.old entfernen");
            if (cleanBrowserCaches) phases.Add("Browser-Caches");
            if (cleanMinidumps) phases.Add("Absturz-Dumps");
            if (cleanPrefetch) phases.Add("Prefetch");
            if (cleanDeliveryOptimization) phases.Add("Delivery Optimization");
            if (cleanOldLogs) phases.Add("Alte Protokolle");
            if (emptyAllRecycleBins) phases.Add("Papierkorb aller Laufwerke");
            if (dismComponentCleanup && !dryRun) phases.Add("DISM Komponentenbereinigung");
            if (flushDnsCache) phases.Add("DNS-Cache leeren");
            if (clearEventLogs) phases.Add("Event-Logs leeren");
            return phases;
        }

        // Startet die nächste Phase: wartet eine Pause ab, beachtet einen Abbruch und
        // aktualisiert Phasenanzeige und Phasenliste.
        private async Task BeginPhaseAsync(string description, CancellationToken cancellationToken) {
            await WaitIfPausedAsync(cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();

            currentPhase++;
            int phase = currentPhase;
            int total = totalPhases;
            Action update = () => {
                phaseLabel.Text = $"Phase {phase}/{total}: {description}";
                MarkPhaseRunning(phase - 1);
            };
            if (InvokeRequired) BeginInvoke(update); else update();
        }

        private async Task ProcessDirectory(DirectoryInfo dir, CancellationToken cancellationToken) {
            try {
                var files = dir.EnumerateFiles("*", SearchOption.AllDirectories).ToList();

                // Verzeichnisse nach Tiefe sortieren (tiefste zuerst), damit untergeordnete
                // Ordner vor den übergeordneten gelöscht werden.
                var directories = dir.EnumerateDirectories("*", SearchOption.AllDirectories)
                    .OrderByDescending(d => d.FullName.Count(c => c == Path.DirectorySeparatorChar))
                    .ToList();

                Invoke(new Action(() => {
                    progressBar1.Value = 0;
                    progressBar1.Maximum = files.Count + directories.Count;
                }));

                // Dateien parallel löschen (unabhängig voneinander), Verzeichnisse danach
                // sequenziell in Tiefensortierung (Reihenfolge ist hier relevant).
                await DeleteFilesParallelAsync(files, cancellationToken);
                await DeleteItemsAsync(directories, cancellationToken);
            } catch (OperationCanceledException) {
                throw;
            } catch (UnauthorizedAccessException ex) {
                LogToDetailTextBox($"Zugriff verweigert auf {dir.FullName}: {ex.Message}");
            } catch (Exception ex) {
                LogError(ex);
            }
        }

        private async Task DeleteItemsAsync<T>(IEnumerable<T> items, CancellationToken cancellationToken) where T : FileSystemInfo {
            foreach (var item in items) {
                await WaitIfPausedAsync(cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                string itemName = item.FullName;

                if (ShouldSkipDueToAge(item)) {
                    System.Threading.Interlocked.Increment(ref skippedItemCount);
                    LogToDetailTextBox($"Übersprungen (zu jung): {itemName}");
                    RecordActivity(ActivityStatus.Skipped, item);
                    QueueProgressStep();
                    continue;
                }

                // Größe nur für Dateien erfassen. Verzeichnisse sind beim Löschen bereits leer
                // (alle Dateien werden zuvor separat gelöscht), eine rekursive Größenermittlung
                // hier wäre langsam und würde die Dateigrößen doppelt zählen.
                long itemSize = (item is FileInfo file) ? GetFileLength(file) : 0;

                bool isDeleted = await TryDeleteItemAsync(item);

                if (isDeleted) {
                    freedBytes += itemSize;
                    System.Threading.Interlocked.Increment(ref deletedItemCount);
                    LogToDetailTextBox($"Gelöscht: {itemName}");
                }

                QueueProgressStep();
            }

            FlushUiUpdates(force: true);
        }

        private async Task DeleteFilesParallelAsync(IReadOnlyList<FileInfo> files, CancellationToken cancellationToken) {
            if (files.Count == 0) {
                return;
            }

            // Parallelitätsgrad an die CPU koppeln, aber begrenzen, um das Dateisystem
            // nicht zu überlasten. Bei sehr wenigen Dateien lohnt Parallelität nicht.
            int maxParallel = Math.Min(Environment.ProcessorCount * 2, 8);
            if (files.Count < 2 || maxParallel <= 1) {
                await DeleteItemsAsync(files, cancellationToken);
                return;
            }

            using (var throttler = new System.Threading.SemaphoreSlim(maxParallel)) {
                var tasks = new List<Task>(files.Count);

                foreach (var file in files) {
                    await WaitIfPausedAsync(cancellationToken);
                    cancellationToken.ThrowIfCancellationRequested();
                    await throttler.WaitAsync(cancellationToken);

                    tasks.Add(Task.Run(async () => {
                        try {
                            await DeleteSingleFileAsync(file, cancellationToken);
                        } finally {
                            throttler.Release();
                        }
                    }, cancellationToken));
                }

                await Task.WhenAll(tasks);
            }

            FlushUiUpdates(force: true);
        }

        private async Task DeleteSingleFileAsync(FileInfo file, CancellationToken cancellationToken) {
            string itemName = file.FullName;

            cancellationToken.ThrowIfCancellationRequested();

            if (ShouldSkipDueToAge(file)) {
                System.Threading.Interlocked.Increment(ref skippedItemCount);
                LogToDetailTextBox($"Übersprungen (zu jung): {itemName}");
                RecordActivity(ActivityStatus.Skipped, file);
                QueueProgressStep();
                return;
            }

            long itemSize = GetFileLength(file);

            bool isDeleted = await TryDeleteItemAsync(file);

            if (isDeleted) {
                System.Threading.Interlocked.Add(ref freedBytes, itemSize);
                System.Threading.Interlocked.Increment(ref deletedItemCount);
                LogToDetailTextBox($"Gelöscht: {itemName}");
            }

            QueueProgressStep();
        }

        private void QueueProgressStep() {
            System.Threading.Interlocked.Increment(ref pendingProgressSteps);

            // UI nur periodisch aktualisieren, um blockierende Invoke-Aufrufe pro Element zu vermeiden.
            if (uiThrottle.ElapsedMilliseconds >= UiUpdateIntervalMs) {
                FlushUiUpdates(force: false);
            }
        }

        private void FlushUiUpdates(bool force) {
            int steps = System.Threading.Interlocked.Exchange(ref pendingProgressSteps, 0);
            uiThrottle.Restart();

            if (steps <= 0 && !force) {
                FlushDetailLog(force);
                return;
            }

            Action update = () => {
                progressBar1.Value += steps;
            };

            if (force) {
                if (InvokeRequired) Invoke(update); else update();
            } else {
                if (InvokeRequired) BeginInvoke(update); else update();
            }

            FlushDetailLog(force);
        }

        private bool ShouldSkipDueToAge(FileSystemInfo item) {
            // Altersfilter gilt nur für Dateien; leere Ordner werden weiterhin entfernt.
            if (minFileAgeDays <= 0 || !(item is FileInfo file)) {
                return false;
            }

            try {
                // Kein Refresh nötig: Die FileInfo stammt frisch aus der Enumeration,
                // die Zeitstempel sind bereits geladen.
                return Services.FileAgeFilter.ShouldSkip(file.LastWriteTime, file.LastAccessTime, minFileAgeDays, DateTime.Now);
            } catch {
                return false;
            }
        }

        private long GetFileLength(FileInfo file) {
            try {
                return file.Length; // Wert stammt aus der Enumeration, kein zusätzlicher I/O-Zugriff.
            } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) {
                return 0;
            }
        }

        private long GetItemSize(FileSystemInfo item) {
            try {
                if (item is FileInfo file) {
                    file.Refresh();
                    return file.Exists ? file.Length : 0;
                }

                if (item is DirectoryInfo directory && directory.Exists) {
                    return directory.EnumerateFiles("*", SearchOption.AllDirectories).Sum(f => {
                        try { return f.Length; } catch { return 0L; }
                    });
                }
            } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) {
                // Größe kann nicht ermittelt werden – als 0 werten.
            }

            return 0;
        }

        private async Task<bool> TryDeleteItemAsync(FileSystemInfo item) {
            bool deleted = await TryDeleteItemCoreAsync(item);
            if (deleted) {
                RecordActivity(ActivityStatus.Deleted, item);
            }
            return deleted;
        }

        private async Task<bool> TryDeleteItemCoreAsync(FileSystemInfo item) {
            const int maxAttempts = 3;
            string itemName = item.FullName;
            bool isFile = item is FileInfo;
            bool recursive = minFileAgeDays <= 0;
            Exception lastException = null;

            // Benutzerdefinierte Ausschlussmuster respektieren: passende Elemente werden
            // nicht gelöscht, sondern als übersprungen gezählt.
            if (excludeMatcher != null && excludeMatcher.IsExcluded(itemName)) {
                System.Threading.Interlocked.Increment(ref skippedItemCount);
                LogToDetailTextBox($"Durch Ausschlussregel übersprungen: {itemName}");
                RecordActivity(ActivityStatus.Skipped, item);
                return false;
            }

            // Dry-Run: nichts löschen, nur protokollieren und als "gelöscht" werten,
            // damit Größen-/Zählstatistik die Vorschau korrekt darstellt.
            if (dryRun) {
                LogToDetailTextBox($"[Vorschau] Würde löschen: {itemName}");
                return true;
            }

            for (int attempt = 1; attempt <= maxAttempts; attempt++) {
                try {
                    // Optional in den Papierkorb verschieben (umkehrbar) statt endgültig löschen.
                    if (deleteToRecycleBin) {
                        if (!isFile && !recursive && Directory.EnumerateFileSystemEntries(itemName).Any()) {
                            return false;
                        }
                        if (!DeleteToRecycleBin(itemName)) {
                            break;
                        }
                        return true;
                    }

                    // Fast-Path: direkt löschen, ohne vorheriges Refresh/Exists/Attributes-Lesen.
                    // Das ist der häufigste Fall und vermeidet mehrere Metadaten-Zugriffe pro Element.
                    if (isFile) {
                        File.Delete(itemName);
                    } else {
                        // Bei aktivem Altersfilter nur leere Verzeichnisse entfernen,
                        // damit noch zu junge Dateien nicht rekursiv mitgelöscht werden.
                        if (!recursive && Directory.EnumerateFileSystemEntries(itemName).Any()) {
                            return false;
                        }
                        Directory.Delete(itemName, recursive);
                    }

                    return true;
                } catch (Exception ex) when (ex is FileNotFoundException || ex is DirectoryNotFoundException) {
                    // Bereits von einem übergeordneten Löschvorgang entfernt – kein Fehler.
                    return false;
                } catch (UnauthorizedAccessException ex) {
                    lastException = ex;
                    // Häufigste Ursache: ReadOnly-Attribut. Einmalig entfernen und erneut versuchen.
                    if (!TryRemoveReadOnly(item)) {
                        break; // Attribut war nicht gesetzt → echtes Rechteproblem, Retry zwecklos.
                    }
                } catch (IOException ex) {
                    lastException = ex;
                    if (attempt < maxAttempts) {
                        await Task.Delay(100 * attempt); // Wachsende Wartezeit vor erneutem Versuch
                    }
                }
            }

            // Gesperrte/in Verwendung befindliche Dateien sind erwartbar und werden
            // ruhiger als "übersprungen" protokolliert, damit echte Fehler nicht untergehen.
            if (lastException is IOException) {
                System.Threading.Interlocked.Increment(ref skippedItemCount);
                LogToDetailTextBox($"Übersprungen (in Verwendung): {itemName}");
                RecordActivity(ActivityStatus.Skipped, item);

                // Optional: gesperrte Datei zum Löschen beim nächsten Neustart vormerken.
                if (schedulePendingDeletes && isFile && !dryRun
                    && Services.PendingDeleteService.ScheduleDeleteOnReboot(itemName)) {
                    LogToDetailTextBox($"Zum Löschen beim nächsten Neustart vorgemerkt: {itemName}");
                }
            } else {
                System.Threading.Interlocked.Increment(ref failedItemCount);
                LogToDetailTextBox($"Löschen von {itemName} fehlgeschlagen: {lastException?.Message}");
                RecordActivity(ActivityStatus.Failed, item);
            }
            return false;
        }

        private bool TryRemoveReadOnly(FileSystemInfo item) {
            try {
                item.Refresh();
                if (item.Attributes.HasFlag(FileAttributes.ReadOnly)) {
                    item.Attributes &= ~FileAttributes.ReadOnly;
                    return true;
                }
            } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException) {
                // Ignorieren – der nächste Löschversuch meldet den Fehler.
            }
            return false;
        }

        private bool DeleteToRecycleBin(string path) {
            try {
                var fileOp = new NativeMethods.SHFILEOPSTRUCT {
                    wFunc = NativeMethods.FO_DELETE,
                    // pFrom muss mit einem zusätzlichen Nullzeichen terminiert sein (doppelt null).
                    pFrom = path + '\0' + '\0',
                    fFlags = (ushort)(NativeMethods.FOF_ALLOWUNDO | NativeMethods.FOF_NOCONFIRMATION | NativeMethods.FOF_NOERRORUI | NativeMethods.FOF_SILENT)
                };

                int result = NativeMethods.SHFileOperation(ref fileOp);
                return result == 0 && fileOp.fAnyOperationsAborted == 0;
            } catch (Exception ex) {
                LogToDetailTextBox($"Papierkorb-Verschiebung fehlgeschlagen für {path}: {ex.Message}");
                return false;
            }
        }

        private async Task CompleteProcess() {
            // Sicherstellen, dass der parallele Health-Check abgeschlossen ist, und
            // dessen Ergebnisse jetzt – nach den gelöschten Dateien – gebündelt ausgeben.
            try {
                healthCheckTask?.Wait(TimeSpan.FromSeconds(30));
            } catch (Exception ex) {
                Services.Logger.Warn($"HealthCheck-Wait: {ex.Message}", "DeleteForm");
            }

            FlushHealthCheckResults();

            // Kompakte Zusammenfassung ans Ende stellen, damit das Ergebnis auf einen Blick erkennbar ist.
            WriteSummary();

            RecordRunHistory(canceled: false);

            string freedText = $"Freigegebener Speicher: {FormatBytes(freedBytes)}";

            // Abschluss-Benachrichtigung im Infobereich (falls aktiviert) – besonders
            // nützlich im minimierten Zustand oder beim unbeaufsichtigten Lauf.
            if (Properties.Settings.Default.notifyByTray) {
                bool withErr = failedItemCount > 0;
                ShowTrayNotification(
                    withErr ? "DataClean – mit Fehlern abgeschlossen" : "DataClean – Bereinigung abgeschlossen",
                    $"{deletedItemCount} Elemente entfernt. {freedText}.",
                    withErr ? ToolTipIcon.Warning : ToolTipIcon.Info);
            }

            // Optionalen Report (CSV/HTML) auf die Platte schreiben.
            if (generateReport) {
                WriteReport();
            }

            // Verbleibende gepufferte Log-Zeilen synchron in die TextBox schreiben.
            FlushUiUpdates(force: true);

            bool withErrors = failedItemCount > 0;
            Invoke(new Action(() => {
                progressBar1.Value = progressBar1.Maximum = 1;
                overallFraction = 1d;
                SetOverallProgress(1d);
                phaseLabel.Text = string.Empty;
                FinishRunTables(canceled: false);
                ShowFinishedState(withErrors
                    ? "Mit Fehlern abgeschlossen"
                    : (dryRun ? "Vorschau abgeschlossen" : "Bereinigung abgeschlossen"));
                UpdateStatusLabel(label1, (withErrors
                    ? "Mit Fehlern abgeschlossen. Siehe Details."
                    : "Erfolgreich abgeschlossen.") + $" {freedText}", withErrors ? themeError : themeSuccess);
            }));

            await Task.Delay(2500);

            if (severeHealthIssueDetected) {
                Invoke(new Action(() => {
                    UpdateStatusLabel(label1, "Schwerwiegender Systemfehler erkannt – Programm bleibt geöffnet. Siehe Details.", themeError);
                    if (currentPage != Page.Settings) {
                        ShowPage(Page.Log);
                    }
                }));
                return; // Kein automatisches Schließen bei schwerem Fehler.
            }

            // Automatisch schließen, sofern der Benutzer nicht mit Details, Protokoll
            // oder Einstellungen arbeitet.
            if (!StayOpenAfterRun) {
                await Task.Delay(5000);
            }
            if (!StayOpenAfterRun) {
                Application.Exit();
            }
        }

        // Abschlusslogik nach einem vom Benutzer ausgelösten Abbruch.
        private void CompleteProcessCanceled() {
            Services.Logger.Warn("Bereinigung wurde vom Benutzer abgebrochen.", "DeleteForm");

            RecordRunHistory(canceled: true);

            LogToDetailTextBox("=== Vorgang abgebrochen \u2013 nicht alle Bereiche wurden bereinigt ===");
            FlushUiUpdates(force: true);

            Invoke(new Action(() => {
                FinishRunTables(canceled: true);
                ShowFinishedState("Vorgang abgebrochen");
                UpdateStatusLabel(label1, $"Nicht alle Bereiche wurden bereinigt. Freigegebener Speicher: {FormatBytes(freedBytes)}", themeError);
            }));
        }

        // Schreibt die Lauf-Historie (einmalig pro Lauf) in den persistenten Speicher.
        private bool runHistoryRecorded = false;
        private void RecordRunHistory(bool canceled) {
            if (runHistoryRecorded) {
                return;
            }
            runHistoryRecorded = true;

            try {
                Services.RunHistoryStore.Append(new Services.RunHistoryEntry {
                    StartTime = runStartTime,
                    EndTime = DateTime.Now,
                    FreedBytes = freedBytes,
                    DeletedItems = deletedItemCount,
                    SkippedItems = skippedItemCount,
                    FailedItems = failedItemCount,
                    DryRun = dryRun,
                    Unattended = unattended,
                    Canceled = canceled,
                    Tasks = string.Join(", ", performedTasks)
                });
            } catch (Exception ex) {
                Services.Logger.Error(ex, "DeleteForm.RecordRunHistory");
            }
        }

        // Erstellt die Abschluss-Zusammenfassung mit Statistik.
        private void WriteSummary() {
            TimeSpan duration = DateTime.Now - runStartTime;

            LogToDetailTextBox(string.Empty);
            LogToDetailTextBox("===== Zusammenfassung =====");
            if (dryRun) {
                LogToDetailTextBox("Vorschau-Modus (Dry-Run): Es wurden keine Daten gelöscht.");
            }
            LogToDetailTextBox($"Gelöscht: {deletedItemCount} Elemente / {FormatBytes(freedBytes)}");
            LogToDetailTextBox($"Übersprungen: {skippedItemCount}  ·  Fehler: {failedItemCount}");
            LogToDetailTextBox($"Dauer: {duration:hh\\:mm\\:ss}");
            if (performedTasks.Count > 0) {
                LogToDetailTextBox($"Zusatzaufgaben: {string.Join(", ", performedTasks)}");
            }
            if (!IsAdministrator()) {
                LogToDetailTextBox("Hinweis: Ohne Administratorrechte wurden einige Systemverzeichnisse übersprungen. Über \"Erweitertes Löschen\" mit Adminrechten neu starten.");
            }
        }

        // Schreibt einen Report (CSV oder HTML) in das konfigurierte Verzeichnis.
        private void WriteReport() {
            // Report nur einmal schreiben, auch wenn er aus mehreren Pfaden (Erfolg, Abbruch, Fehler) aufgerufen wird.
            if (reportWritten) {
                return;
            }
            reportWritten = true;

            try {
                string targetDir = string.IsNullOrWhiteSpace(reportPath)
                    ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DataClean", "Reports")
                    : reportPath;
                Directory.CreateDirectory(targetDir);

                bool isHtml = string.Equals(reportFormat, "HTML", StringComparison.OrdinalIgnoreCase);
                string fileName = Services.ReportBuilder.BuildFileName(runStartTime, isHtml);
                string fullPath = Path.Combine(targetDir, fileName);

                TimeSpan duration = DateTime.Now - runStartTime;
                string tasks = performedTasks.Count > 0 ? string.Join(", ", performedTasks) : "-";

                var reportData = new Services.ReportData {
                    RunStartTime = runStartTime,
                    MachineName = Environment.MachineName,
                    DryRun = dryRun,
                    DeletedItemCount = deletedItemCount,
                    FreedBytes = freedBytes,
                    SkippedItemCount = skippedItemCount,
                    FailedItemCount = failedItemCount,
                    Duration = duration,
                    Tasks = tasks
                };

                string content = isHtml
                    ? Services.ReportBuilder.BuildHtml(reportData)
                    : Services.ReportBuilder.BuildCsv(reportData);
                File.WriteAllText(fullPath, content, System.Text.Encoding.UTF8);

                LogToDetailTextBox($"Report gespeichert: {fullPath}");

                if (emailReport) {
                    SendEmailNotification($"DataClean Report – {Environment.MachineName}",
                        $"Gelöscht: {deletedItemCount} Elemente / {FormatBytes(freedBytes)}" + Environment.NewLine +
                        $"Übersprungen: {skippedItemCount}, Fehler: {failedItemCount}" + Environment.NewLine +
                        $"Dauer: {duration:hh\\:mm\\:ss}" + Environment.NewLine +
                        $"Report: {fullPath}");
                }
            } catch (Exception ex) {
                LogToDetailTextBox($"Report konnte nicht erstellt werden: {ex.Message}");
            }
        }

        private static string FormatBytes(long bytes) {
            return Services.ByteFormatter.Format(bytes);
        }

        private void UpdateStatusLabel(Label label, string text, Color? color = null) {
            Invoke(new Action(() => {
                label.Text = text;
                if (color.HasValue) label.ForeColor = color.Value;
            }));
        }

        private void LogToDetailTextBox(string message) {
            // Nicht-blockierend puffern; der Hotpath darf nicht pro Element auf den UI-Thread warten.
            lock (detailLogLock) {
                detailLogBuffer.Append(message).Append(Environment.NewLine);
            }
        }

        private void FlushDetailLog(bool force = false) {
            string pending;
            lock (detailLogLock) {
                if (detailLogBuffer.Length == 0) return;
                pending = detailLogBuffer.ToString();
                detailLogBuffer.Clear();
            }

            Action append = () => detailTextBox.AppendText(pending);
            if (force) {
                if (InvokeRequired) Invoke(append); else append();
            } else {
                if (InvokeRequired) BeginInvoke(append); else append();
            }
        }

        private void LogError(Exception ex) {
            Services.Logger.Error(ex);
            LogToDetailTextBox($"Fehler: {ex.Message}");
        }

        private bool IsAdministrator() {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private void RetryButton_Click(object sender, EventArgs e) {
            if (!IsAdministrator()) {
                var exeName = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
                var startInfo = new System.Diagnostics.ProcessStartInfo(exeName) {
                    Verb = "runas"
                };

                try {
                    System.Diagnostics.Process.Start(startInfo);
                    Application.Exit();
                } catch (Exception ex) {
                    UI.AppDialog.Show(this, $"Fehler beim Neustarten mit Admin-Rechten: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            } else {
                StartCleanup();
            }
        }

        private void CancelButton_Click(object sender, EventArgs e) {
            if (cleanupCts == null || cleanupCts.IsCancellationRequested) {
                return;
            }

            cleanupCts.Cancel();
            cancelButton.Enabled = false;
            pauseButton.Enabled = false;
            UpdateStatusLabel(label1, "Abbruch angefordert \u2013 laufender Vorgang wird beendet...", themeStatusText);
            LogToDetailTextBox("=== ABBRUCH durch Benutzer angefordert ===");
        }

        private void LinkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            OpenMailClient(linkLabel2.Text);
        }

        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            OpenWebPage("https://schuppelius.org");
        }

        private void OpenMailClient(string email) {
            try {
                System.Diagnostics.Process.Start($"mailto:{email}?subject=Anfrage");
            } catch (Exception ex) {
                UI.AppDialog.Show(this, $"Fehler beim Öffnen des Mailprogramms: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenWebPage(string url) {
            try {
                System.Diagnostics.Process.Start(url);
            } catch (Exception ex) {
                UI.AppDialog.Show(this, $"Fehler beim Öffnen der Webseite: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void EnsureHandleCreated() {
            if (!this.IsHandleCreated) {
                this.CreateHandle();
            }
        }
    }
}
