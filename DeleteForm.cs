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
    public partial class DeleteForm : Form {
        // Definition der SHEmptyRecycleBin Funktion aus der shell32.dll
        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, uint dwFlags);

        // Weist einem Steuerelement ein Windows-Theme zu (z. B. dunkle Scrollbars/ProgressBar im Dark-Mode).
        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
        private static extern int SetWindowTheme(IntPtr hWnd, string pszSubAppName, string pszSubIdList);

        // SHFileOperation zum Verschieben in den Papierkorb
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct SHFILEOPSTRUCT {
            public IntPtr hwnd;
            public uint wFunc;
            public string pFrom;
            public string pTo;
            public ushort fFlags;
            public int fAnyOperationsAborted;
            public IntPtr hNameMappings;
            public string lpszProgressTitle;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHFileOperation(ref SHFILEOPSTRUCT lpFileOp);

        private const uint FO_DELETE = 0x0003;
        private const ushort FOF_ALLOWUNDO = 0x0040;
        private const ushort FOF_NOCONFIRMATION = 0x0010;
        private const ushort FOF_NOERRORUI = 0x0400;
        private const ushort FOF_SILENT = 0x0004;

        // Flags für die SHEmptyRecycleBin-Funktion
        private const uint SHERB_NOCONFIRMATION = 0x00000001;  // Kein Bestätigungsdialog
        private const uint SHERB_NOPROGRESSUI = 0x00000002;    // Keine Fortschrittsanzeige
        private const uint SHERB_NOSOUND = 0x00000004;         // Kein Sound nach dem Löschen


        private readonly List<DirectoryInfo> protectedDirectories;
        private readonly List<DirectoryInfo> allowedSubdirectories;

        private List<DirectoryInfo> cleaningDirectories;

        private bool isErrorTextBoxVisible = false;
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
        private bool generateReport = false;
        private string reportPath = string.Empty;
        private string reportFormat = "HTML";
        private bool emailReport = false;

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

        public DeleteForm() {
            protectedDirectories = new List<DirectoryInfo>();
            allowedSubdirectories = new List<DirectoryInfo>();

            LoadDirectoriesFromConfig(Properties.Settings.Default.protectedDirectories, protectedDirectories);
            LoadDirectoriesFromConfig(Properties.Settings.Default.allowedSubdirectories, allowedSubdirectories);

            this.emptyTrash = Properties.Settings.Default.emptyTrash;
            this.minFileAgeDays = Properties.Settings.Default.minFileAgeDays;
            this.removeWindowsOld = Properties.Settings.Default.removeWindowsOld;
            this.flushDnsCache = Properties.Settings.Default.flushDnsCache;
            this.clearEventLogs = Properties.Settings.Default.clearEventLogs;
            this.healthCheckEnabled = Properties.Settings.Default.healthCheckEnabled;
            this.minFreeDiskPercent = Properties.Settings.Default.minFreeDiskPercent;

            this.dryRun = Properties.Settings.Default.dryRun;
            this.createRestorePoint = Properties.Settings.Default.createRestorePoint;
            this.deleteToRecycleBin = Properties.Settings.Default.deleteToRecycleBin;
            this.cleanBrowserCaches = Properties.Settings.Default.cleanBrowserCaches;
            this.dismComponentCleanup = Properties.Settings.Default.dismComponentCleanup;
            this.cleanMinidumps = Properties.Settings.Default.cleanMinidumps;
            this.cleanPrefetch = Properties.Settings.Default.cleanPrefetch;
            this.cleanDeliveryOptimization = Properties.Settings.Default.cleanDeliveryOptimization;
            this.cleanOldLogs = Properties.Settings.Default.cleanOldLogs;
            this.oldLogAgeDays = Properties.Settings.Default.oldLogAgeDays;
            this.emptyAllRecycleBins = Properties.Settings.Default.emptyAllRecycleBins;
            this.generateReport = Properties.Settings.Default.generateReport;
            this.reportPath = Properties.Settings.Default.reportPath;
            this.reportFormat = Properties.Settings.Default.reportFormat;
            this.emailReport = Properties.Settings.Default.emailReport;

            InitializeComponent();

            ApplyUserTheme();

            StyleRoundedButtons();

            // Wert-Label direkt hinter die Beschriftung "Vorgang:" setzen,
            // damit der Abstand unabhängig von Schriftbreite/DPI konstant bleibt.
            label1.Left = label5.Right + 2;
            directoryValueLabel.Left = label2.Right + 2;

            // Detailbereich ist anfangs ausgeblendet – Form entsprechend zusammenklappen.
            this.Height -= detailTextBox.Height + 10;

            this.FormClosing += DeleteForm_FormClosing;
        }

        // Farbpalette, die sich an den Windows-Benutzereinstellungen (Hell-/Dunkel-Modus) orientiert.
        private Color themeBackground;
        private Color themeSurface;
        private Color themeText;
        private Color themeSubtleText;
        private Color themeSeparator;
        private Color themePrimary;
        private Color themePrimaryHover;
        private Color themePrimaryText;
        private Color themeSecondary;
        private Color themeSecondaryHover;
        private Color themeSecondaryText;
        private Color themeLink;
        private Color themeLinkActive;
        private Color themeError;
        private Color themeSuccess;
        private static bool IsWindowsLightTheme() {
            try {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")) {
                    if (key?.GetValue("AppsUseLightTheme") is int value) {
                        return value != 0;
                    }
                }
            } catch {
                // Bei Fehlern (z. B. Registry nicht lesbar) hellen Modus annehmen.
            }
            return true;
        }

        private void ApplyUserTheme() {
            bool isLight = IsWindowsLightTheme();
            if (isLight) {
                themeBackground = Color.White;
                themeSurface = Color.FromArgb(249, 249, 249);
                themeText = Color.FromArgb(32, 32, 32);
                themeSubtleText = Color.FromArgb(96, 96, 96);
                themeSeparator = Color.FromArgb(229, 229, 229);
                themePrimary = Color.FromArgb(51, 51, 51);
                themePrimaryHover = Color.FromArgb(77, 77, 77);
                themePrimaryText = Color.White;
                themeSecondary = Color.FromArgb(243, 243, 243);
                themeSecondaryHover = Color.FromArgb(229, 229, 229);
                themeSecondaryText = Color.FromArgb(51, 51, 51);
                themeLink = Color.FromArgb(89, 89, 89);
                themeLinkActive = Color.FromArgb(51, 51, 51);
                themeError = Color.FromArgb(197, 48, 38);
                themeSuccess = Color.FromArgb(31, 136, 61);
            } else {
                themeBackground = Color.FromArgb(32, 32, 32);
                themeSurface = Color.FromArgb(43, 43, 43);
                themeText = Color.FromArgb(240, 240, 240);
                themeSubtleText = Color.FromArgb(170, 170, 170);
                themeSeparator = Color.FromArgb(64, 64, 64);
                themePrimary = Color.FromArgb(80, 80, 80);
                themePrimaryHover = Color.FromArgb(104, 104, 104);
                themePrimaryText = Color.White;
                themeSecondary = Color.FromArgb(56, 56, 56);
                themeSecondaryHover = Color.FromArgb(72, 72, 72);
                themeSecondaryText = Color.FromArgb(240, 240, 240);
                themeLink = Color.FromArgb(180, 180, 180);
                themeLinkActive = Color.FromArgb(240, 240, 240);
                themeError = Color.FromArgb(255, 123, 114);
                themeSuccess = Color.FromArgb(63, 185, 80);
            }

            this.BackColor = themeBackground;

            headerPanel.BackColor = themeBackground;
            headerTitleLabel.ForeColor = themeText;
            headerSeparatorPanel.BackColor = themeSeparator;
            separatorPanel.BackColor = themeSeparator;
            footerPanel.BackColor = themeSurface;

            label1.ForeColor = themeText;
            label2.ForeColor = themeSubtleText;
            directoryValueLabel.ForeColor = themeText;
            label5.ForeColor = themeSubtleText;
            label3.ForeColor = themeSubtleText;
            label4.ForeColor = themeSubtleText;

            linkLabel1.LinkColor = themeLink;
            linkLabel1.ActiveLinkColor = themeLinkActive;
            linkLabel2.LinkColor = themeLink;
            linkLabel2.ActiveLinkColor = themeLinkActive;

            detailTextBox.BackColor = themeSurface;
            detailTextBox.ForeColor = themeText;

            ApplyNativeControlTheme(isLight);
        }

        // Weist nativen Win32-Controls (ProgressBar, TextBox-Scrollbars) im Dunkelmodus
        // das passende Theme zu, da diese BackColor ignorieren.
        private void ApplyNativeControlTheme(bool isLight) {
            // TextBox: dunkle bzw. helle Scrollbars über das Explorer-Theme.
            string scrollTheme = isLight ? "Explorer" : "DarkMode_Explorer";
            ApplyWindowTheme(detailTextBox, scrollTheme, null);

            // ProgressBars: im Dunkelmodus das visuelle Theme entfernen, damit
            // BackColor/ForeColor überhaupt wirken; im Hellmodus Standard-Theme beibehalten.
            progressBar1.BackColor = themeSurface;
            progressBar2.BackColor = themeSurface;
            progressBar1.ForeColor = isLight ? Color.FromArgb(0, 120, 215) : Color.FromArgb(76, 194, 255);
            progressBar2.ForeColor = progressBar1.ForeColor;

            if (isLight) {
                ApplyWindowTheme(progressBar1, "Explorer", null);
                ApplyWindowTheme(progressBar2, "Explorer", null);
            } else {
                ApplyWindowTheme(progressBar1, "", "");
                ApplyWindowTheme(progressBar2, "", "");
            }
        }

        private void ApplyWindowTheme(Control control, string appName, string idList) {
            if (control == null) {
                return;
            }
            if (!control.IsHandleCreated) {
                control.HandleCreated += (s, e) => SetWindowTheme(((Control)s).Handle, appName, idList);
            } else {
                SetWindowTheme(control.Handle, appName, idList);
            }
        }

        private void StyleRoundedButtons() {
            ApplyRoundedButton(toggleButton,
                themePrimary,
                themePrimaryHover,
                themePrimaryText);

            ApplyRoundedButton(retryButton,
                themeSecondary,
                themeSecondaryHover,
                themeSecondaryText);
        }

        private void ApplyRoundedButton(Button button, Color normal, Color hover, Color foreColor) {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = hover;
            button.FlatAppearance.MouseDownBackColor = hover;
            button.ForeColor = foreColor;
            button.BackColor = normal;
            button.UseVisualStyleBackColor = false;
            button.Cursor = Cursors.Hand;

            // Kantengeglättetes Zeichnen statt harter Region-Beschneidung,
            // damit die abgerundeten Ecken sauber (nicht ausgefranst) wirken.
            button.Region = null;
            button.Paint += (s, e) => PaintRoundedButton(button, e, normal, hover, foreColor);
            button.MouseEnter += (s, e) => button.Invalidate();
            button.MouseLeave += (s, e) => button.Invalidate();
            button.EnabledChanged += (s, e) => button.Invalidate();
        }

        private void PaintRoundedButton(Button button, PaintEventArgs e, Color normal, Color hover, Color foreColor) {
            const int radius = 8;
            var g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(button.Parent != null ? button.Parent.BackColor : BackColor);

            bool isHover = button.ClientRectangle.Contains(button.PointToClient(Cursor.Position));
            Color fill = (button.Enabled && isHover) ? hover : normal;
            Color text = button.Enabled ? foreColor : Color.FromArgb(160, 160, 160);

            var rect = new Rectangle(0, 0, button.Width - 1, button.Height - 1);
            using (var path = CreateRoundedPath(rect, radius))
            using (var brush = new SolidBrush(fill)) {
                g.FillPath(brush, path);
            }

            TextRenderer.DrawText(g, button.Text, button.Font, button.ClientRectangle, text,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        private System.Drawing.Drawing2D.GraphicsPath CreateRoundedPath(Rectangle bounds, int radius) {
            int d = radius * 2;
            var path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
            path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
            path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        private void DeleteForm_FormClosing(object sender, FormClosingEventArgs e) {
            // Sicherstellen, dass die Windows-Update-Dienste beim Beenden wieder laufen,
            // auch wenn die Anwendung während der Bereinigung geschlossen wird.
            if (updateServicesStopped) {
                SetWindowsUpdateServices(true);
            }
        }

        private void DeleteForm_Shown(object sender, EventArgs e) {
            LoadConfiguration();
            Task.Run(() => CleanDirectories());

            if (healthCheckEnabled) {
                healthCheckTask = Task.Run(() => RunHealthCheck());
            }
        }

        private void LoadConfiguration() {
            EnsureHandleCreated();

            cleaningDirectories = new List<DirectoryInfo>();

            string directoriesConfig = Properties.Settings.Default.cleanDirectories;
            if (string.IsNullOrEmpty(directoriesConfig)) {
                MessageBox.Show("Keine Verzeichnisse in der Konfigurationsdatei gefunden.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            switch (path) {
                case "%TEMP%":
                    return Path.GetTempPath();
                case "%INTERNET_CACHE%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.InternetCache);
                case "%HISTORY%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.History);
                case "%DOWNLOADS%":
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
                case "%DESKTOP%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                case "%DOCUMENTS%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                case "%MUSIC%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.MyMusic);
                case "%PICTURES%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
                case "%VIDEOS%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.MyVideos);
                case "%WINDOWS%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.Windows);
                case "%SYSTEM%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.System);
                case "%PROGRAM_FILES%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                case "%PROGRAM_FILES_X86%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
                case "%USERPROFILE%":
                    return Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                case "%WINDOWS_TEMP%":
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Temp");
                case "%PREFETCH%":
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Prefetch");
                case "%SOFTWARE_DISTRIBUTION%":
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SoftwareDistribution", "Download");
                case "%CRASH_DUMPS%":
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CrashDumps");
                case "%THUMBNAIL_CACHE%":
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Windows", "Explorer");
                case "%EDGE_CACHE%":
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Edge", "User Data", "Default", "Cache");
                case "%CHROME_CACHE%":
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Google", "Chrome", "User Data", "Default", "Cache");
                case "%FIREFOX_CACHE%":
                    return ResolveFirefoxCache();
                default:
                    return path;
            }
        }

        private string ResolveFirefoxCache() {
            string firefoxProfiles = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Mozilla", "Firefox", "Profiles");
            if (!Directory.Exists(firefoxProfiles)) {
                return firefoxProfiles;
            }

            // Standardprofil (*.default-release) bevorzugen, sonst erstes Profil mit Cache verwenden.
            var profile = Directory.EnumerateDirectories(firefoxProfiles)
                .FirstOrDefault(p => p.EndsWith("default-release", StringComparison.OrdinalIgnoreCase))
                ?? Directory.EnumerateDirectories(firefoxProfiles).FirstOrDefault();

            return profile != null ? Path.Combine(profile, "cache2") : firefoxProfiles;
        }

        private bool IsProtectedDirectory(string directory) {
            if (allowedSubdirectories.Any(allowedDir => directory.StartsWith(allowedDir.FullName, StringComparison.OrdinalIgnoreCase))) {
                return false; // Verzeichnis ist erlaubt, also nicht geschützt
            }

            return protectedDirectories.Any(protectedDir => directory.StartsWith(protectedDir.FullName, StringComparison.OrdinalIgnoreCase));
        }

        private void EmptyRecycleBin() {
            try {
                // Leert den Papierkorb ohne Bestätigung, ohne UI und ohne Sound
                int result = SHEmptyRecycleBin(IntPtr.Zero, null, SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND);

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
                var issues = HealthCheck.Run(minFreeDiskPercent);

                if (issues.Count == 0) {
                    LogToDetailTextBox("Systemprüfung: keine Auffälligkeiten.");
                    return;
                }

                foreach (var issue in issues) {
                    LogToDetailTextBox($"Systemprüfung: {issue}");
                }

                var critical = issues.Where(i => i.Severity == HealthSeverity.Critical).ToList();
                if (critical.Count > 0) {
                    severeHealthIssueDetected = true;
                    NotifySevereIssues(critical);
                }
            } catch (Exception ex) {
                LogError(ex);
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
            try {
                Invoke(new Action(() => {
                    notifyIcon1.Visible = true;
                    notifyIcon1.BalloonTipTitle = title;
                    notifyIcon1.BalloonTipText = body;
                    notifyIcon1.BalloonTipIcon = ToolTipIcon.Error;
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
                    UpdateStatusLabel(label1, "Entferne Windows.old...", themeText);
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
                    LogToDetailTextBox($"Windows.old entfernt: {FormatBytes(size)}");
                } else {
                    LogToDetailTextBox("Windows.old konnte nicht vollständig entfernt werden (Administratorrechte erforderlich).");
                }
            } catch (Exception ex) {
                LogError(ex);
            }
        }

        private void FlushDnsCache() {
            try {
                Invoke(new Action(() => UpdateStatusLabel(label1, "Leere DNS-Cache...", themeText)));

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
            try {
                Invoke(new Action(() => UpdateStatusLabel(label1, "Leere Windows-Event-Logs...", themeText)));

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
                        $"Dienst {service} wird {(start ? "gestartet" : "gestoppt")}...", themeText)));

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
                Invoke(new Action(() => UpdateStatusLabel(label1, "Erstelle Systemwiederherstellungspunkt...", themeText)));
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

        private async Task CleanDirectories() {
            try {
                Invoke(new Action(() => progressBar2.Maximum = cleaningDirectories.Count));

                if (dryRun) {
                    Invoke(new Action(() => UpdateStatusLabel(label1, "Vorschau-Modus – es wird nichts gelöscht.", themeText)));
                    LogToDetailTextBox("=== VORSCHAU-MODUS (Dry-Run) – es werden keine Daten gelöscht ===");
                }

                // Optional einen Systemwiederherstellungspunkt anlegen, bevor größere
                // Löschvorgänge beginnen (nicht im Vorschau-Modus).
                if (createRestorePoint && !dryRun) {
                    await CreateRestorePointAsync();
                }

                // Wenn der Windows-Update-Download-Ordner bereinigt werden soll, müssen die
                // sperrenden Dienste (Windows Update, BITS) zuvor gestoppt werden.
                bool needsUpdateServices = cleaningDirectories.Any(d =>
                    d.FullName.IndexOf("SoftwareDistribution", StringComparison.OrdinalIgnoreCase) >= 0);

                if (needsUpdateServices) {
                    SetWindowsUpdateServices(false);
                }

                try {
                    foreach (var scanDir in cleaningDirectories) {
                        Invoke(new Action(() => UpdateStatusLabel(directoryValueLabel, $"{scanDir}")));
                        await ProcessDirectory(scanDir);
                        Invoke(new Action(() => progressBar2.PerformStep()));
                    }
                } finally {
                    if (needsUpdateServices) {
                        SetWindowsUpdateServices(true);
                    }
                }

                if (emptyTrash) {
                    Invoke(new Action(() => {
                        UpdateStatusLabel(directoryValueLabel, $"Papierkorb");
                        UpdateStatusLabel(label1, "Leere Papierkorb....", themeText);
                        EmptyRecycleBin();
                        UpdateStatusLabel(label1, "Leere Papierkorb... beendet", themeText);
                    }));
                }

                if (removeWindowsOld) {
                    await RemoveWindowsOldAsync();
                }

                if (cleanBrowserCaches) {
                    await CleanBrowserCachesAsync();
                }

                if (cleanMinidumps) {
                    await CleanMinidumpsAsync();
                }

                if (cleanPrefetch) {
                    await CleanPrefetchAsync();
                }

                if (cleanDeliveryOptimization) {
                    await CleanDeliveryOptimizationAsync();
                }

                if (cleanOldLogs) {
                    await CleanOldLogsAsync();
                }

                if (emptyAllRecycleBins) {
                    EmptyAllRecycleBins();
                }

                if (dismComponentCleanup && !dryRun) {
                    await RunDismComponentCleanupAsync();
                }

                if (flushDnsCache) {
                    FlushDnsCache();
                }

                if (clearEventLogs) {
                    ClearEventLogs();
                }

                Thread.Sleep(2500);
                CompleteProcess();
            } catch (Exception ex) {
                LogError(ex);
            }
        }

        private async Task ProcessDirectory(DirectoryInfo dir) {
            try {
                var files = dir.EnumerateFiles("*", SearchOption.AllDirectories).ToList();

                // Verzeichnisse nach Tiefe sortieren (tiefste zuerst), damit untergeordnete
                // Ordner vor den übergeordneten gelöscht werden.
                var directories = dir.EnumerateDirectories("*", SearchOption.AllDirectories)
                    .OrderByDescending(d => d.FullName.Count(c => c == Path.DirectorySeparatorChar))
                    .ToList();

                Invoke(new Action(() => { progressBar1.Maximum = files.Count + directories.Count; }));

                // Dateien parallel löschen (unabhängig voneinander), Verzeichnisse danach
                // sequenziell in Tiefensortierung (Reihenfolge ist hier relevant).
                await DeleteFilesParallelAsync(files);
                await DeleteItemsAsync(directories);
            } catch (UnauthorizedAccessException ex) {
                LogToDetailTextBox($"Zugriff verweigert auf {dir.FullName}: {ex.Message}");
            } catch (Exception ex) {
                LogError(ex);
            }
        }

        private async Task DeleteItemsAsync<T>(IEnumerable<T> items) where T : FileSystemInfo {
            foreach (var item in items) {
                string itemName = item.FullName;

                if (ShouldSkipDueToAge(item)) {
                    LogToDetailTextBox($"Übersprungen (zu jung): {itemName}");
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
                    LogToDetailTextBox($"Gelöscht: {itemName}");
                }

                QueueProgressStep(itemName);
            }

            FlushUiUpdates(itemName: null, force: true);
        }

        private async Task DeleteFilesParallelAsync(IReadOnlyList<FileInfo> files) {
            if (files.Count == 0) {
                return;
            }

            // Parallelitätsgrad an die CPU koppeln, aber begrenzen, um das Dateisystem
            // nicht zu überlasten. Bei sehr wenigen Dateien lohnt Parallelität nicht.
            int maxParallel = Math.Min(Environment.ProcessorCount * 2, 8);
            if (files.Count < 2 || maxParallel <= 1) {
                await DeleteItemsAsync(files);
                return;
            }

            using (var throttler = new System.Threading.SemaphoreSlim(maxParallel)) {
                var tasks = new List<Task>(files.Count);

                foreach (var file in files) {
                    await throttler.WaitAsync();

                    tasks.Add(Task.Run(async () => {
                        try {
                            await DeleteSingleFileAsync(file);
                        } finally {
                            throttler.Release();
                        }
                    }));
                }

                await Task.WhenAll(tasks);
            }

            FlushUiUpdates(itemName: null, force: true);
        }

        private async Task DeleteSingleFileAsync(FileInfo file) {
            string itemName = file.FullName;

            if (ShouldSkipDueToAge(file)) {
                LogToDetailTextBox($"Übersprungen (zu jung): {itemName}");
                QueueProgressStep();
                return;
            }

            long itemSize = GetFileLength(file);

            bool isDeleted = await TryDeleteItemAsync(file);

            if (isDeleted) {
                System.Threading.Interlocked.Add(ref freedBytes, itemSize);
                LogToDetailTextBox($"Gelöscht: {itemName}");
            }

            QueueProgressStep(itemName);
        }

        private void QueueProgressStep(string statusText = null) {
            int steps = System.Threading.Interlocked.Increment(ref pendingProgressSteps);

            // UI nur periodisch aktualisieren, um blockierende Invoke-Aufrufe pro Element zu vermeiden.
            if (uiThrottle.ElapsedMilliseconds >= UiUpdateIntervalMs) {
                FlushUiUpdates(statusText, force: false);
            }
        }

        private void FlushUiUpdates(string itemName, bool force) {
            int steps = System.Threading.Interlocked.Exchange(ref pendingProgressSteps, 0);
            uiThrottle.Restart();

            if (steps <= 0 && itemName == null && !force) {
                FlushDetailLog(force);
                return;
            }

            Action update = () => {
                int target = Math.Min(progressBar1.Value + steps, progressBar1.Maximum);
                if (target > progressBar1.Value) progressBar1.Value = target;
                if (itemName != null) label1.Text = itemName;
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
                DateTime threshold = DateTime.Now.AddDays(-minFileAgeDays);
                return file.LastWriteTime > threshold && file.LastAccessTime > threshold;
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
            const int maxAttempts = 3;
            string itemName = item.FullName;
            bool isFile = item is FileInfo;
            bool recursive = minFileAgeDays <= 0;
            Exception lastException = null;

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

            UpdateStatusLabel(label1, $"{itemName}", themeError);
            LogToDetailTextBox($"Löschen von {itemName} fehlgeschlagen: {lastException?.Message}");
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
                var fileOp = new SHFILEOPSTRUCT {
                    wFunc = FO_DELETE,
                    // pFrom muss mit einem zusätzlichen Nullzeichen terminiert sein (doppelt null).
                    pFrom = path + '\0' + '\0',
                    fFlags = (ushort)(FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT)
                };

                int result = SHFileOperation(ref fileOp);
                return result == 0 && fileOp.fAnyOperationsAborted == 0;
            } catch (Exception ex) {
                LogToDetailTextBox($"Papierkorb-Verschiebung fehlgeschlagen für {path}: {ex.Message}");
                return false;
            }
        }

        private void CompleteProcess() {
            string freedText = $"Freigegebener Speicher: {FormatBytes(freedBytes)}";
            LogToDetailTextBox(freedText);

            // Verbleibende gepufferte Log-Zeilen synchron in die TextBox schreiben,
            // bevor deren Inhalt auf Fehler geprüft wird.
            FlushUiUpdates(itemName: null, force: true);

            Invoke(new Action(() => {
                progressBar1.Value = progressBar1.Maximum = 1;
                progressBar2.Value = progressBar2.Maximum = 1;
                retryButton.Enabled = !IsAdministrator();
                bool withErrors = detailTextBox.Text
                    .Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries)
                    .Any(line => line.StartsWith("Löschen von", StringComparison.Ordinal) && line.Contains("fehlgeschlagen"));
                UpdateStatusLabel(label1, (withErrors
                    ? "Mit Fehlern abgeschlossen. Siehe Details."
                    : "Erfolgreich abgeschlossen.") + $" {freedText}", withErrors ? themeError : themeSuccess);
            }));

            Thread.Sleep(2500);  // Simulate delay

            // Sicherstellen, dass der parallele Health-Check abgeschlossen ist,
            // bevor über das Schließen entschieden wird.
            try {
                healthCheckTask?.Wait(TimeSpan.FromSeconds(30));
            } catch (Exception ex) {
                System.Diagnostics.Debug.WriteLine($"HealthCheck-Wait: {ex.Message}");
            }

            if (severeHealthIssueDetected) {
                Invoke(new Action(() => {
                    UpdateStatusLabel(label1, "Schwerwiegender Systemfehler erkannt – Programm bleibt geöffnet. Siehe Details.", themeError);
                    if (!isErrorTextBoxVisible) {
                        isErrorTextBoxVisible = true;
                        this.Height += detailTextBox.Height + 10;
                    }
                    detailTextBox.Visible = true;
                    toggleButton.Text = "Beenden";
                    retryButton.Enabled = !IsAdministrator();
                }));
                return; // Kein automatisches Schließen bei schwerem Fehler.
            }

            if (toggleButton.Text == "Details verbergen")
                Invoke(new Action(() => toggleButton.Text = "Beenden"));
            else {
                Thread.Sleep(5000);  // Simulate delay
                Application.Exit();
            }
        }

        private static string FormatBytes(long bytes) {
            string[] units = { "B", "KB", "MB", "GB", "TB" };
            double size = bytes;
            int unit = 0;
            while (size >= 1024 && unit < units.Length - 1) {
                size /= 1024;
                unit++;
            }
            return $"{size:0.##} {units[unit]}";
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
                    MessageBox.Show($"Fehler beim Neustarten mit Admin-Rechten: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            } else {
                Task.Run(() => CleanDirectories());
            }
        }

        private void ToggleButton_Click(object sender, EventArgs e) {
            if (toggleButton.Text == "Beenden") {
                notifyIcon1.Visible = false;
                Application.Exit();
                return;
            }

            isErrorTextBoxVisible = !isErrorTextBoxVisible;

            int extra = detailTextBox.Height + 10;
            this.Height += isErrorTextBoxVisible ? extra : -extra;
            detailTextBox.Visible = isErrorTextBoxVisible;
            toggleButton.Text = isErrorTextBoxVisible ? "Details verbergen" : "Details anzeigen";
        }

        private void LinkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            OpenMailClient(linkLabel2.Text);
        }

        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            OpenWebPage(linkLabel1.Text);
        }

        private void OpenMailClient(string email) {
            try {
                System.Diagnostics.Process.Start($"mailto:{email}?subject=Anfrage");
            } catch (Exception ex) {
                MessageBox.Show($"Fehler beim Öffnen des Mailprogramms: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenWebPage(string url) {
            try {
                System.Diagnostics.Process.Start(url);
            } catch (Exception ex) {
                MessageBox.Show($"Fehler beim Öffnen der Webseite: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void EnsureHandleCreated() {
            if (!this.IsHandleCreated) {
                this.CreateHandle();
            }
        }
    }
}
