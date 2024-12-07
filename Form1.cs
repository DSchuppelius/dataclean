using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataClean {
    public partial class Form1 : Form {
        // Definition der SHEmptyRecycleBin Funktion aus der shell32.dll
        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern int SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, uint dwFlags);

        // Flags für die SHEmptyRecycleBin-Funktion
        private const uint SHERB_NOCONFIRMATION = 0x00000001;  // Kein Bestätigungsdialog
        private const uint SHERB_NOPROGRESSUI = 0x00000002;    // Keine Fortschrittsanzeige
        private const uint SHERB_NOSOUND = 0x00000004;         // Kein Sound nach dem Löschen


        private readonly List<DirectoryInfo> protectedDirectories;
        private readonly List<DirectoryInfo> allowedSubdirectories;

        private List<DirectoryInfo> cleaningDirectories;

        private bool isErrorTextBoxVisible = false;
        private bool emptyTrash = false;

        public Form1() {
            protectedDirectories = new List<DirectoryInfo>();
            allowedSubdirectories = new List<DirectoryInfo>();

            LoadDirectoriesFromConfig(Properties.Settings.Default.protectedDirectories, protectedDirectories);
            LoadDirectoriesFromConfig(Properties.Settings.Default.allowedSubdirectories, allowedSubdirectories);

            this.emptyTrash = Properties.Settings.Default.emptyTrash;

            InitializeComponent();
        }

        private void Form1_Shown(object sender, EventArgs e) {
            LoadConfiguration();
            Task.Run(() => CleanDirectories());
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
                default:
                    return path;
            }
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

                if (result == 0) {
                    LogToDetailTextBox("Papierkorb erfolgreich geleert.");
                } else {
                    LogToDetailTextBox($"Fehler beim Leeren des Papierkorbs. Fehlercode: {result}");
                }
            } catch (Exception ex) {
                LogError(ex);
            }
        }

        private async Task CleanDirectories() {
            try {
                Invoke(new Action(() => progressBar2.Maximum = cleaningDirectories.Count));

                foreach (var scanDir in cleaningDirectories) {
                    Invoke(new Action(() => UpdateStatusLabel(label2, $"Aktuelles Verzeichnis: {scanDir}")));
                    await ProcessDirectory(scanDir);
                    Invoke(new Action(() => progressBar2.PerformStep()));
                }

                if (emptyTrash) {
                    Invoke(new Action(() => {
                        UpdateStatusLabel(label2, $"Aktuelles Verzeichnis: Papierkorb");
                        UpdateStatusLabel(label1, "Leere Papierkorb....", Color.Black);
                        EmptyRecycleBin();
                        UpdateStatusLabel(label1, "Leere Papierkorb... beendet", Color.Black);
                    }));
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
                var directories = dir.EnumerateDirectories("*", SearchOption.AllDirectories).ToList();

                Invoke(new Action(() => { progressBar1.Maximum = files.Count + directories.Count; }));

                await DeleteItemsAsync(files);
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
                bool isDeleted = false;

                try {
                    if (item is FileInfo) {
                        File.Delete(itemName);
                        isDeleted = true;
                    } else if (item is DirectoryInfo) {
                        Directory.Delete(itemName, true);
                        isDeleted = true;
                    }

                    if (isDeleted) {
                        LogToDetailTextBox($"Gelöscht: {itemName}");
                        UpdateStatusLabel(label1, $"{itemName}", Color.Black);
                    }
                } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) {
                    UpdateStatusLabel(label1, $"{itemName}", Color.Red);
                    LogToDetailTextBox($"Löschen von {itemName} fehlgeschlagen: {ex.Message}");
                    await Task.Delay(5); // Async Pause
                }

                Invoke(new Action(() => progressBar1.PerformStep()));
            }
        }

        private void CompleteProcess() {
            Invoke(new Action(() => {
                progressBar1.Value = progressBar1.Maximum = 1;
                progressBar2.Value = progressBar2.Maximum = 1;
                retryButton.Enabled = !IsAdministrator();
                UpdateStatusLabel(label1, string.IsNullOrEmpty(detailTextBox.Text)
                    ? "Erfolgreich abgeschlossen"
                    : "Mit Fehlern abgeschlossen. Siehe Details.", string.IsNullOrEmpty(detailTextBox.Text)
                        ? Color.Green
                        : Color.Red
                );
            }));

            Thread.Sleep(2500);  // Simulate delay

            if (toggleButton.Text == "Details verbergen")
                Invoke(new Action(() => toggleButton.Text = "Beenden"));
            else {
                Thread.Sleep(5000);  // Simulate delay
                Application.Exit();
            }
        }

        private void UpdateStatusLabel(Label label, string text, Color? color = null) {
            Invoke(new Action(() => {
                label.Text = text;
                if (color.HasValue) label.ForeColor = color.Value;
            }));
        }

        private void LogToDetailTextBox(string message) {
            Invoke(new Action(() => detailTextBox.AppendText($"{message}{Environment.NewLine}")));
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
            if (toggleButton.Text == "Beenden")
                Application.Exit();

            isErrorTextBoxVisible = !isErrorTextBoxVisible;
            detailTextBox.Visible = isErrorTextBoxVisible;
            this.Height += isErrorTextBoxVisible ? detailTextBox.Height + 5 : -(detailTextBox.Height + 5);
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
