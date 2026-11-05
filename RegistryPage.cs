using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DataClean.Services;

namespace DataClean {
    /// <summary>
    /// Seite im Hauptfenster für die Registry-Prüfung. Sie läuft nur auf ausdrücklichen Klick und nie
    /// im unbeaufsichtigten Lauf: Dort gäbe es niemanden, der die Funde beurteilt.
    /// </summary>
    public partial class RegistryPage : UserControl, UI.IThemedControl {
        private readonly Dictionary<RegistryCategory, UI.ThemedCheckBox> categoryBoxes;

        // Funde, die nicht entfernt werden sollen. Bleibt über das Ein- und Ausblenden von Bereichen erhalten.
        private readonly HashSet<RegistryIssue> unselected = new HashSet<RegistryIssue>();

        private RegistryScope scope = RegistryScope.System;
        private string backupDirectory = RegistryCleaner.DefaultBackupDirectory;
        private List<RegistryIssue> issues = new List<RegistryIssue>();
        private CancellationTokenSource scanCancellation;
        private bool busy;
        private bool updatingChecks;
        private bool previewOnly;

        public RegistryPage() {
            InitializeComponent();

            categoryBoxes = new Dictionary<RegistryCategory, UI.ThemedCheckBox> {
                { RegistryCategory.StartupEntries, chkStartup },
                { RegistryCategory.ApplicationPaths, chkApplicationPaths },
                { RegistryCategory.ObsoleteSoftware, chkObsoleteSoftware },
                { RegistryCategory.Applications, chkApplications },
                { RegistryCategory.SharedDlls, chkSharedDlls },
                { RegistryCategory.Fonts, chkFonts },
                { RegistryCategory.HelpFiles, chkHelpFiles },
                { RegistryCategory.SoundEvents, chkSoundEvents },
                { RegistryCategory.MuiCache, chkMuiCache },
                { RegistryCategory.ComClasses, chkComClasses },
                { RegistryCategory.TypeLibraries, chkTypeLibraries }
            };
            foreach (UI.ThemedCheckBox box in categoryBoxes.Values) {
                box.CheckedChanged += CategoryBox_CheckedChanged;
            }

            Disposed += (s, e) => scanCancellation?.Cancel();

            ApplyTheme();
            lblSubheading.Text = DefaultSubtitle;
            UpdateActions();
        }

        /// <summary>Texte für das Protokoll der Anwendung.</summary>
        public event Action<string> LogMessage;

        /// <summary>Im Vorschau-Modus wird nur geprüft; entfernen lässt sich nichts.</summary>
        [System.ComponentModel.Browsable(false)]
        [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public bool PreviewOnly {
            get { return previewOnly; }
            set {
                if (previewOnly == value) {
                    return;
                }
                previewOnly = value;
                lblSubheading.Text = DefaultSubtitle;
                UpdateActions();
            }
        }

        public void ApplyTheme() {
            BackColor = UI.AppTheme.Background;
        }

        protected override void WndProc(ref Message m) {
            if (UI.EdgeTransparentPanel.HandleEdgeHitTest(this, ref m)) {
                return;
            }
            base.WndProc(ref m);
        }

        private string DefaultSubtitle {
            get {
                if (previewOnly) {
                    return "Vorschau-Modus – es wird nichts entfernt.";
                }
                return IsElevated()
                    ? "Findet Einträge, deren Datei nachweislich fehlt. Vor dem Entfernen wird gesichert."
                    : "Ohne Administratorrechte lassen sich nur Einträge des aktuellen Benutzers entfernen.";
            }
        }

        private static bool IsElevated() {
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent()) {
                return new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        private async void ScanButton_Click(object sender, EventArgs e) {
            // Während der Prüfung ist derselbe Knopf "Abbrechen".
            if (scanCancellation != null) {
                scanCancellation.Cancel();
                return;
            }

            List<RegistryCategory> categories = categoryBoxes.Where(pair => pair.Value.Checked).Select(pair => pair.Key).ToList();
            if (categories.Count == 0) {
                lblSubheading.Text = "Bitte mindestens einen Bereich auswählen.";
                return;
            }

            issues = new List<RegistryIssue>();
            unselected.Clear();
            PopulateList();
            lblHeading.Text = "Prüfung läuft";

            RegistryScope activeScope = scope;
            var progress = new Progress<RegistryCategory>(category => {
                if (!IsDisposed) {
                    lblSubheading.Text = "Prüfe " + RegistryCategories.Title(category) + " …";
                }
            });

            scanCancellation = new CancellationTokenSource();
            CancellationToken token = scanCancellation.Token;
            SetBusy(true);
            try {
                List<RegistryIssue> found = await Task.Run(() => new RegistryScanner(activeScope, new RegistryPathProbe()).Scan(categories, progress, token));
                if (IsDisposed) {
                    return;
                }
                issues = found.OrderBy(issue => issue.Category).ThenBy(issue => issue.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
                unselected.UnionWith(issues.Where(issue => RegistryCategories.IsCautious(issue.Category)));
                ShowFindings();
            } catch (OperationCanceledException) {
                if (!IsDisposed) {
                    lblHeading.Text = "Prüfung abgebrochen";
                    lblSubheading.Text = DefaultSubtitle;
                }
            } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is SecurityException || ex is ArgumentException) {
                Logger.Error(ex, "Registry");
                if (!IsDisposed) {
                    lblHeading.Text = "Prüfung fehlgeschlagen";
                    lblSubheading.Text = ex.Message;
                }
            } finally {
                scanCancellation.Dispose();
                scanCancellation = null;
                if (!IsDisposed) {
                    SetBusy(false);
                }
            }
        }

        private async void RemoveButton_Click(object sender, EventArgs e) {
            List<RegistryIssue> selected = SelectedIssues();
            if (selected.Count == 0 || previewOnly) {
                return;
            }

            int cautious = selected.Count(issue => RegistryCategories.IsCautious(issue.Category));
            string text = CountText(selected.Count, "Registry-Eintrag wird", "Registry-Einträge werden") + " entfernt. Vorher entsteht eine Sicherung, die sich hier wieder einspielen lässt."
                + (cautious > 0 ? $"{Environment.NewLine}{Environment.NewLine}{cautious} davon stammen aus den Bereichen „Mit Vorsicht“." : string.Empty)
                + $"{Environment.NewLine}{Environment.NewLine}Möchten Sie fortfahren?";
            DialogResult answer = UI.AppDialog.Show(FindForm(), text, "Registry bereinigen",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (answer != DialogResult.Yes) {
                return;
            }

            RegistryScope activeScope = scope;
            string activeDirectory = backupDirectory;
            lblHeading.Text = "Einträge werden entfernt";
            lblSubheading.Text = "Sicherung wird angelegt …";
            SetBusy(true);
            try {
                RegistryCleanResult result = await Task.Run(() => new RegistryCleaner(activeScope, new RegistryPathProbe(), activeDirectory).Clean(selected));
                if (!IsDisposed) {
                    ShowCleanResult(result);
                }
            } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is SecurityException || ex is NotSupportedException) {
                Logger.Error(ex, "Registry");
                if (!IsDisposed) {
                    lblHeading.Text = "Nichts entfernt";
                    lblSubheading.Text = "Die Sicherung ist fehlgeschlagen: " + ex.Message;
                    LogMessage?.Invoke("Registry: Sicherung fehlgeschlagen, nichts entfernt: " + ex.Message);
                }
            } finally {
                if (!IsDisposed) {
                    SetBusy(false);
                }
            }
        }

        private async void RestoreButton_Click(object sender, EventArgs e) {
            string backup = RegistryCleaner.FindLatestBackup(backupDirectory);
            if (backup == null) {
                return;
            }

            DialogResult answer = UI.AppDialog.Show(FindForm(),
                $"Die letzte Sicherung wird wieder eingespielt:{Environment.NewLine}{Path.GetFileName(backup)}{Environment.NewLine}{Environment.NewLine}Möchten Sie fortfahren?",
                "Sicherung einspielen", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (answer != DialogResult.Yes) {
                return;
            }

            lblHeading.Text = "Sicherung wird eingespielt";
            lblSubheading.Text = Path.GetFileName(backup);
            SetBusy(true);
            try {
                ProcessResult result = await Task.Run(() => RegistryBackup.Restore(backup));
                if (IsDisposed) {
                    return;
                }
                // Die Liste zeigt nun nicht mehr den aktuellen Stand.
                issues = new List<RegistryIssue>();
                unselected.Clear();
                PopulateList();
                if (result.ExitCode == 0) {
                    lblHeading.Text = "Sicherung eingespielt";
                    lblSubheading.Text = Path.GetFileName(backup);
                    LogMessage?.Invoke("Registry: Sicherung eingespielt: " + backup);
                } else {
                    lblHeading.Text = "Einspielen fehlgeschlagen";
                    lblSubheading.Text = "Für Einträge des gesamten Rechners sind Administratorrechte nötig. Die Sicherung bleibt erhalten.";
                    Logger.Warn($"Sicherung nicht eingespielt ({result.ExitCode}): {result.ErrorOutput}", "Registry");
                }
            } catch (Exception ex) when (ex is System.ComponentModel.Win32Exception || ex is InvalidOperationException) {
                // reg.exe ließ sich nicht starten, etwa weil eine Richtlinie es sperrt.
                Logger.Error(ex, "Registry");
                if (!IsDisposed) {
                    lblHeading.Text = "Einspielen fehlgeschlagen";
                    lblSubheading.Text = ex.Message;
                }
            } finally {
                if (!IsDisposed) {
                    SetBusy(false);
                }
            }
        }

        private void ShowFindings() {
            UpdateCategoryCounts();
            PopulateList();
            if (issues.Count == 0) {
                lblHeading.Text = "Keine verwaisten Einträge";
                lblSubheading.Text = "In den geprüften Bereichen zeigt kein Eintrag nachweislich ins Leere.";
            } else {
                lblHeading.Text = CountText(issues.Count, "verwaister Eintrag", "verwaiste Einträge");
                lblSubheading.Text = DefaultSubtitle;
            }
        }

        private void ShowCleanResult(RegistryCleanResult result) {
            issues = issues.Except(result.Removed).Except(result.Skipped).ToList();
            UpdateCategoryCounts();
            PopulateList();

            string removed = CountText(result.Removed.Count, "Eintrag", "Einträge") + " entfernt";
            lblHeading.Text = result.Failed.Count > 0 ? removed + ", " + result.Failed.Count + " fehlgeschlagen" : removed;
            if (result.Failed.Count > 0 && !IsElevated()) {
                lblSubheading.Text = "Für Einträge des gesamten Rechners sind Administratorrechte nötig.";
            } else {
                lblSubheading.Text = result.BackupPath != null
                    ? "Sicherung: " + result.BackupPath
                    : "Die Dateien sind inzwischen wieder vorhanden, es wurde nichts entfernt.";
            }

            LogMessage?.Invoke($"Registry: {result.Removed.Count} entfernt, {result.Failed.Count} fehlgeschlagen, {result.Skipped.Count} übersprungen. Sicherung: {result.BackupPath}");
            foreach (KeyValuePair<RegistryIssue, string> failure in result.Failed) {
                LogMessage?.Invoke($"Registry: nicht entfernt: {failure.Key.KeyPath}: {failure.Value}");
            }
        }

        private static string CountText(int count, string singular, string plural) {
            return count + " " + (count == 1 ? singular : plural);
        }

        private void UpdateCategoryCounts() {
            foreach (KeyValuePair<RegistryCategory, UI.ThemedCheckBox> pair in categoryBoxes) {
                int count = issues.Count(issue => issue.Category == pair.Key);
                pair.Value.Text = RegistryCategories.Title(pair.Key) + (count > 0 ? " (" + count + ")" : string.Empty);
            }
        }

        private void PopulateList() {
            updatingChecks = true;
            issuesList.BeginUpdate();
            try {
                issuesList.Items.Clear();
                foreach (RegistryIssue issue in issues.Where(item => categoryBoxes[item.Category].Checked)) {
                    var row = new ListViewItem(new[] { issue.Name, RegistryCategories.Title(issue.Category), issue.MissingPath }) {
                        Tag = issue,
                        Checked = !unselected.Contains(issue),
                        ToolTipText = scope.FullName(issue.Hive, issue.KeyPath)
                            + (issue.IsWholeKey ? string.Empty : Environment.NewLine + "Wert: " + issue.ValueName)
                            + Environment.NewLine + "Fehlt: " + issue.MissingPath
                    };
                    issuesList.Items.Add(row);
                }
            } finally {
                issuesList.EndUpdate();
                updatingChecks = false;
            }
            issuesList.FitColumns();
            UpdateActions();
        }

        private List<RegistryIssue> SelectedIssues() {
            return issuesList.Items.Cast<ListViewItem>().Where(row => row.Checked).Select(row => (RegistryIssue)row.Tag).ToList();
        }

        private void IssuesList_ItemChecked(object sender, ItemCheckedEventArgs e) {
            if (updatingChecks) {
                return;
            }
            var issue = (RegistryIssue)e.Item.Tag;
            if (e.Item.Checked) {
                unselected.Remove(issue);
            } else {
                unselected.Add(issue);
            }
            UpdateActions();
        }

        private void SelectAllCheck_CheckedChanged(object sender, EventArgs e) {
            if (updatingChecks) {
                return;
            }
            updatingChecks = true;
            issuesList.BeginUpdate();
            try {
                foreach (ListViewItem row in issuesList.Items) {
                    row.Checked = selectAllCheck.Checked;
                    var issue = (RegistryIssue)row.Tag;
                    if (selectAllCheck.Checked) {
                        unselected.Remove(issue);
                    } else {
                        unselected.Add(issue);
                    }
                }
            } finally {
                issuesList.EndUpdate();
                updatingChecks = false;
            }
            UpdateActions();
        }

        private void CategoryBox_CheckedChanged(object sender, EventArgs e) {
            if (issues.Count > 0) {
                PopulateList();
            }
        }

        private void SetBusy(bool value) {
            busy = value;
            progressBar.Visible = value;
            progressBar.Marquee = value;
            UpdateActions();
        }

        private void UpdateActions() {
            bool scanning = scanCancellation != null;
            int selected = issuesList.Items.Cast<ListViewItem>().Count(row => row.Checked);

            foreach (UI.ThemedCheckBox box in categoryBoxes.Values) {
                box.Enabled = !busy;
            }
            issuesList.Enabled = !busy;

            scanButton.Text = scanning ? "Abbrechen" : (issues.Count > 0 ? "Erneut prüfen" : "Prüfen");
            scanButton.Kind = issues.Count > 0 || scanning ? UI.ButtonKind.Secondary : UI.ButtonKind.Primary;
            scanButton.Enabled = !busy || scanning;

            removeButton.Visible = !previewOnly;
            removeButton.Text = selected > 0 ? CountText(selected, "Eintrag", "Einträge") + " entfernen" : "Auswahl entfernen";
            removeButton.Enabled = !busy && selected > 0;

            restoreButton.Visible = !previewOnly;
            restoreButton.Enabled = !busy && RegistryCleaner.FindLatestBackup(backupDirectory) != null;

            selectAllCheck.Visible = !previewOnly;
            selectAllCheck.Enabled = !busy && issuesList.Items.Count > 0;
            if (!updatingChecks) {
                updatingChecks = true;
                selectAllCheck.Checked = issuesList.Items.Count > 0 && selected == issuesList.Items.Count;
                updatingChecks = false;
            }
        }
    }
}
