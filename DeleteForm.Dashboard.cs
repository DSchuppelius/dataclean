using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataClean {
    // Darstellungslogik des Dashboards: Seitenwechsel, Kennzahlen, Aktivitätstabelle,
    // Phasenliste und Pausenfunktion. Die eigentliche Bereinigung liegt in DeleteForm.cs.
    public partial class DeleteForm {
        private enum Page {
            Overview,
            Cleanup,
            Log,
            Settings
        }

        private enum ActivityStatus {
            Deleted,
            Skipped,
            Failed
        }

        private sealed class ActivityEntry {
            public ActivityStatus Status;
            public string Path;
            public long Size;
        }

        // Obergrenze im Speicher; das vollständige Protokoll steht weiterhin auf der Protokoll-Seite.
        private const int MaxActivityEntries = 5000;
        private const int ActivityTrimThreshold = 500;
        private const int DashboardRefreshMs = 200;

        private static readonly CultureInfo GermanCulture = CultureInfo.GetCultureInfo("de-DE");

        private readonly List<ActivityEntry> pendingActivity = new List<ActivityEntry>();
        private readonly object activityLock = new object();
        private readonly List<ActivityEntry> activityEntries = new List<ActivityEntry>();
        private readonly List<ActivityEntry> visibleActivity = new List<ActivityEntry>();

        private Timer dashboardTimer;
        private UI.NavItem navOverview;
        private UI.NavItem navCleanup;
        private UI.NavItem navLog;
        private UI.NavItem navSettings;
        private Page currentPage = Page.Overview;

        private volatile bool isPaused;
        private volatile bool keepOpenAfterRun;

        // Fensterhöhe: ohne Details bleibt die Übersicht kompakt, alle anderen Ansichten
        // brauchen die volle Höhe. Die vom Benutzer gezogene Höhe wird wiederhergestellt.
        private const int ExpandedMinimumHeight = 640;
        private volatile bool detailsExpanded;
        private bool applyingWindowLayout;
        private bool windowIsExpanded = true;
        private int compactWindowHeight;
        private int expandedWindowHeight;

        private int directoriesDone;
        private int directoriesTotal;
        private double overallFraction;

        // Farben der Statuszeile unter der Überschrift.
        private Color themeStatusText;
        private Color themeSuccess;
        private Color themeError;

        private bool StayOpenAfterRun {
            get { return keepOpenAfterRun || detailsExpanded; }
        }

        private string RunningTitle {
            get { return dryRun ? "Vorschau läuft" : "Bereinigung läuft"; }
        }

        private void InitializeDashboard() {
            ApplyStatusColors();

            FooterControl = footerPanel;
            activityFilterCombo.SelectedIndex = 0;

            if (dryRun) {
                statFreed.Caption = "Freigebbar";
                statDeleted.Caption = "Würden gelöscht";
                activityFilterCombo.Items[1] = "Vorschau";
            }
            headerTitleLabel.Text = RunningTitle;

            dashboardTimer = new Timer(components) { Interval = DashboardRefreshMs };
            dashboardTimer.Tick += (s, e) => RefreshDashboard();

            // Das Design kann geändert worden sein – geöffnete Seiten direkt anpassen.
            settingsPage.Saved += (s, e) => ReapplyTheme();
        }

        private void ApplyStatusColors() {
            themeStatusText = UI.AppTheme.SubtleText;
            themeSuccess = UI.AppTheme.Success;
            themeError = UI.AppTheme.Error;
        }

        protected override void OnThemeReapplied() {
            ApplyStatusColors();
        }

        protected override void OnLoad(EventArgs e) {
            base.OnLoad(e);
            navOverview = AddNavItem(UI.IconKind.Home, "Übersicht", (s, ev) => ShowPage(Page.Overview));
            navCleanup = AddNavItem(UI.IconKind.Trash, "Bereinigung", (s, ev) => ShowPage(Page.Cleanup));
            navLog = AddNavItem(UI.IconKind.Document, "Protokoll", (s, ev) => ShowPage(Page.Log));
            navSettings = AddNavItem(UI.IconKind.Gear, "Einstellungen", (s, ev) => ShowPage(Page.Settings));

            // Die Designgröße enthält die Aktivitätstabelle; kompakt entfällt sie samt Abstand.
            expandedWindowHeight = Height;
            compactWindowHeight = Height - (actionsPanel.Top - activityCard.Top);

            ShowPage(Page.Overview);
            CenterToScreen();
        }

        private UI.NavItem NavItemFor(Page page) {
            switch (page) {
                case Page.Cleanup: return navCleanup;
                case Page.Log: return navLog;
                case Page.Settings: return navSettings;
                default: return navOverview;
            }
        }

        private void ShowPage(Page page) {
            if (currentPage == Page.Settings && page != Page.Settings && !settingsPage.ConfirmLeave()) {
                SelectNavItem(navSettings);
                return;
            }

            currentPage = page;

            // Beim Lesen oder Bearbeiten nicht automatisch schließen.
            if (page == Page.Log || page == Page.Settings) {
                keepOpenAfterRun = true;
            }
            if (page == Page.Log) {
                FlushDetailLog();
            }
            if (page == Page.Settings) {
                settingsPage.EnsureLoaded();
            }

            overviewPage.Visible = page == Page.Overview;
            cleanupPage.Visible = page == Page.Cleanup;
            logPage.Visible = page == Page.Log;
            settingsPage.Visible = page == Page.Settings;
            SelectNavItem(NavItemFor(page));
            ApplyWindowLayout();
        }

        private void DetailsLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            detailsExpanded = !detailsExpanded;
            ApplyWindowLayout();
        }

        // Stimmt Aktivitätstabelle, Link und Fensterhöhe auf Seite und Fensterzustand ab.
        // Maximiert ist genug Platz vorhanden, dort bleibt die Tabelle immer sichtbar.
        private void ApplyWindowLayout() {
            if (applyingWindowLayout || compactWindowHeight <= 0 || WindowState == FormWindowState.Minimized) {
                return;
            }

            applyingWindowLayout = true;
            try {
                bool maximized = WindowState == FormWindowState.Maximized;
                bool expanded = maximized || detailsExpanded || currentPage != Page.Overview;

                activityCard.Visible = maximized || detailsExpanded;
                detailsLink.Visible = !maximized;
                detailsLink.Text = detailsExpanded ? "Details verbergen" : "Details anzeigen";
                LockVerticalResize = !expanded;

                if (maximized) {
                    return;
                }

                Rectangle area = Screen.FromControl(this).WorkingArea;
                if (expanded) {
                    MinimumSize = new Size(MinimumSize.Width, Math.Min(ExpandedMinimumHeight, area.Height));
                    if (!windowIsExpanded) {
                        Height = Math.Min(expandedWindowHeight, area.Height);
                        if (Bottom > area.Bottom) {
                            Top = Math.Max(area.Top, area.Bottom - Height);
                        }
                    }
                } else {
                    if (windowIsExpanded) {
                        expandedWindowHeight = Height;
                    }
                    MinimumSize = new Size(MinimumSize.Width, compactWindowHeight);
                    Height = compactWindowHeight;
                }
                windowIsExpanded = expanded;
            } finally {
                applyingWindowLayout = false;
            }
        }

        private void ExitApplication() {
            if (!settingsPage.ConfirmLeave()) {
                return;
            }
            notifyIcon1.Visible = false;
            Application.Exit();
        }

        private void HistoryButton_Click(object sender, EventArgs e) {
            using (var form = new HistoryForm()) {
                form.ShowDialog(this);
            }
        }

        private void FinishButton_Click(object sender, EventArgs e) {
            ExitApplication();
        }

        private void PauseButton_Click(object sender, EventArgs e) {
            isPaused = !isPaused;
            pauseButton.Text = isPaused ? "Fortsetzen" : "Pausieren";
            pauseButton.Icon = isPaused ? UI.IconKind.Play : UI.IconKind.Pause;
            headerTitleLabel.Text = isPaused ? (dryRun ? "Vorschau pausiert" : "Bereinigung pausiert") : RunningTitle;
            LogToDetailTextBox(isPaused ? "=== Bereinigung pausiert ===" : "=== Bereinigung fortgesetzt ===");
        }

        // Hält den Lösch-Hotpath an, solange pausiert ist. Bereits laufende Einzelschritte
        // (z. B. DISM) werden nicht unterbrochen; ein Abbruch beendet das Warten sofort.
        private async Task WaitIfPausedAsync(System.Threading.CancellationToken cancellationToken) {
            while (isPaused) {
                await Task.Delay(150, cancellationToken);
            }
        }

        // Setzt die Anzeige zu Beginn eines Laufs zurück und startet die periodische Aktualisierung.
        private void PrepareDashboardForRun() {
            isPaused = false;
            overallFraction = 0d;
            directoriesDone = 0;
            directoriesTotal = cleaningDirectories?.Count ?? 0;

            headerTitleLabel.Text = RunningTitle;
            progressBar1.Value = 0;
            SetOverallProgress(0d);
            phaseLabel.Text = string.Empty;

            pauseButton.Text = "Pausieren";
            pauseButton.Icon = UI.IconKind.Pause;
            pauseButton.Enabled = true;
            pauseButton.Visible = true;
            cancelButton.Enabled = true;
            cancelButton.Visible = true;
            finishButton.Visible = false;
            retryButton.Visible = false;

            PopulateRunTables();
            dashboardTimer.Start();
        }

        // Schaltet die Aktionsleiste auf "Lauf beendet" und aktualisiert die Kennzahlen ein letztes Mal.
        private void ShowFinishedState(string title) {
            dashboardTimer.Stop();
            isPaused = false;
            headerTitleLabel.Text = title;
            pauseButton.Visible = false;
            cancelButton.Visible = false;
            retryButton.Visible = !IsAdministrator();
            finishButton.Visible = true;
            RefreshDashboard();
        }

        private void RefreshDashboard() {
            FlushDetailLog();
            DrainActivity();

            statFreed.Value = FormatBytes(System.Threading.Interlocked.Read(ref freedBytes));
            int deleted = System.Threading.Volatile.Read(ref deletedItemCount);
            int skipped = System.Threading.Volatile.Read(ref skippedItemCount);
            int failed = System.Threading.Volatile.Read(ref failedItemCount);
            statDeleted.Value = deleted.ToString("N0", GermanCulture);
            statSkipped.Value = skipped.ToString("N0", GermanCulture);

            UpdateInfoLine(skipped, failed);
            UpdateOverallProgress();
        }

        private void UpdateInfoLine(int skipped, int failed) {
            if (skipped <= 0 && failed <= 0) {
                infoLine.Visible = false;
                return;
            }

            var parts = new List<string>();
            if (skipped > 0) {
                parts.Add(skipped == 1
                    ? "1 Element übersprungen (in Verwendung, kein Zugriff oder Filter)."
                    : $"{skipped.ToString("N0", GermanCulture)} Elemente übersprungen (in Verwendung, kein Zugriff oder Filter).");
            }
            if (failed > 0) {
                parts.Add(failed == 1
                    ? "1 Element konnte nicht gelöscht werden."
                    : $"{failed.ToString("N0", GermanCulture)} Elemente konnten nicht gelöscht werden.");
            }

            infoLine.Tone = failed > 0 ? UI.Tone.Error : UI.Tone.Warning;
            infoLine.Text = string.Join(" ", parts);
            infoLine.Visible = true;
        }

        // Gesamtfortschritt über alle Phasen. Nur die erste Phase (Verzeichnisse) liefert
        // Zwischenwerte; die übrigen Phasen zählen beim Erreichen der nächsten Phase.
        private void UpdateOverallProgress() {
            if (totalPhases <= 0 || currentPhase <= 0) {
                return;
            }

            double phaseFraction = 0d;
            if (currentPhase == 1) {
                phaseFraction = directoriesTotal > 0
                    ? Math.Min(1d, (directoriesDone + progressBar1.Fraction) / directoriesTotal)
                    : 1d;
            }

            double overall = (currentPhase - 1 + phaseFraction) / totalPhases;
            overallFraction = Math.Max(overallFraction, overall);
            SetOverallProgress(overallFraction);
        }

        private void SetOverallProgress(double fraction) {
            progressBar2.Value = (int)Math.Round(fraction * progressBar2.Maximum);
            percentLabel.Text = $"{(int)Math.Floor(fraction * 100)} %";
        }

        // ---------------------------------------------------------------- Aktivitätstabelle

        // Nicht-blockierend puffern; die Tabelle wird periodisch auf dem UI-Thread befüllt.
        private void RecordActivity(ActivityStatus status, string path, long size) {
            var entry = new ActivityEntry { Status = status, Path = path, Size = size };
            lock (activityLock) {
                pendingActivity.Add(entry);
            }
        }

        private void RecordActivity(ActivityStatus status, FileSystemInfo item) {
            RecordActivity(status, item.FullName, item is FileInfo file ? GetFileLength(file) : 0);
        }

        private void DrainActivity() {
            List<ActivityEntry> batch;
            lock (activityLock) {
                if (pendingActivity.Count == 0) {
                    return;
                }
                batch = new List<ActivityEntry>(pendingActivity);
                pendingActivity.Clear();
            }

            activityEntries.AddRange(batch);
            if (activityEntries.Count > MaxActivityEntries + ActivityTrimThreshold) {
                activityEntries.RemoveRange(0, activityEntries.Count - MaxActivityEntries);
                RebuildVisibleActivity();
                return;
            }

            visibleActivity.AddRange(batch.Where(MatchesActivityFilter));
            UpdateActivityList();
        }

        private bool MatchesActivityFilter(ActivityEntry entry) {
            switch (activityFilterCombo.SelectedIndex) {
                case 1: return entry.Status == ActivityStatus.Deleted;
                case 2: return entry.Status == ActivityStatus.Skipped;
                case 3: return entry.Status == ActivityStatus.Failed;
                default: return true;
            }
        }

        private void RebuildVisibleActivity() {
            visibleActivity.Clear();
            visibleActivity.AddRange(activityEntries.Where(MatchesActivityFilter));
            UpdateActivityList();
        }

        private void UpdateActivityList() {
            activityList.VirtualListSize = visibleActivity.Count;
            activityList.FitColumns();
            activityList.Invalidate();
        }

        private void ActivityFilterCombo_SelectedIndexChanged(object sender, EventArgs e) {
            RebuildVisibleActivity();
        }

        // Neueste Einträge stehen oben, damit der laufende Vorgang ohne Scrollen sichtbar bleibt.
        private void ActivityList_RetrieveVirtualItem(object sender, RetrieveVirtualItemEventArgs e) {
            int index = visibleActivity.Count - 1 - e.ItemIndex;
            if (index < 0 || index >= visibleActivity.Count) {
                e.Item = new ListViewItem(new[] { string.Empty, string.Empty, string.Empty });
                return;
            }

            ActivityEntry entry = visibleActivity[index];
            UI.StatusKind status;
            string text;
            switch (entry.Status) {
                case ActivityStatus.Skipped:
                    status = UI.StatusKind.Skipped;
                    text = "Übersprungen";
                    break;
                case ActivityStatus.Failed:
                    status = UI.StatusKind.Error;
                    text = "Fehler";
                    break;
                default:
                    status = dryRun ? UI.StatusKind.Preview : UI.StatusKind.Success;
                    text = dryRun ? "Vorschau" : "Gelöscht";
                    break;
            }

            e.Item = new ListViewItem(new[] { text, entry.Path, entry.Size > 0 ? FormatBytes(entry.Size) : "–" }) {
                Tag = status
            };
        }

        // ---------------------------------------------------------------- Seite "Bereinigung"

        private void PopulateRunTables() {
            phasesList.BeginUpdate();
            phasesList.Items.Clear();
            foreach (string phase in GetEnabledPhaseNames()) {
                phasesList.Items.Add(CreateStatusRow(UI.StatusKind.Pending, "Ausstehend", phase));
            }
            phasesList.EndUpdate();

            directoriesList.BeginUpdate();
            directoriesList.Items.Clear();
            if (cleaningDirectories != null) {
                foreach (DirectoryInfo directory in cleaningDirectories) {
                    directoriesList.Items.Add(CreateStatusRow(UI.StatusKind.Pending, "Ausstehend", directory.FullName));
                }
            }
            directoriesList.EndUpdate();

            phasesList.FitColumns();
            directoriesList.FitColumns();
        }

        private static ListViewItem CreateStatusRow(UI.StatusKind status, string statusText, string name) {
            return new ListViewItem(new[] { statusText, name }) { Tag = status };
        }

        private static void SetRowStatus(ListView list, int index, UI.StatusKind status, string text) {
            if (index < 0 || index >= list.Items.Count) {
                return;
            }
            ListViewItem row = list.Items[index];
            row.Tag = status;
            row.Text = text;
        }

        private void MarkPhaseRunning(int phaseIndex) {
            for (int i = 0; i < phasesList.Items.Count && i <= phaseIndex; i++) {
                if (i == phaseIndex) {
                    SetRowStatus(phasesList, i, UI.StatusKind.Running, "Läuft");
                } else {
                    SetRowStatus(phasesList, i, UI.StatusKind.Success, "Fertig");
                }
            }
        }

        // Schließt beide Tabellen am Laufende ab: laufende Zeilen werden "Fertig" oder "Abgebrochen".
        private void FinishRunTables(bool canceled) {
            foreach (ListView list in new ListView[] { phasesList, directoriesList }) {
                for (int i = 0; i < list.Items.Count; i++) {
                    UI.StatusKind status = list.Items[i].Tag is UI.StatusKind kind ? kind : UI.StatusKind.None;
                    if (status == UI.StatusKind.Running) {
                        SetRowStatus(list, i, canceled ? UI.StatusKind.Warning : UI.StatusKind.Success, canceled ? "Abgebrochen" : "Fertig");
                    } else if (status == UI.StatusKind.Pending && !canceled) {
                        SetRowStatus(list, i, UI.StatusKind.Success, "Fertig");
                    }
                }
            }
        }
    }
}
