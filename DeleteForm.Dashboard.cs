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
        private const int ResultVisibleMs = 2500;
        private const int AutoCloseDelayMs = 7500;
        private const int IdlePollMs = 250;
        private const int BusyElapsedAfterSeconds = 3;

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

        // Schritt ohne Dateifortschritt (z. B. DISM). Wird nur auf dem UI-Thread geändert.
        private string busyStatus;
        private DateTime busyStartedAt;
        private double busyPercent = -1d;
        private string defaultStatusText;

        // Von Hand gestartet: Das Fenster wartet auf den Start, es läuft noch nichts.
        private bool waitingForStart;

        // Farben der Statuszeile unter der Überschrift.
        private Color themeStatusText;
        private Color themeSuccess;
        private Color themeError;

        // Solange jemand Details oder eine andere Seite offen hat, wartet das automatische Schließen.
        private bool StayOpenAfterRun {
            get { return detailsExpanded || currentPage != Page.Overview; }
        }

        private string RunningTitle {
            get { return dryRun ? "Vorschau läuft" : "Bereinigung läuft"; }
        }

        private void InitializeDashboard() {
            ApplyStatusColors();
            defaultStatusText = label1.Text;

            FooterControl = footerPanel;

            // Das Logo steckt bereits als Anwendungssymbol in der Exe.
            using (var logo = new Icon(Icon, 128, 128)) {
                SidebarLogo = logo.ToBitmap();
            }
            using (var small = new Icon(Icon, 32, 32)) {
                UI.AppTheme.BrandImage = small.ToBitmap();
            }
            activityFilterCombo.SelectedIndex = 0;

            ApplyRunModeTexts();
            heroLabel.Text = HeroText;
            headerTitleLabel.Text = RunningTitle;

            dashboardTimer = new Timer(components) { Interval = DashboardRefreshMs };
            dashboardTimer.Tick += (s, e) => RefreshDashboard();

            settingsPage.Saved += (s, e) => {
                // Das Design kann geändert worden sein – geöffnete Seiten direkt anpassen.
                ReapplyTheme();
                // Wartet die Anwendung auf den Start, zeigt sie, was dieser jetzt tun würde.
                if (waitingForStart) {
                    ShowReadyState();
                }
            };
        }

        private void ApplyRunModeTexts() {
            statDeleted.Caption = dryRun ? "Würden gelöscht" : "Elemente gelöscht";
            activityFilterCombo.Items[1] = dryRun ? "Vorschau" : "Gelöscht";
        }

        private string HeroText {
            get { return FormatBytes(System.Threading.Interlocked.Read(ref freedBytes)) + (dryRun ? " freigebbar" : " freigegeben"); }
        }

        private void ApplyStatusColors() {
            themeStatusText = UI.AppTheme.SubtleText;
            themeSuccess = UI.AppTheme.Success;
            themeError = UI.AppTheme.Error;
        }

        protected override void OnThemeReapplied() {
            ApplyStatusColors();
        }

        // Die Palette setzt alle Beschriftungen auf ihre Grundfarbe; Erfolg und Fehler sollen bleiben.
        public override void ReapplyTheme() {
            Color title = headerTitleLabel.ForeColor;
            Color status = label1.ForeColor;
            Color oldSuccess = themeSuccess;
            Color oldError = themeError;

            base.ReapplyTheme();

            headerTitleLabel.ForeColor = StateColorAfterThemeChange(title, oldSuccess, oldError, UI.AppTheme.Link);
            label1.ForeColor = StateColorAfterThemeChange(status, oldSuccess, oldError, themeStatusText);
        }

        private Color StateColorAfterThemeChange(Color before, Color oldSuccess, Color oldError, Color neutral) {
            if (before == oldSuccess) {
                return themeSuccess;
            }
            return before == oldError ? themeError : neutral;
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

            // Schon vor dem ersten Zeichnen, damit nicht kurz "Bereinigung läuft" zu sehen ist.
            if (!unattended) {
                ShowReadyState();
            }
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
            if (settingsPage.ConfirmLeave()) {
                CloseApplication();
            }
        }

        // Nur der unbeaufsichtigte Lauf schließt sich selbst: Wer die Anwendung von Hand startet und
        // das Löschen bestätigt, beendet sie auch selbst. Unbeaufsichtigt schließt sie, sobald sie die
        // angegebene Zeit unbenutzt war. Wer Details oder eine andere Seite offen hat, liest in Ruhe;
        // danach läuft die Wartezeit neu an.
        private async Task CloseIfUnattendedAsync(int idleMs) {
            if (!unattended) {
                return;
            }

            int idle = 0;
            while (idle < idleMs || StayOpenAfterRun) {
                await Task.Delay(IdlePollMs);
                idle = StayOpenAfterRun ? 0 : idle + IdlePollMs;
            }
            CloseApplication();
        }

        // Hinweise der Systemprüfung stehen bis zum Schließen unter dem Fortschritt.
        private void ShowHealthWarnings() {
            List<HealthIssue> warnings = pendingHealthIssues?.Where(i => i.Severity == HealthSeverity.Warning).ToList();
            if (warnings == null || warnings.Count == 0) {
                return;
            }

            string more = warnings.Count == 2 ? " (+1 weiterer Hinweis)" : (warnings.Count > 2 ? $" (+{warnings.Count - 1} weitere Hinweise)" : string.Empty);
            infoLine.Tone = UI.Tone.Warning;
            infoLine.Text = warnings[0].Message + more;
            infoLine.Visible = true;
        }

        // Beendet die Anwendung auf dem UI-Thread und entfernt das Symbol im Infobereich, das sonst
        // nach einer Abschlussmeldung bis zum Überfahren mit der Maus stehen bliebe.
        private void CloseApplication() {
            if (IsDisposed || !IsHandleCreated) {
                return;
            }

            Action close = () => {
                notifyIcon1.Visible = false;
                Application.Exit();
            };
            if (InvokeRequired) {
                BeginInvoke(close);
            } else {
                close();
            }
        }

        private void HistoryButton_Click(object sender, EventArgs e) {
            using (var form = new HistoryForm()) {
                form.ShowDialog(this);
            }
        }

        private void FinishButton_Click(object sender, EventArgs e) {
            ExitApplication();
        }

        private void StartButton_Click(object sender, EventArgs e) {
            StartRun();
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

        // Setzt die Anzeige zurück, die ein vorheriger Lauf hinterlassen haben kann.
        private void ResetDashboard() {
            isPaused = false;
            overallFraction = 0d;
            directoriesDone = 0;
            directoriesTotal = cleaningDirectories?.Count ?? 0;

            ClearBusyStep();
            ApplyRunModeTexts();
            headerTitleLabel.ForeColor = UI.AppTheme.Link;
            heroLabel.Text = HeroText;
            label1.ForeColor = themeStatusText;
            statRuntime.Value = "00:00";
            statDeleted.Value = "0";
            statSkipped.Value = "0";
            infoLine.Visible = false;
            progressBar1.Value = 0;
            SetOverallProgress(0d);
            phaseLabel.Text = string.Empty;
            directoryValueLabel.Text = string.Empty;

            UpdateActivityList();
            PopulateRunTables();
        }

        // Von Hand gestartet wartet die Anwendung hier auf den Start. Noch ist nichts passiert,
        // die Einstellungen lassen sich also anpassen, ohne dass etwas gelöscht wird.
        private void ShowReadyState() {
            dashboardTimer.Stop();
            ResetRunState();
            LoadOptions();
            LoadConfiguration();
            ResetDashboard();
            waitingForStart = true;

            int count = cleaningDirectories.Count;
            headerTitleLabel.Text = dryRun ? "Vorschau bereit" : "Bereit";
            heroLabel.Text = count == 1 ? "1 Verzeichnis ausgewählt" : $"{count} Verzeichnisse ausgewählt";
            label1.Text = dryRun
                ? "Vorschau-Modus – es wird nichts gelöscht."
                : "Noch wurde nichts gelöscht. Passen Sie bei Bedarf die Einstellungen an und starten Sie dann die Bereinigung.";
            label2.Text = "Geplante Schritte";
            directoryValueLabel.Text = string.Join(", ", GetEnabledPhaseNames());

            pauseButton.Visible = false;
            cancelButton.Visible = false;
            retryButton.Visible = false;
            startButton.Text = dryRun ? "Vorschau starten" : "Bereinigung starten";
            startButton.Kind = UI.ButtonKind.Primary;
            startButton.Visible = true;
            finishButton.Kind = UI.ButtonKind.Secondary;
            finishButton.Visible = true;
            // Die Leiste füllt sich von rechts: Ganz rechts steht die Hauptaktion.
            actionsPanel.Controls.SetChildIndex(startButton, 0);
        }

        // Setzt die Anzeige zu Beginn eines Laufs zurück und startet die periodische Aktualisierung.
        private void PrepareDashboardForRun() {
            ResetDashboard();

            headerTitleLabel.Text = RunningTitle;
            label1.Text = defaultStatusText;

            pauseButton.Text = "Pausieren";
            pauseButton.Icon = UI.IconKind.Pause;
            pauseButton.Enabled = true;
            pauseButton.Visible = true;
            cancelButton.Enabled = true;
            cancelButton.Visible = true;
            startButton.Visible = false;
            finishButton.Visible = false;
            retryButton.Visible = false;

            dashboardTimer.Start();
        }

        // Schaltet die Aktionsleiste auf "Lauf beendet" und aktualisiert die Kennzahlen ein letztes Mal.
        private void ShowFinishedState(string title) {
            dashboardTimer.Stop();
            isPaused = false;
            ClearBusyStep();
            headerTitleLabel.Text = title;
            bool problem = System.Threading.Volatile.Read(ref failedItemCount) > 0 || title == "Bereinigung fehlgeschlagen" || title == "Vorgang abgebrochen";
            headerTitleLabel.ForeColor = problem ? themeError : themeSuccess;
            pauseButton.Visible = false;
            cancelButton.Visible = false;
            retryButton.Visible = !IsAdministrator();
            // Von Hand gestartet bleibt das Fenster offen, geänderte Einstellungen lassen sich direkt nutzen.
            startButton.Text = "Erneut starten";
            startButton.Kind = UI.ButtonKind.Secondary;
            startButton.Visible = !unattended;
            finishButton.Kind = UI.ButtonKind.Primary;
            finishButton.Visible = true;
            actionsPanel.Controls.SetChildIndex(finishButton, 0);
            RefreshDashboard();
        }

        private void RunOnUiThread(Action action) {
            if (InvokeRequired) {
                Invoke(action);
            } else {
                action();
            }
        }

        // Zeigt für lange Schritte, dass die Anwendung arbeitet: Laufbalken, Schrittname und Laufzeit.
        private void BeginBusyStep(string status, string target) {
            RunOnUiThread(() => {
                busyStatus = status;
                busyStartedAt = DateTime.Now;
                busyPercent = -1d;
                label2.Text = "Aktueller Schritt";
                directoryValueLabel.Text = target;
                label1.ForeColor = themeStatusText;
                label1.Text = BuildBusyText();
                progressBar1.Value = 0;
                progressBar1.Marquee = true;
                // Ein laufendes externes Programm lässt sich nicht anhalten; der Button würde nichts bewirken.
                pauseButton.Enabled = false;
            });
        }

        // Liefert ein Programm Prozentwerte, wird aus dem Laufbalken eine echte Fortschrittsanzeige.
        private void ReportBusyProgress(double percent) {
            BeginInvoke(new Action(() => {
                if (busyStatus == null) {
                    return;
                }
                busyPercent = percent;
                progressBar1.Marquee = false;
                progressBar1.Maximum = 1000;
                progressBar1.Value = (int)Math.Round(percent * 10d);
                label1.Text = BuildBusyText();
            }));
        }

        private void EndBusyStep() {
            RunOnUiThread(() => {
                if (busyStatus == null) {
                    return;
                }
                ClearBusyStep();
                progressBar1.Value = 0;
                label1.Text = dryRun ? "Vorschau-Modus – es wird nichts gelöscht." : defaultStatusText;
                pauseButton.Enabled = cleanupCts != null && !cleanupCts.IsCancellationRequested;
            });
        }

        private void ClearBusyStep() {
            busyStatus = null;
            busyPercent = -1d;
            progressBar1.Marquee = false;
            label2.Text = "Aktuelles Verzeichnis";
        }

        private string BuildBusyText() {
            string text = cleanupCts != null && cleanupCts.IsCancellationRequested
                ? "Abbruch angefordert – laufender Schritt wird noch abgeschlossen"
                : busyStatus;
            if (busyPercent >= 0d) {
                text += $" · {busyPercent:0} %";
            }
            TimeSpan elapsed = DateTime.Now - busyStartedAt;
            if (elapsed.TotalSeconds >= BusyElapsedAfterSeconds) {
                text += $" · {(int)elapsed.TotalMinutes:00}:{elapsed.Seconds:00}";
            }
            return text;
        }

        private void RefreshDashboard() {
            FlushDetailLog();
            DrainActivity();

            if (busyStatus != null) {
                label1.Text = BuildBusyText();
            } else if (currentPhase == 1 && directoriesTotal > 0 && !dryRun && dashboardTimer.Enabled
                       && (cleanupCts == null || !cleanupCts.IsCancellationRequested)) {
                label1.Text = $"Verzeichnis {Math.Min(directoriesDone + 1, directoriesTotal)} von {directoriesTotal}";
            }

            heroLabel.Text = HeroText;
            TimeSpan runtime = DateTime.Now - runStartTime;
            statRuntime.Value = $"{(int)runtime.TotalMinutes:00}:{runtime.Seconds:00}";
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

            infoLine.Tone = failed > 0 ? UI.Tone.Error : UI.Tone.Muted;
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
            } else if (busyPercent >= 0d) {
                phaseFraction = busyPercent / 100d;
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
