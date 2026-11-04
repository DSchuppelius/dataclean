using System;
using System.Windows.Forms;

namespace DataClean {
    /// <summary>
    /// Seite im Hauptfenster zum Bearbeiten der benutzerbezogenen Bereinigungs- und Report-Optionen.
    /// Anwendungsweite Einstellungen (z. B. geschützte Verzeichnisse) werden hier bewusst nicht
    /// angeboten, da sie zur Laufzeit schreibgeschützt sind.
    /// </summary>
    public partial class SettingsPage : UserControl, UI.IThemedControl {
        private bool loaded;
        private bool loading;
        private bool dirty;

        public SettingsPage() {
            InitializeComponent();
            ApplyTheme();
        }

        /// <summary>Wird nach erfolgreichem Speichern ausgelöst.</summary>
        public event EventHandler Saved;

        public void ApplyTheme() {
            BackColor = UI.AppTheme.Background;
        }

        protected override void WndProc(ref Message m) {
            if (UI.EdgeTransparentPanel.HandleEdgeHitTest(this, ref m)) {
                return;
            }
            base.WndProc(ref m);
        }

        /// <summary>
        /// Lädt die Werte beim ersten Öffnen. Erst dann, weil die Abfrage der geplanten Aufgabe
        /// einen externen Prozess startet und den Programmstart nicht verzögern soll.
        /// </summary>
        public void EnsureLoaded() {
            if (loaded) {
                return;
            }
            loaded = true;
            ReloadFromSettings();
            TrackChanges(this);
        }

        /// <summary>
        /// Fragt bei ungespeicherten Änderungen nach. Gibt false zurück, wenn die Seite
        /// geöffnet bleiben soll.
        /// </summary>
        public bool ConfirmLeave() {
            if (!dirty) {
                return true;
            }

            DialogResult answer = UI.AppDialog.Show(FindForm(),
                $"Die Änderungen wurden noch nicht gespeichert.{Environment.NewLine}{Environment.NewLine}Möchten Sie sie jetzt speichern?",
                "Einstellungen", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            if (answer == DialogResult.Yes) {
                return TrySave();
            }
            if (answer == DialogResult.No) {
                ReloadFromSettings();
                return true;
            }
            return false;
        }

        private void ReloadFromSettings() {
            loading = true;
            try {
                LoadFromOptions(CleanupOptions.FromSettings());
                LoadScheduleState();
                LoadProfiles();
            } finally {
                loading = false;
            }
            SetDirty(false);
        }

        // Meldet jede Änderung eines Eingabefelds; die Profilauswahl selbst ist keine Einstellung.
        private void TrackChanges(Control root) {
            foreach (Control control in root.Controls) {
                if (control == cmbProfile) {
                    continue;
                }

                if (control is CheckBox checkBox) {
                    checkBox.CheckedChanged += Input_Changed;
                } else if (control is NumericUpDown number) {
                    number.ValueChanged += Input_Changed;
                } else if (control is UI.ThemedTimePicker time) {
                    time.ValueChanged += Input_Changed;
                } else if (control is ComboBox combo) {
                    combo.SelectedIndexChanged += Input_Changed;
                } else if (control is TextBox text) {
                    text.TextChanged += Input_Changed;
                } else if (control.HasChildren) {
                    TrackChanges(control);
                }
            }
        }

        private void Input_Changed(object sender, EventArgs e) {
            if (!loading) {
                SetDirty(true);
            }
        }

        private void SetDirty(bool value) {
            dirty = value;
            btnSave.Enabled = value;
            btnRevert.Enabled = value;
            if (value) {
                saveStatus.Visible = false;
            }
        }

        private void LoadFromOptions(CleanupOptions o) {
            ApplyOptionsToUi(o);

            var s = Properties.Settings.Default;
            txtSmtpUser.Text = s.smtpUser ?? string.Empty;
            txtSmtpPassword.Text = Services.SecureSettings.Decrypt(s.smtpPasswordEncrypted);
            chkVerboseLogging.Checked = s.verboseLogging;

            switch ((s.themeOverride ?? "System").ToLowerInvariant()) {
                case "light": cmbTheme.SelectedItem = "Hell"; break;
                case "dark": cmbTheme.SelectedItem = "Dunkel"; break;
                default: cmbTheme.SelectedItem = "System"; break;
            }
        }

        private void SaveToOptions() {
            var o = CollectOptionsFromUi();

            o.Save();

            // SMTP-Zugangsdaten gesondert speichern – das Passwort per DPAPI verschlüsselt.
            var s = Properties.Settings.Default;
            s.smtpUser = txtSmtpUser.Text.Trim();
            s.smtpPasswordEncrypted = Services.SecureSettings.Encrypt(txtSmtpPassword.Text);
            s.verboseLogging = chkVerboseLogging.Checked;
            switch (cmbTheme.SelectedItem as string) {
                case "Hell": s.themeOverride = "Light"; break;
                case "Dunkel": s.themeOverride = "Dark"; break;
                default: s.themeOverride = "System"; break;
            }
            s.Save();

            Services.Logger.MinimumLevel = chkVerboseLogging.Checked
                ? Services.LogLevel.Info
                : Services.LogLevel.Warn;
        }

        /// <summary>
        /// Liest die aktuellen Werte der Oberfläche in ein CleanupOptions-Objekt.
        /// </summary>
        private CleanupOptions CollectOptionsFromUi() {
            var o = CleanupOptions.FromSettings();

            o.DryRun = chkDryRun.Checked;
            o.CreateRestorePoint = chkCreateRestorePoint.Checked;
            o.DeleteToRecycleBin = chkDeleteToRecycleBin.Checked;
            o.CleanBrowserCaches = chkCleanBrowserCaches.Checked;
            o.DismComponentCleanup = chkDismComponentCleanup.Checked;
            o.CleanMinidumps = chkCleanMinidumps.Checked;
            o.CleanPrefetch = chkCleanPrefetch.Checked;
            o.CleanDeliveryOptimization = chkCleanDeliveryOptimization.Checked;
            o.CleanOldLogs = chkCleanOldLogs.Checked;
            o.OldLogAgeDays = (int)numOldLogAgeDays.Value;
            o.EmptyAllRecycleBins = chkEmptyAllRecycleBins.Checked;
            o.ConfirmBeforeDelete = chkConfirmBeforeDelete.Checked;
            o.GenerateReport = chkGenerateReport.Checked;
            o.ReportPath = txtReportPath.Text.Trim();
            o.ReportFormat = (cmbReportFormat.SelectedItem as string) ?? "HTML";
            o.EmailReport = chkEmailReport.Checked;

            o.ExcludePatterns = txtExcludePatterns.Text.Trim();

            o.SkipIfEnoughFreeSpace = chkSkipIfEnoughFreeSpace.Checked;

            o.SchedulePendingDeletes = chkSchedulePendingDeletes.Checked;

            return o;
        }

        /// <summary>
        /// Überträgt die Werte eines CleanupOptions-Objekts in die Oberfläche
        /// (ohne SMTP-Felder, die separat verwaltet werden).
        /// </summary>
        private void ApplyOptionsToUi(CleanupOptions o) {
            chkDryRun.Checked = o.DryRun;
            chkCreateRestorePoint.Checked = o.CreateRestorePoint;
            chkDeleteToRecycleBin.Checked = o.DeleteToRecycleBin;
            chkCleanBrowserCaches.Checked = o.CleanBrowserCaches;
            chkDismComponentCleanup.Checked = o.DismComponentCleanup;
            chkCleanMinidumps.Checked = o.CleanMinidumps;
            chkCleanPrefetch.Checked = o.CleanPrefetch;
            chkCleanDeliveryOptimization.Checked = o.CleanDeliveryOptimization;
            chkCleanOldLogs.Checked = o.CleanOldLogs;
            numOldLogAgeDays.Value = Math.Max(numOldLogAgeDays.Minimum,
                Math.Min(numOldLogAgeDays.Maximum, o.OldLogAgeDays));
            chkEmptyAllRecycleBins.Checked = o.EmptyAllRecycleBins;
            chkConfirmBeforeDelete.Checked = o.ConfirmBeforeDelete;
            chkGenerateReport.Checked = o.GenerateReport;
            txtReportPath.Text = o.ReportPath ?? string.Empty;
            cmbReportFormat.SelectedItem = string.Equals(o.ReportFormat, "CSV", StringComparison.OrdinalIgnoreCase)
                ? "CSV" : "HTML";
            chkEmailReport.Checked = o.EmailReport;
            txtExcludePatterns.Text = o.ExcludePatterns ?? string.Empty;
            chkSkipIfEnoughFreeSpace.Checked = o.SkipIfEnoughFreeSpace;
            chkSchedulePendingDeletes.Checked = o.SchedulePendingDeletes;
        }

        private void LoadProfiles() {
            cmbProfile.Items.Clear();
            foreach (var name in Services.CleanupProfileStore.GetNames()) {
                cmbProfile.Items.Add(name);
            }
        }

        private void btnLoadProfile_Click(object sender, EventArgs e) {
            string name = cmbProfile.SelectedItem as string ?? cmbProfile.Text;
            var profile = Services.CleanupProfileStore.Get(name);
            if (profile == null) {
                UI.AppDialog.Show(FindForm(), "Bitte zuerst ein vorhandenes Profil auswählen.",
                    "Profil laden", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var o = CleanupOptions.FromSettings();
            profile.ApplyTo(o);
            ApplyOptionsToUi(o);
        }

        private void btnSaveProfile_Click(object sender, EventArgs e) {
            string name = (cmbProfile.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(name)) {
                UI.AppDialog.Show(FindForm(), "Bitte einen Profilnamen eingeben oder auswählen.",
                    "Profil speichern", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var profile = Services.CleanupProfile.FromOptions(name, CollectOptionsFromUi());
            if (Services.CleanupProfileStore.Save(profile)) {
                LoadProfiles();
                cmbProfile.SelectedItem = name;
            } else {
                UI.AppDialog.Show(FindForm(), "Das Profil konnte nicht gespeichert werden.",
                    "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDeleteProfile_Click(object sender, EventArgs e) {
            string name = cmbProfile.SelectedItem as string ?? cmbProfile.Text;
            if (string.IsNullOrWhiteSpace(name)) {
                return;
            }

            if (UI.AppDialog.Show(FindForm(), $"Profil \"{name}\" wirklich löschen?",
                "Profil löschen", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) {
                return;
            }

            if (Services.CleanupProfileStore.Delete(name)) {
                LoadProfiles();
                cmbProfile.Text = string.Empty;
            }
        }

        private void btnBrowseReportPath_Click(object sender, EventArgs e) {
            using (var dlg = new UI.FolderPickerDialog("Zielordner für Reports auswählen", txtReportPath.Text)) {
                if (dlg.ShowDialog(FindForm()) == DialogResult.OK) {
                    txtReportPath.Text = dlg.SelectedPath;
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e) {
            TrySave();
        }

        private bool TrySave() {
            try {
                SaveToOptions();
                ApplyScheduleState();
            } catch (Exception ex) {
                Services.Logger.Error(ex, "SettingsPage");
                UI.AppDialog.Show(FindForm(), $"Einstellungen konnten nicht gespeichert werden:{Environment.NewLine}{ex.Message}",
                    "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            SetDirty(false);
            saveStatus.Text = "Gespeichert. Änderungen gelten ab dem nächsten Lauf.";
            saveStatus.Visible = true;
            Saved?.Invoke(this, EventArgs.Empty);
            return true;
        }

        private void btnRevert_Click(object sender, EventArgs e) {
            ReloadFromSettings();
        }

        private void chkScheduleEnabled_CheckedChanged(object sender, EventArgs e) {
            dtpScheduleTime.Enabled = chkScheduleEnabled.Checked;
            lblScheduleTime.Enabled = chkScheduleEnabled.Checked;
        }

        private void LoadScheduleState() {
            bool exists = Services.SchedulerService.TaskExists();
            chkScheduleEnabled.Checked = exists;
            dtpScheduleTime.Enabled = exists;
            lblScheduleTime.Enabled = exists;

            // Standard-Uhrzeit (z. B. 03:00 Uhr), falls noch keine Aufgabe existiert.
            dtpScheduleTime.Value = DateTime.Today.AddHours(3);
        }

        private void ApplyScheduleState() {
            bool exists = Services.SchedulerService.TaskExists();
            string message;

            if (chkScheduleEnabled.Checked) {
                // Aufgabe (neu) anlegen bzw. mit aktueller Uhrzeit aktualisieren.
                var time = dtpScheduleTime.Value.TimeOfDay;
                if (!Services.SchedulerService.CreateDailyTask(time, out message)) {
                    throw new InvalidOperationException(
                        "Die geplante Aufgabe konnte nicht angelegt werden. Möglicherweise sind Administratorrechte erforderlich." +
                        (string.IsNullOrWhiteSpace(message) ? string.Empty : Environment.NewLine + message));
                }
            } else if (exists) {
                // Vorhandene Aufgabe entfernen.
                if (!Services.SchedulerService.DeleteTask(out message)) {
                    throw new InvalidOperationException(
                        "Die geplante Aufgabe konnte nicht entfernt werden. Möglicherweise sind Administratorrechte erforderlich." +
                        (string.IsNullOrWhiteSpace(message) ? string.Empty : Environment.NewLine + message));
                }
            }
        }
    }
}
