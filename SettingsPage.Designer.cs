namespace DataClean {
    partial class SettingsPage {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        private void InitializeComponent() {
            this.lblHeading = new System.Windows.Forms.Label();
            this.lblSubheading = new System.Windows.Forms.Label();
            this.settingsScroll = new DataClean.UI.ScrollHostPanel();
            this.settingsLayout = new System.Windows.Forms.TableLayoutPanel();
            this.grpProfile = new DataClean.UI.CardPanel();
            this.cmbProfile = new DataClean.UI.ThemedComboBox();
            this.btnLoadProfile = new DataClean.UI.ThemedButton();
            this.btnSaveProfile = new DataClean.UI.ThemedButton();
            this.btnDeleteProfile = new DataClean.UI.ThemedButton();
            this.grpCleanup = new DataClean.UI.CardPanel();
            this.chkDryRun = new DataClean.UI.ThemedCheckBox();
            this.chkCreateRestorePoint = new DataClean.UI.ThemedCheckBox();
            this.chkDeleteToRecycleBin = new DataClean.UI.ThemedCheckBox();
            this.chkConfirmBeforeDelete = new DataClean.UI.ThemedCheckBox();
            this.chkCleanBrowserCaches = new DataClean.UI.ThemedCheckBox();
            this.chkDismComponentCleanup = new DataClean.UI.ThemedCheckBox();
            this.chkCleanMinidumps = new DataClean.UI.ThemedCheckBox();
            this.chkCleanPrefetch = new DataClean.UI.ThemedCheckBox();
            this.chkCleanDeliveryOptimization = new DataClean.UI.ThemedCheckBox();
            this.chkEmptyAllRecycleBins = new DataClean.UI.ThemedCheckBox();
            this.chkCleanOldLogs = new DataClean.UI.ThemedCheckBox();
            this.numOldLogAgeDays = new DataClean.UI.ThemedNumericUpDown();
            this.lblOldLogAgeDays = new System.Windows.Forms.Label();
            this.grpReport = new DataClean.UI.CardPanel();
            this.chkGenerateReport = new DataClean.UI.ThemedCheckBox();
            this.lblReportPath = new System.Windows.Forms.Label();
            this.txtReportPath = new DataClean.UI.ThemedTextBox();
            this.btnBrowseReportPath = new DataClean.UI.ThemedButton();
            this.lblReportFormat = new System.Windows.Forms.Label();
            this.cmbReportFormat = new DataClean.UI.ThemedComboBox();
            this.chkEmailReport = new DataClean.UI.ThemedCheckBox();
            this.lblSmtpUser = new System.Windows.Forms.Label();
            this.txtSmtpUser = new DataClean.UI.ThemedTextBox();
            this.lblSmtpPassword = new System.Windows.Forms.Label();
            this.txtSmtpPassword = new DataClean.UI.ThemedTextBox();
            this.grpExclude = new DataClean.UI.CardPanel();
            this.txtExcludePatterns = new DataClean.UI.ThemedTextBox();
            this.lblExcludeHint = new System.Windows.Forms.Label();
            this.grpSchedule = new DataClean.UI.CardPanel();
            this.chkScheduleEnabled = new DataClean.UI.ThemedCheckBox();
            this.lblScheduleTime = new System.Windows.Forms.Label();
            this.dtpScheduleTime = new DataClean.UI.ThemedTimePicker();
            this.chkSkipIfEnoughFreeSpace = new DataClean.UI.ThemedCheckBox();
            this.grpGeneral = new DataClean.UI.CardPanel();
            this.chkVerboseLogging = new DataClean.UI.ThemedCheckBox();
            this.chkSchedulePendingDeletes = new DataClean.UI.ThemedCheckBox();
            this.lblTheme = new System.Windows.Forms.Label();
            this.cmbTheme = new DataClean.UI.ThemedComboBox();
            this.saveStatus = new DataClean.UI.IconLabel();
            this.btnSave = new DataClean.UI.ThemedButton();
            this.btnRevert = new DataClean.UI.ThemedButton();
            ((System.ComponentModel.ISupportInitialize)(this.numOldLogAgeDays)).BeginInit();
            this.settingsScroll.SuspendLayout();
            this.settingsLayout.SuspendLayout();
            this.grpProfile.SuspendLayout();
            this.grpCleanup.SuspendLayout();
            this.grpReport.SuspendLayout();
            this.grpExclude.SuspendLayout();
            this.grpSchedule.SuspendLayout();
            this.grpGeneral.SuspendLayout();
            this.SuspendLayout();
            //
            // lblHeading
            //
            this.lblHeading.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblHeading.Font = new System.Drawing.Font("Segoe UI", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHeading.Location = new System.Drawing.Point(20, 11);
            this.lblHeading.Name = "lblHeading";
            this.lblHeading.Size = new System.Drawing.Size(767, 44);
            this.lblHeading.TabIndex = 0;
            this.lblHeading.Text = "Einstellungen";
            //
            // lblSubheading
            //
            this.lblSubheading.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblSubheading.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSubheading.Location = new System.Drawing.Point(22, 54);
            this.lblSubheading.Name = "lblSubheading";
            this.lblSubheading.Size = new System.Drawing.Size(765, 26);
            this.lblSubheading.TabIndex = 1;
            this.lblSubheading.Tag = "subtle";
            this.lblSubheading.Text = "Änderungen gelten ab dem nächsten Lauf.";
            //
            // settingsScroll
            //
            this.settingsScroll.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.settingsScroll.AutoScroll = true;
            this.settingsScroll.Controls.Add(this.settingsLayout);
            this.settingsScroll.Location = new System.Drawing.Point(0, 93);
            this.settingsScroll.Name = "settingsScroll";
            this.settingsScroll.Padding = new System.Windows.Forms.Padding(24, 0, 7, 0);
            this.settingsScroll.Size = new System.Drawing.Size(794, 476);
            this.settingsScroll.TabIndex = 2;
            //
            // settingsLayout
            //
            this.settingsLayout.ColumnCount = 2;
            this.settingsLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.settingsLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.settingsLayout.Controls.Add(this.grpProfile, 0, 0);
            this.settingsLayout.Controls.Add(this.grpCleanup, 0, 1);
            this.settingsLayout.Controls.Add(this.grpReport, 0, 2);
            this.settingsLayout.Controls.Add(this.grpExclude, 1, 2);
            this.settingsLayout.Controls.Add(this.grpSchedule, 0, 3);
            this.settingsLayout.Controls.Add(this.grpGeneral, 1, 3);
            this.settingsLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.settingsLayout.Location = new System.Drawing.Point(24, 0);
            this.settingsLayout.Margin = new System.Windows.Forms.Padding(0);
            this.settingsLayout.Name = "settingsLayout";
            this.settingsLayout.RowCount = 4;
            this.settingsLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 64F));
            this.settingsLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 178F));
            this.settingsLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 222F));
            this.settingsLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 140F));
            this.settingsLayout.SetColumnSpan(this.grpProfile, 2);
            this.settingsLayout.SetColumnSpan(this.grpCleanup, 2);
            this.settingsLayout.Size = new System.Drawing.Size(763, 604);
            this.settingsLayout.TabIndex = 0;
            //
            // grpProfile
            //
            this.grpProfile.Controls.Add(this.cmbProfile);
            this.grpProfile.Controls.Add(this.btnLoadProfile);
            this.grpProfile.Controls.Add(this.btnSaveProfile);
            this.grpProfile.Controls.Add(this.btnDeleteProfile);
            this.grpProfile.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grpProfile.Location = new System.Drawing.Point(0, 0);
            this.grpProfile.Margin = new System.Windows.Forms.Padding(0, 0, 0, 12);
            this.grpProfile.Name = "grpProfile";
            this.grpProfile.Size = new System.Drawing.Size(763, 52);
            this.grpProfile.TabIndex = 0;
            this.grpProfile.Title = "Profil";
            //
            // cmbProfile
            //
            this.cmbProfile.FormattingEnabled = true;
            this.cmbProfile.Location = new System.Drawing.Point(82, 13);
            this.cmbProfile.Name = "cmbProfile";
            this.cmbProfile.Size = new System.Drawing.Size(270, 26);
            this.cmbProfile.TabIndex = 0;
            //
            // btnLoadProfile
            //
            this.btnLoadProfile.Location = new System.Drawing.Point(364, 12);
            this.btnLoadProfile.Name = "btnLoadProfile";
            this.btnLoadProfile.Size = new System.Drawing.Size(96, 28);
            this.btnLoadProfile.TabIndex = 1;
            this.btnLoadProfile.Text = "Laden";
            this.btnLoadProfile.Click += new System.EventHandler(this.btnLoadProfile_Click);
            //
            // btnSaveProfile
            //
            this.btnSaveProfile.Location = new System.Drawing.Point(468, 12);
            this.btnSaveProfile.Name = "btnSaveProfile";
            this.btnSaveProfile.Size = new System.Drawing.Size(96, 28);
            this.btnSaveProfile.TabIndex = 2;
            this.btnSaveProfile.Text = "Speichern";
            this.btnSaveProfile.Click += new System.EventHandler(this.btnSaveProfile_Click);
            //
            // btnDeleteProfile
            //
            this.btnDeleteProfile.Kind = DataClean.UI.ButtonKind.Danger;
            this.btnDeleteProfile.Location = new System.Drawing.Point(572, 12);
            this.btnDeleteProfile.Name = "btnDeleteProfile";
            this.btnDeleteProfile.Size = new System.Drawing.Size(96, 28);
            this.btnDeleteProfile.TabIndex = 3;
            this.btnDeleteProfile.Text = "Löschen";
            this.btnDeleteProfile.Click += new System.EventHandler(this.btnDeleteProfile_Click);
            //
            // grpCleanup
            //
            this.grpCleanup.Controls.Add(this.chkDryRun);
            this.grpCleanup.Controls.Add(this.chkCreateRestorePoint);
            this.grpCleanup.Controls.Add(this.chkDeleteToRecycleBin);
            this.grpCleanup.Controls.Add(this.chkConfirmBeforeDelete);
            this.grpCleanup.Controls.Add(this.chkCleanBrowserCaches);
            this.grpCleanup.Controls.Add(this.chkDismComponentCleanup);
            this.grpCleanup.Controls.Add(this.chkCleanMinidumps);
            this.grpCleanup.Controls.Add(this.chkCleanPrefetch);
            this.grpCleanup.Controls.Add(this.chkCleanDeliveryOptimization);
            this.grpCleanup.Controls.Add(this.chkEmptyAllRecycleBins);
            this.grpCleanup.Controls.Add(this.chkCleanOldLogs);
            this.grpCleanup.Controls.Add(this.numOldLogAgeDays);
            this.grpCleanup.Controls.Add(this.lblOldLogAgeDays);
            this.grpCleanup.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grpCleanup.Location = new System.Drawing.Point(0, 64);
            this.grpCleanup.Margin = new System.Windows.Forms.Padding(0, 0, 0, 12);
            this.grpCleanup.Name = "grpCleanup";
            this.grpCleanup.Size = new System.Drawing.Size(763, 166);
            this.grpCleanup.TabIndex = 1;
            this.grpCleanup.Title = "Bereinigungsoptionen";
            this.grpCleanup.TitleIcon = DataClean.UI.IconKind.Broom;
            //
            // chkDryRun
            //
            this.chkDryRun.AutoSize = true;
            this.chkDryRun.Location = new System.Drawing.Point(20, 46);
            this.chkDryRun.Name = "chkDryRun";
            this.chkDryRun.Size = new System.Drawing.Size(240, 22);
            this.chkDryRun.TabIndex = 0;
            this.chkDryRun.Text = "Vorschau-Modus (nichts wird gelöscht)";
            //
            // chkCreateRestorePoint
            //
            this.chkCreateRestorePoint.AutoSize = true;
            this.chkCreateRestorePoint.Location = new System.Drawing.Point(20, 74);
            this.chkCreateRestorePoint.Name = "chkCreateRestorePoint";
            this.chkCreateRestorePoint.Size = new System.Drawing.Size(240, 22);
            this.chkCreateRestorePoint.TabIndex = 1;
            this.chkCreateRestorePoint.Text = "Systemwiederherstellungspunkt anlegen";
            //
            // chkDeleteToRecycleBin
            //
            this.chkDeleteToRecycleBin.AutoSize = true;
            this.chkDeleteToRecycleBin.Location = new System.Drawing.Point(20, 102);
            this.chkDeleteToRecycleBin.Name = "chkDeleteToRecycleBin";
            this.chkDeleteToRecycleBin.Size = new System.Drawing.Size(190, 22);
            this.chkDeleteToRecycleBin.TabIndex = 2;
            this.chkDeleteToRecycleBin.Text = "In den Papierkorb verschieben";
            //
            // chkConfirmBeforeDelete
            //
            this.chkConfirmBeforeDelete.AutoSize = true;
            this.chkConfirmBeforeDelete.Location = new System.Drawing.Point(20, 130);
            this.chkConfirmBeforeDelete.Name = "chkConfirmBeforeDelete";
            this.chkConfirmBeforeDelete.Size = new System.Drawing.Size(240, 22);
            this.chkConfirmBeforeDelete.TabIndex = 3;
            this.chkConfirmBeforeDelete.Text = "Vor dem Löschen nachfragen (Vorschau)";
            //
            // chkCleanBrowserCaches
            //
            this.chkCleanBrowserCaches.AutoSize = true;
            this.chkCleanBrowserCaches.Location = new System.Drawing.Point(272, 46);
            this.chkCleanBrowserCaches.Name = "chkCleanBrowserCaches";
            this.chkCleanBrowserCaches.Size = new System.Drawing.Size(150, 22);
            this.chkCleanBrowserCaches.TabIndex = 4;
            this.chkCleanBrowserCaches.Text = "Browser-Caches leeren";
            //
            // chkDismComponentCleanup
            //
            this.chkDismComponentCleanup.AutoSize = true;
            this.chkDismComponentCleanup.Location = new System.Drawing.Point(272, 74);
            this.chkDismComponentCleanup.Name = "chkDismComponentCleanup";
            this.chkDismComponentCleanup.Size = new System.Drawing.Size(200, 22);
            this.chkDismComponentCleanup.TabIndex = 5;
            this.chkDismComponentCleanup.Text = "DISM Komponentenbereinigung";
            //
            // chkCleanMinidumps
            //
            this.chkCleanMinidumps.AutoSize = true;
            this.chkCleanMinidumps.Location = new System.Drawing.Point(272, 102);
            this.chkCleanMinidumps.Name = "chkCleanMinidumps";
            this.chkCleanMinidumps.Size = new System.Drawing.Size(165, 22);
            this.chkCleanMinidumps.TabIndex = 6;
            this.chkCleanMinidumps.Text = "Absturz-Dumps entfernen";
            //
            // chkCleanPrefetch
            //
            this.chkCleanPrefetch.AutoSize = true;
            this.chkCleanPrefetch.Location = new System.Drawing.Point(272, 130);
            this.chkCleanPrefetch.Name = "chkCleanPrefetch";
            this.chkCleanPrefetch.Size = new System.Drawing.Size(130, 22);
            this.chkCleanPrefetch.TabIndex = 7;
            this.chkCleanPrefetch.Text = "Prefetch bereinigen";
            //
            // chkCleanDeliveryOptimization
            //
            this.chkCleanDeliveryOptimization.AutoSize = true;
            this.chkCleanDeliveryOptimization.Location = new System.Drawing.Point(500, 46);
            this.chkCleanDeliveryOptimization.Name = "chkCleanDeliveryOptimization";
            this.chkCleanDeliveryOptimization.Size = new System.Drawing.Size(180, 22);
            this.chkCleanDeliveryOptimization.TabIndex = 8;
            this.chkCleanDeliveryOptimization.Text = "Delivery Optimization leeren";
            //
            // chkEmptyAllRecycleBins
            //
            this.chkEmptyAllRecycleBins.AutoSize = true;
            this.chkEmptyAllRecycleBins.Location = new System.Drawing.Point(500, 74);
            this.chkEmptyAllRecycleBins.Name = "chkEmptyAllRecycleBins";
            this.chkEmptyAllRecycleBins.Size = new System.Drawing.Size(205, 22);
            this.chkEmptyAllRecycleBins.TabIndex = 9;
            this.chkEmptyAllRecycleBins.Text = "Papierkorb aller Laufwerke leeren";
            //
            // chkCleanOldLogs
            //
            this.chkCleanOldLogs.AutoSize = true;
            this.chkCleanOldLogs.Location = new System.Drawing.Point(500, 102);
            this.chkCleanOldLogs.Name = "chkCleanOldLogs";
            this.chkCleanOldLogs.Size = new System.Drawing.Size(180, 22);
            this.chkCleanOldLogs.TabIndex = 10;
            this.chkCleanOldLogs.Text = "Alte Protokolle entfernen ab";
            //
            // numOldLogAgeDays
            //
            this.numOldLogAgeDays.Location = new System.Drawing.Point(524, 129);
            this.numOldLogAgeDays.Maximum = new decimal(new int[] {
            3650,
            0,
            0,
            0});
            this.numOldLogAgeDays.Name = "numOldLogAgeDays";
            this.numOldLogAgeDays.Size = new System.Drawing.Size(80, 23);
            this.numOldLogAgeDays.TabIndex = 11;
            this.numOldLogAgeDays.Value = new decimal(new int[] {
            30,
            0,
            0,
            0});
            //
            // lblOldLogAgeDays
            //
            this.lblOldLogAgeDays.AutoSize = true;
            this.lblOldLogAgeDays.Location = new System.Drawing.Point(610, 133);
            this.lblOldLogAgeDays.Name = "lblOldLogAgeDays";
            this.lblOldLogAgeDays.Size = new System.Drawing.Size(32, 15);
            this.lblOldLogAgeDays.TabIndex = 12;
            this.lblOldLogAgeDays.Text = "Tage";
            //
            // grpReport
            //
            this.grpReport.Controls.Add(this.chkGenerateReport);
            this.grpReport.Controls.Add(this.lblReportPath);
            this.grpReport.Controls.Add(this.txtReportPath);
            this.grpReport.Controls.Add(this.btnBrowseReportPath);
            this.grpReport.Controls.Add(this.lblReportFormat);
            this.grpReport.Controls.Add(this.cmbReportFormat);
            this.grpReport.Controls.Add(this.chkEmailReport);
            this.grpReport.Controls.Add(this.lblSmtpUser);
            this.grpReport.Controls.Add(this.txtSmtpUser);
            this.grpReport.Controls.Add(this.lblSmtpPassword);
            this.grpReport.Controls.Add(this.txtSmtpPassword);
            this.grpReport.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grpReport.Location = new System.Drawing.Point(0, 242);
            this.grpReport.Margin = new System.Windows.Forms.Padding(0, 0, 6, 12);
            this.grpReport.Name = "grpReport";
            this.grpReport.Size = new System.Drawing.Size(375, 210);
            this.grpReport.TabIndex = 2;
            this.grpReport.Title = "Report";
            this.grpReport.TitleIcon = DataClean.UI.IconKind.Document;
            //
            // chkGenerateReport
            //
            this.chkGenerateReport.AutoSize = true;
            this.chkGenerateReport.Location = new System.Drawing.Point(20, 46);
            this.chkGenerateReport.Name = "chkGenerateReport";
            this.chkGenerateReport.Size = new System.Drawing.Size(130, 22);
            this.chkGenerateReport.TabIndex = 0;
            this.chkGenerateReport.Text = "Report generieren";
            //
            // lblReportPath
            //
            this.lblReportPath.AutoSize = true;
            this.lblReportPath.Location = new System.Drawing.Point(20, 79);
            this.lblReportPath.Name = "lblReportPath";
            this.lblReportPath.Size = new System.Drawing.Size(64, 15);
            this.lblReportPath.TabIndex = 1;
            this.lblReportPath.Text = "Zielordner:";
            //
            // txtReportPath
            //
            this.txtReportPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtReportPath.Location = new System.Drawing.Point(124, 75);
            this.txtReportPath.Name = "txtReportPath";
            this.txtReportPath.Size = new System.Drawing.Size(189, 23);
            this.txtReportPath.TabIndex = 2;
            //
            // btnBrowseReportPath
            //
            this.btnBrowseReportPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowseReportPath.Icon = DataClean.UI.IconKind.Folder;
            this.btnBrowseReportPath.Location = new System.Drawing.Point(319, 74);
            this.btnBrowseReportPath.Name = "btnBrowseReportPath";
            this.btnBrowseReportPath.Size = new System.Drawing.Size(36, 25);
            this.btnBrowseReportPath.TabIndex = 3;
            this.btnBrowseReportPath.Click += new System.EventHandler(this.btnBrowseReportPath_Click);
            //
            // lblReportFormat
            //
            this.lblReportFormat.AutoSize = true;
            this.lblReportFormat.Location = new System.Drawing.Point(20, 112);
            this.lblReportFormat.Name = "lblReportFormat";
            this.lblReportFormat.Size = new System.Drawing.Size(48, 15);
            this.lblReportFormat.TabIndex = 4;
            this.lblReportFormat.Text = "Format:";
            //
            // cmbReportFormat
            //
            this.cmbReportFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbReportFormat.FormattingEnabled = true;
            this.cmbReportFormat.Items.AddRange(new object[] {
            "HTML",
            "CSV"});
            this.cmbReportFormat.Location = new System.Drawing.Point(124, 107);
            this.cmbReportFormat.Name = "cmbReportFormat";
            this.cmbReportFormat.Size = new System.Drawing.Size(74, 26);
            this.cmbReportFormat.TabIndex = 5;
            //
            // chkEmailReport
            //
            this.chkEmailReport.AutoSize = true;
            this.chkEmailReport.Location = new System.Drawing.Point(210, 109);
            this.chkEmailReport.Name = "chkEmailReport";
            this.chkEmailReport.Size = new System.Drawing.Size(125, 22);
            this.chkEmailReport.TabIndex = 6;
            this.chkEmailReport.Text = "Per E-Mail senden";
            //
            // lblSmtpUser
            //
            this.lblSmtpUser.AutoSize = true;
            this.lblSmtpUser.Location = new System.Drawing.Point(20, 145);
            this.lblSmtpUser.Name = "lblSmtpUser";
            this.lblSmtpUser.Size = new System.Drawing.Size(92, 15);
            this.lblSmtpUser.TabIndex = 7;
            this.lblSmtpUser.Text = "SMTP-Benutzer:";
            //
            // txtSmtpUser
            //
            this.txtSmtpUser.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSmtpUser.Location = new System.Drawing.Point(124, 141);
            this.txtSmtpUser.Name = "txtSmtpUser";
            this.txtSmtpUser.Size = new System.Drawing.Size(231, 23);
            this.txtSmtpUser.TabIndex = 8;
            //
            // lblSmtpPassword
            //
            this.lblSmtpPassword.AutoSize = true;
            this.lblSmtpPassword.Location = new System.Drawing.Point(20, 178);
            this.lblSmtpPassword.Name = "lblSmtpPassword";
            this.lblSmtpPassword.Size = new System.Drawing.Size(93, 15);
            this.lblSmtpPassword.TabIndex = 9;
            this.lblSmtpPassword.Text = "SMTP-Passwort:";
            //
            // txtSmtpPassword
            //
            this.txtSmtpPassword.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSmtpPassword.Location = new System.Drawing.Point(124, 174);
            this.txtSmtpPassword.Name = "txtSmtpPassword";
            this.txtSmtpPassword.Size = new System.Drawing.Size(231, 23);
            this.txtSmtpPassword.TabIndex = 10;
            this.txtSmtpPassword.UseSystemPasswordChar = true;
            //
            // grpExclude
            //
            this.grpExclude.Controls.Add(this.txtExcludePatterns);
            this.grpExclude.Controls.Add(this.lblExcludeHint);
            this.grpExclude.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grpExclude.Location = new System.Drawing.Point(387, 242);
            this.grpExclude.Margin = new System.Windows.Forms.Padding(6, 0, 0, 12);
            this.grpExclude.Name = "grpExclude";
            this.grpExclude.Size = new System.Drawing.Size(376, 210);
            this.grpExclude.TabIndex = 3;
            this.grpExclude.Title = "Ausschlussmuster";
            this.grpExclude.TitleIcon = DataClean.UI.IconKind.Skip;
            //
            // txtExcludePatterns
            //
            this.txtExcludePatterns.AcceptsReturn = true;
            this.txtExcludePatterns.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtExcludePatterns.Location = new System.Drawing.Point(20, 44);
            this.txtExcludePatterns.Multiline = true;
            this.txtExcludePatterns.Name = "txtExcludePatterns";
            this.txtExcludePatterns.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtExcludePatterns.Size = new System.Drawing.Size(336, 128);
            this.txtExcludePatterns.TabIndex = 0;
            this.txtExcludePatterns.WordWrap = false;
            //
            // lblExcludeHint
            //
            this.lblExcludeHint.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblExcludeHint.AutoEllipsis = true;
            this.lblExcludeHint.Location = new System.Drawing.Point(20, 177);
            this.lblExcludeHint.Name = "lblExcludeHint";
            this.lblExcludeHint.Size = new System.Drawing.Size(336, 18);
            this.lblExcludeHint.TabIndex = 1;
            this.lblExcludeHint.Tag = "subtle";
            this.lblExcludeHint.Text = "Ein Muster pro Zeile, z. B. *.log oder **\\Cache\\**";
            this.lblExcludeHint.UseMnemonic = false;
            //
            // grpSchedule
            //
            this.grpSchedule.Controls.Add(this.chkScheduleEnabled);
            this.grpSchedule.Controls.Add(this.lblScheduleTime);
            this.grpSchedule.Controls.Add(this.dtpScheduleTime);
            this.grpSchedule.Controls.Add(this.chkSkipIfEnoughFreeSpace);
            this.grpSchedule.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grpSchedule.Location = new System.Drawing.Point(0, 464);
            this.grpSchedule.Margin = new System.Windows.Forms.Padding(0, 0, 6, 0);
            this.grpSchedule.Name = "grpSchedule";
            this.grpSchedule.Size = new System.Drawing.Size(375, 140);
            this.grpSchedule.TabIndex = 4;
            this.grpSchedule.Title = "Automatische Ausführung";
            this.grpSchedule.TitleIcon = DataClean.UI.IconKind.Clock;
            //
            // chkScheduleEnabled
            //
            this.chkScheduleEnabled.AutoSize = true;
            this.chkScheduleEnabled.Location = new System.Drawing.Point(20, 46);
            this.chkScheduleEnabled.Name = "chkScheduleEnabled";
            this.chkScheduleEnabled.Size = new System.Drawing.Size(300, 22);
            this.chkScheduleEnabled.TabIndex = 0;
            this.chkScheduleEnabled.Text = "Täglich automatisch ausführen (Aufgabenplanung)";
            this.chkScheduleEnabled.CheckedChanged += new System.EventHandler(this.chkScheduleEnabled_CheckedChanged);
            //
            // lblScheduleTime
            //
            this.lblScheduleTime.AutoSize = true;
            this.lblScheduleTime.Location = new System.Drawing.Point(44, 79);
            this.lblScheduleTime.Name = "lblScheduleTime";
            this.lblScheduleTime.Size = new System.Drawing.Size(47, 15);
            this.lblScheduleTime.TabIndex = 1;
            this.lblScheduleTime.Text = "Uhrzeit:";
            //
            // dtpScheduleTime
            //
            this.dtpScheduleTime.Location = new System.Drawing.Point(98, 75);
            this.dtpScheduleTime.Name = "dtpScheduleTime";
            this.dtpScheduleTime.Size = new System.Drawing.Size(120, 23);
            this.dtpScheduleTime.TabIndex = 2;
            //
            // chkSkipIfEnoughFreeSpace
            //
            this.chkSkipIfEnoughFreeSpace.AutoSize = true;
            this.chkSkipIfEnoughFreeSpace.Location = new System.Drawing.Point(20, 106);
            this.chkSkipIfEnoughFreeSpace.Name = "chkSkipIfEnoughFreeSpace";
            this.chkSkipIfEnoughFreeSpace.Size = new System.Drawing.Size(310, 22);
            this.chkSkipIfEnoughFreeSpace.TabIndex = 3;
            this.chkSkipIfEnoughFreeSpace.Text = "Nur bereinigen, wenn Speicher unter Schwellenwert";
            //
            // grpGeneral
            //
            this.grpGeneral.Controls.Add(this.chkVerboseLogging);
            this.grpGeneral.Controls.Add(this.chkSchedulePendingDeletes);
            this.grpGeneral.Controls.Add(this.lblTheme);
            this.grpGeneral.Controls.Add(this.cmbTheme);
            this.grpGeneral.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grpGeneral.Location = new System.Drawing.Point(387, 464);
            this.grpGeneral.Margin = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.grpGeneral.Name = "grpGeneral";
            this.grpGeneral.Size = new System.Drawing.Size(376, 140);
            this.grpGeneral.TabIndex = 5;
            this.grpGeneral.Title = "Allgemein";
            this.grpGeneral.TitleIcon = DataClean.UI.IconKind.Gear;
            //
            // chkVerboseLogging
            //
            this.chkVerboseLogging.AutoSize = true;
            this.chkVerboseLogging.Location = new System.Drawing.Point(20, 46);
            this.chkVerboseLogging.Name = "chkVerboseLogging";
            this.chkVerboseLogging.Size = new System.Drawing.Size(155, 22);
            this.chkVerboseLogging.TabIndex = 0;
            this.chkVerboseLogging.Text = "Ausführliches Protokoll";
            //
            // chkSchedulePendingDeletes
            //
            this.chkSchedulePendingDeletes.AutoSize = true;
            this.chkSchedulePendingDeletes.Location = new System.Drawing.Point(20, 74);
            this.chkSchedulePendingDeletes.Name = "chkSchedulePendingDeletes";
            this.chkSchedulePendingDeletes.Size = new System.Drawing.Size(320, 22);
            this.chkSchedulePendingDeletes.TabIndex = 1;
            this.chkSchedulePendingDeletes.Text = "Gesperrte Dateien beim nächsten Neustart löschen";
            //
            // lblTheme
            //
            this.lblTheme.AutoSize = true;
            this.lblTheme.Location = new System.Drawing.Point(20, 107);
            this.lblTheme.Name = "lblTheme";
            this.lblTheme.Size = new System.Drawing.Size(46, 15);
            this.lblTheme.TabIndex = 2;
            this.lblTheme.Text = "Design:";
            //
            // cmbTheme
            //
            this.cmbTheme.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTheme.FormattingEnabled = true;
            this.cmbTheme.Items.AddRange(new object[] {
            "System",
            "Hell",
            "Dunkel"});
            this.cmbTheme.Location = new System.Drawing.Point(74, 102);
            this.cmbTheme.Name = "cmbTheme";
            this.cmbTheme.Size = new System.Drawing.Size(146, 26);
            this.cmbTheme.TabIndex = 3;
            //
            // saveStatus
            //
            this.saveStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.saveStatus.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.saveStatus.Icon = DataClean.UI.IconKind.CheckCircle;
            this.saveStatus.Location = new System.Drawing.Point(24, 585);
            this.saveStatus.Name = "saveStatus";
            this.saveStatus.Size = new System.Drawing.Size(500, 24);
            this.saveStatus.TabIndex = 3;
            this.saveStatus.Tone = DataClean.UI.Tone.Success;
            this.saveStatus.Visible = false;
            //
            // btnSave
            //
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.Enabled = false;
            this.btnSave.Kind = DataClean.UI.ButtonKind.Primary;
            this.btnSave.Location = new System.Drawing.Point(533, 581);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(120, 32);
            this.btnSave.TabIndex = 4;
            this.btnSave.Text = "Speichern";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            //
            // btnRevert
            //
            this.btnRevert.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRevert.Enabled = false;
            this.btnRevert.Location = new System.Drawing.Point(667, 581);
            this.btnRevert.Name = "btnRevert";
            this.btnRevert.Size = new System.Drawing.Size(120, 32);
            this.btnRevert.TabIndex = 5;
            this.btnRevert.Text = "Verwerfen";
            this.btnRevert.Click += new System.EventHandler(this.btnRevert_Click);
            //
            // SettingsPage
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lblHeading);
            this.Controls.Add(this.lblSubheading);
            this.Controls.Add(this.settingsScroll);
            this.Controls.Add(this.saveStatus);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnRevert);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "SettingsPage";
            this.Size = new System.Drawing.Size(802, 633);
            ((System.ComponentModel.ISupportInitialize)(this.numOldLogAgeDays)).EndInit();
            this.settingsScroll.ResumeLayout(false);
            this.settingsLayout.ResumeLayout(false);
            this.grpProfile.ResumeLayout(false);
            this.grpCleanup.ResumeLayout(false);
            this.grpCleanup.PerformLayout();
            this.grpReport.ResumeLayout(false);
            this.grpReport.PerformLayout();
            this.grpExclude.ResumeLayout(false);
            this.grpExclude.PerformLayout();
            this.grpSchedule.ResumeLayout(false);
            this.grpSchedule.PerformLayout();
            this.grpGeneral.ResumeLayout(false);
            this.grpGeneral.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblHeading;
        private System.Windows.Forms.Label lblSubheading;
        private DataClean.UI.ScrollHostPanel settingsScroll;
        private System.Windows.Forms.TableLayoutPanel settingsLayout;
        private DataClean.UI.CardPanel grpProfile;
        private DataClean.UI.ThemedComboBox cmbProfile;
        private DataClean.UI.ThemedButton btnLoadProfile;
        private DataClean.UI.ThemedButton btnSaveProfile;
        private DataClean.UI.ThemedButton btnDeleteProfile;
        private DataClean.UI.CardPanel grpCleanup;
        private DataClean.UI.ThemedCheckBox chkDryRun;
        private DataClean.UI.ThemedCheckBox chkCreateRestorePoint;
        private DataClean.UI.ThemedCheckBox chkDeleteToRecycleBin;
        private DataClean.UI.ThemedCheckBox chkConfirmBeforeDelete;
        private DataClean.UI.ThemedCheckBox chkCleanBrowserCaches;
        private DataClean.UI.ThemedCheckBox chkDismComponentCleanup;
        private DataClean.UI.ThemedCheckBox chkCleanMinidumps;
        private DataClean.UI.ThemedCheckBox chkCleanPrefetch;
        private DataClean.UI.ThemedCheckBox chkCleanDeliveryOptimization;
        private DataClean.UI.ThemedCheckBox chkEmptyAllRecycleBins;
        private DataClean.UI.ThemedCheckBox chkCleanOldLogs;
        private DataClean.UI.ThemedNumericUpDown numOldLogAgeDays;
        private System.Windows.Forms.Label lblOldLogAgeDays;
        private DataClean.UI.CardPanel grpReport;
        private DataClean.UI.ThemedCheckBox chkGenerateReport;
        private System.Windows.Forms.Label lblReportPath;
        private DataClean.UI.ThemedTextBox txtReportPath;
        private DataClean.UI.ThemedButton btnBrowseReportPath;
        private System.Windows.Forms.Label lblReportFormat;
        private DataClean.UI.ThemedComboBox cmbReportFormat;
        private DataClean.UI.ThemedCheckBox chkEmailReport;
        private System.Windows.Forms.Label lblSmtpUser;
        private DataClean.UI.ThemedTextBox txtSmtpUser;
        private System.Windows.Forms.Label lblSmtpPassword;
        private DataClean.UI.ThemedTextBox txtSmtpPassword;
        private DataClean.UI.CardPanel grpExclude;
        private DataClean.UI.ThemedTextBox txtExcludePatterns;
        private System.Windows.Forms.Label lblExcludeHint;
        private DataClean.UI.CardPanel grpSchedule;
        private DataClean.UI.ThemedCheckBox chkScheduleEnabled;
        private System.Windows.Forms.Label lblScheduleTime;
        private DataClean.UI.ThemedTimePicker dtpScheduleTime;
        private DataClean.UI.ThemedCheckBox chkSkipIfEnoughFreeSpace;
        private DataClean.UI.CardPanel grpGeneral;
        private DataClean.UI.ThemedCheckBox chkVerboseLogging;
        private DataClean.UI.ThemedCheckBox chkSchedulePendingDeletes;
        private System.Windows.Forms.Label lblTheme;
        private DataClean.UI.ThemedComboBox cmbTheme;
        private DataClean.UI.IconLabel saveStatus;
        private DataClean.UI.ThemedButton btnSave;
        private DataClean.UI.ThemedButton btnRevert;
    }
}
