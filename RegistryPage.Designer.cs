namespace DataClean {
    partial class RegistryPage {
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
            this.progressBar = new DataClean.UI.NeonProgressBar();
            this.chkStartup = new DataClean.UI.ThemedCheckBox();
            this.chkApplicationPaths = new DataClean.UI.ThemedCheckBox();
            this.chkObsoleteSoftware = new DataClean.UI.ThemedCheckBox();
            this.chkApplications = new DataClean.UI.ThemedCheckBox();
            this.chkSharedDlls = new DataClean.UI.ThemedCheckBox();
            this.chkFonts = new DataClean.UI.ThemedCheckBox();
            this.chkHelpFiles = new DataClean.UI.ThemedCheckBox();
            this.chkSoundEvents = new DataClean.UI.ThemedCheckBox();
            this.chkMuiCache = new DataClean.UI.ThemedCheckBox();
            this.lblCaution = new System.Windows.Forms.Label();
            this.chkComClasses = new DataClean.UI.ThemedCheckBox();
            this.chkTypeLibraries = new DataClean.UI.ThemedCheckBox();
            this.lblCautionHint = new System.Windows.Forms.Label();
            this.columnDivider = new System.Windows.Forms.Panel();
            this.issuesList = new DataClean.UI.ThemedListView();
            this.colIssueName = new System.Windows.Forms.ColumnHeader();
            this.colIssueCategory = new System.Windows.Forms.ColumnHeader();
            this.colIssuePath = new System.Windows.Forms.ColumnHeader();
            this.selectAllCheck = new DataClean.UI.ThemedCheckBox();
            this.actionsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.removeButton = new DataClean.UI.ThemedButton();
            this.scanButton = new DataClean.UI.ThemedButton();
            this.restoreButton = new DataClean.UI.ThemedButton();
            this.actionsPanel.SuspendLayout();
            this.SuspendLayout();
            //
            // lblHeading
            //
            this.lblHeading.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblHeading.Font = new System.Drawing.Font("Segoe UI Semibold", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHeading.Location = new System.Drawing.Point(20, 11);
            this.lblHeading.Name = "lblHeading";
            this.lblHeading.Size = new System.Drawing.Size(767, 44);
            this.lblHeading.TabIndex = 0;
            this.lblHeading.Text = "Registry";
            //
            // lblSubheading
            //
            this.lblSubheading.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblSubheading.AutoEllipsis = true;
            this.lblSubheading.Font = new System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSubheading.Location = new System.Drawing.Point(22, 54);
            this.lblSubheading.Name = "lblSubheading";
            this.lblSubheading.Size = new System.Drawing.Size(765, 26);
            this.lblSubheading.TabIndex = 1;
            this.lblSubheading.Tag = "subtle";
            this.lblSubheading.Text = "Findet Einträge, deren Datei nachweislich fehlt. Vor dem Entfernen wird gesichert.";
            //
            // progressBar
            //
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(22, 82);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(765, 4);
            this.progressBar.TabIndex = 2;
            this.progressBar.Visible = false;
            //
            // chkStartup
            //
            this.chkStartup.Checked = true;
            this.chkStartup.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkStartup.Location = new System.Drawing.Point(22, 96);
            this.chkStartup.Name = "chkStartup";
            this.chkStartup.Size = new System.Drawing.Size(190, 22);
            this.chkStartup.TabIndex = 3;
            this.chkStartup.Text = "Autostart";
            //
            // chkApplicationPaths
            //
            this.chkApplicationPaths.Checked = true;
            this.chkApplicationPaths.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkApplicationPaths.Location = new System.Drawing.Point(22, 122);
            this.chkApplicationPaths.Name = "chkApplicationPaths";
            this.chkApplicationPaths.Size = new System.Drawing.Size(190, 22);
            this.chkApplicationPaths.TabIndex = 4;
            this.chkApplicationPaths.Text = "Anwendungspfade";
            //
            // chkObsoleteSoftware
            //
            this.chkObsoleteSoftware.Checked = true;
            this.chkObsoleteSoftware.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkObsoleteSoftware.Location = new System.Drawing.Point(22, 148);
            this.chkObsoleteSoftware.Name = "chkObsoleteSoftware";
            this.chkObsoleteSoftware.Size = new System.Drawing.Size(190, 22);
            this.chkObsoleteSoftware.TabIndex = 5;
            this.chkObsoleteSoftware.Text = "Veraltete Software";
            //
            // chkApplications
            //
            this.chkApplications.Checked = true;
            this.chkApplications.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkApplications.Location = new System.Drawing.Point(22, 174);
            this.chkApplications.Name = "chkApplications";
            this.chkApplications.Size = new System.Drawing.Size(190, 22);
            this.chkApplications.TabIndex = 6;
            this.chkApplications.Text = "Anwendungen";
            //
            // chkSharedDlls
            //
            this.chkSharedDlls.Checked = true;
            this.chkSharedDlls.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSharedDlls.Location = new System.Drawing.Point(22, 200);
            this.chkSharedDlls.Name = "chkSharedDlls";
            this.chkSharedDlls.Size = new System.Drawing.Size(190, 22);
            this.chkSharedDlls.TabIndex = 7;
            this.chkSharedDlls.Text = "Gemeinsame DLLs";
            //
            // chkFonts
            //
            this.chkFonts.Checked = true;
            this.chkFonts.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkFonts.Location = new System.Drawing.Point(22, 226);
            this.chkFonts.Name = "chkFonts";
            this.chkFonts.Size = new System.Drawing.Size(190, 22);
            this.chkFonts.TabIndex = 8;
            this.chkFonts.Text = "Schriftarten";
            //
            // chkHelpFiles
            //
            this.chkHelpFiles.Checked = true;
            this.chkHelpFiles.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkHelpFiles.Location = new System.Drawing.Point(22, 252);
            this.chkHelpFiles.Name = "chkHelpFiles";
            this.chkHelpFiles.Size = new System.Drawing.Size(190, 22);
            this.chkHelpFiles.TabIndex = 9;
            this.chkHelpFiles.Text = "Hilfedateien";
            //
            // chkSoundEvents
            //
            this.chkSoundEvents.Checked = true;
            this.chkSoundEvents.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkSoundEvents.Location = new System.Drawing.Point(22, 278);
            this.chkSoundEvents.Name = "chkSoundEvents";
            this.chkSoundEvents.Size = new System.Drawing.Size(190, 22);
            this.chkSoundEvents.TabIndex = 10;
            this.chkSoundEvents.Text = "Soundereignisse";
            //
            // chkMuiCache
            //
            this.chkMuiCache.Checked = true;
            this.chkMuiCache.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkMuiCache.Location = new System.Drawing.Point(22, 304);
            this.chkMuiCache.Name = "chkMuiCache";
            this.chkMuiCache.Size = new System.Drawing.Size(190, 22);
            this.chkMuiCache.TabIndex = 11;
            this.chkMuiCache.Text = "MUI-Cache";
            //
            // lblCaution
            //
            this.lblCaution.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCaution.Location = new System.Drawing.Point(20, 342);
            this.lblCaution.Name = "lblCaution";
            this.lblCaution.Size = new System.Drawing.Size(192, 20);
            this.lblCaution.TabIndex = 12;
            this.lblCaution.Text = "Mit Vorsicht";
            //
            // chkComClasses
            //
            this.chkComClasses.Location = new System.Drawing.Point(22, 366);
            this.chkComClasses.Name = "chkComClasses";
            this.chkComClasses.Size = new System.Drawing.Size(190, 22);
            this.chkComClasses.TabIndex = 13;
            this.chkComClasses.Text = "ActiveX und COM";
            //
            // chkTypeLibraries
            //
            this.chkTypeLibraries.Location = new System.Drawing.Point(22, 392);
            this.chkTypeLibraries.Name = "chkTypeLibraries";
            this.chkTypeLibraries.Size = new System.Drawing.Size(190, 22);
            this.chkTypeLibraries.TabIndex = 14;
            this.chkTypeLibraries.Text = "Typbibliotheken";
            //
            // lblCautionHint
            //
            this.lblCautionHint.Location = new System.Drawing.Point(20, 420);
            this.lblCautionHint.Name = "lblCautionHint";
            this.lblCautionHint.Size = new System.Drawing.Size(192, 52);
            this.lblCautionHint.TabIndex = 15;
            this.lblCautionHint.Tag = "subtle";
            this.lblCautionHint.Text = "Andere Programme können davon abhängen. Funde sind nicht vorausgewählt.";
            //
            // columnDivider
            //
            this.columnDivider.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)));
            this.columnDivider.Location = new System.Drawing.Point(224, 96);
            this.columnDivider.Name = "columnDivider";
            this.columnDivider.Size = new System.Drawing.Size(1, 472);
            this.columnDivider.TabIndex = 16;
            this.columnDivider.Tag = "separator";
            //
            // issuesList
            //
            this.issuesList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.issuesList.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.issuesList.CheckBoxes = true;
            this.issuesList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colIssueName,
            this.colIssueCategory,
            this.colIssuePath});
            this.issuesList.FullRowSelect = true;
            this.issuesList.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.issuesList.Location = new System.Drawing.Point(236, 92);
            this.issuesList.MultiSelect = false;
            this.issuesList.Name = "issuesList";
            this.issuesList.OwnerDraw = true;
            this.issuesList.PathColumn = 2;
            this.issuesList.ShowItemToolTips = true;
            this.issuesList.ShowSelection = false;
            this.issuesList.Size = new System.Drawing.Size(551, 476);
            this.issuesList.TabIndex = 17;
            this.issuesList.UseCompatibleStateImageBehavior = false;
            this.issuesList.View = System.Windows.Forms.View.Details;
            this.issuesList.ItemChecked += new System.Windows.Forms.ItemCheckedEventHandler(this.IssuesList_ItemChecked);
            //
            // colIssueName
            //
            this.colIssueName.Text = "Eintrag";
            this.colIssueName.Width = 195;
            //
            // colIssueCategory
            //
            this.colIssueCategory.Text = "Bereich";
            this.colIssueCategory.Width = 115;
            //
            // colIssuePath
            //
            this.colIssuePath.Text = "Fehlende Datei";
            this.colIssuePath.Width = 241;
            //
            // selectAllCheck
            //
            this.selectAllCheck.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.selectAllCheck.Location = new System.Drawing.Point(22, 586);
            this.selectAllCheck.Name = "selectAllCheck";
            this.selectAllCheck.Size = new System.Drawing.Size(190, 22);
            this.selectAllCheck.TabIndex = 18;
            this.selectAllCheck.Text = "Alle Funde auswählen";
            this.selectAllCheck.CheckedChanged += new System.EventHandler(this.SelectAllCheck_CheckedChanged);
            //
            // actionsPanel
            //
            this.actionsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.actionsPanel.Controls.Add(this.removeButton);
            this.actionsPanel.Controls.Add(this.scanButton);
            this.actionsPanel.Controls.Add(this.restoreButton);
            this.actionsPanel.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.actionsPanel.Location = new System.Drawing.Point(236, 581);
            this.actionsPanel.Margin = new System.Windows.Forms.Padding(0);
            this.actionsPanel.Name = "actionsPanel";
            this.actionsPanel.Size = new System.Drawing.Size(551, 33);
            this.actionsPanel.TabIndex = 19;
            this.actionsPanel.WrapContents = false;
            //
            // removeButton
            //
            this.removeButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.removeButton.Kind = DataClean.UI.ButtonKind.Danger;
            this.removeButton.Margin = new System.Windows.Forms.Padding(14, 0, 0, 0);
            this.removeButton.Name = "removeButton";
            this.removeButton.Size = new System.Drawing.Size(160, 32);
            this.removeButton.TabIndex = 2;
            this.removeButton.Text = "Auswahl entfernen";
            this.removeButton.UseVisualStyleBackColor = false;
            this.removeButton.Click += new System.EventHandler(this.RemoveButton_Click);
            //
            // scanButton
            //
            this.scanButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.scanButton.Kind = DataClean.UI.ButtonKind.Primary;
            this.scanButton.Margin = new System.Windows.Forms.Padding(14, 0, 0, 0);
            this.scanButton.Name = "scanButton";
            this.scanButton.Size = new System.Drawing.Size(110, 32);
            this.scanButton.TabIndex = 1;
            this.scanButton.Text = "Prüfen";
            this.scanButton.UseVisualStyleBackColor = false;
            this.scanButton.Click += new System.EventHandler(this.ScanButton_Click);
            //
            // restoreButton
            //
            this.restoreButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.restoreButton.Kind = DataClean.UI.ButtonKind.Secondary;
            this.restoreButton.Margin = new System.Windows.Forms.Padding(14, 0, 0, 0);
            this.restoreButton.Name = "restoreButton";
            this.restoreButton.Size = new System.Drawing.Size(150, 32);
            this.restoreButton.TabIndex = 0;
            this.restoreButton.Text = "Sicherung einspielen";
            this.restoreButton.UseVisualStyleBackColor = false;
            this.restoreButton.Click += new System.EventHandler(this.RestoreButton_Click);
            //
            // RegistryPage
            //
            this.Controls.Add(this.lblHeading);
            this.Controls.Add(this.lblSubheading);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.chkStartup);
            this.Controls.Add(this.chkApplicationPaths);
            this.Controls.Add(this.chkObsoleteSoftware);
            this.Controls.Add(this.chkApplications);
            this.Controls.Add(this.chkSharedDlls);
            this.Controls.Add(this.chkFonts);
            this.Controls.Add(this.chkHelpFiles);
            this.Controls.Add(this.chkSoundEvents);
            this.Controls.Add(this.chkMuiCache);
            this.Controls.Add(this.lblCaution);
            this.Controls.Add(this.chkComClasses);
            this.Controls.Add(this.chkTypeLibraries);
            this.Controls.Add(this.lblCautionHint);
            this.Controls.Add(this.columnDivider);
            this.Controls.Add(this.issuesList);
            this.Controls.Add(this.selectAllCheck);
            this.Controls.Add(this.actionsPanel);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "RegistryPage";
            this.Size = new System.Drawing.Size(802, 633);
            this.actionsPanel.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblHeading;
        private System.Windows.Forms.Label lblSubheading;
        private DataClean.UI.NeonProgressBar progressBar;
        private DataClean.UI.ThemedCheckBox chkStartup;
        private DataClean.UI.ThemedCheckBox chkApplicationPaths;
        private DataClean.UI.ThemedCheckBox chkObsoleteSoftware;
        private DataClean.UI.ThemedCheckBox chkApplications;
        private DataClean.UI.ThemedCheckBox chkSharedDlls;
        private DataClean.UI.ThemedCheckBox chkFonts;
        private DataClean.UI.ThemedCheckBox chkHelpFiles;
        private DataClean.UI.ThemedCheckBox chkSoundEvents;
        private DataClean.UI.ThemedCheckBox chkMuiCache;
        private System.Windows.Forms.Label lblCaution;
        private DataClean.UI.ThemedCheckBox chkComClasses;
        private DataClean.UI.ThemedCheckBox chkTypeLibraries;
        private System.Windows.Forms.Label lblCautionHint;
        private System.Windows.Forms.Panel columnDivider;
        private DataClean.UI.ThemedListView issuesList;
        private System.Windows.Forms.ColumnHeader colIssueName;
        private System.Windows.Forms.ColumnHeader colIssueCategory;
        private System.Windows.Forms.ColumnHeader colIssuePath;
        private DataClean.UI.ThemedCheckBox selectAllCheck;
        private System.Windows.Forms.FlowLayoutPanel actionsPanel;
        private DataClean.UI.ThemedButton removeButton;
        private DataClean.UI.ThemedButton scanButton;
        private DataClean.UI.ThemedButton restoreButton;
    }
}
