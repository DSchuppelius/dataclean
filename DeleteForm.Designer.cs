namespace DataClean {
    partial class DeleteForm {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent() {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DeleteForm));
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.overviewPage = new DataClean.UI.PagePanel();
            this.headerTitleLabel = new System.Windows.Forms.Label();
            this.heroLabel = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.statsLayout = new System.Windows.Forms.TableLayoutPanel();
            this.statRuntime = new DataClean.UI.StatCard();
            this.statDeleted = new DataClean.UI.StatCard();
            this.statSkipped = new DataClean.UI.StatCard();
            this.progressCard = new DataClean.UI.CardPanel();
            this.phaseLabel = new System.Windows.Forms.Label();
            this.percentLabel = new System.Windows.Forms.Label();
            this.progressBar2 = new DataClean.UI.NeonProgressBar();
            this.label2 = new System.Windows.Forms.Label();
            this.directoryValueLabel = new System.Windows.Forms.Label();
            this.progressBar1 = new DataClean.UI.NeonProgressBar();
            this.infoLine = new DataClean.UI.IconLabel();
            this.activityCard = new DataClean.UI.CardPanel();
            this.activityFilterCombo = new DataClean.UI.ThemedComboBox();
            this.activityList = new DataClean.UI.ThemedListView();
            this.colActivityStatus = new System.Windows.Forms.ColumnHeader();
            this.colActivityFile = new System.Windows.Forms.ColumnHeader();
            this.colActivitySize = new System.Windows.Forms.ColumnHeader();
            this.detailsLink = new System.Windows.Forms.LinkLabel();
            this.actionsPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.finishButton = new DataClean.UI.ThemedButton();
            this.startButton = new DataClean.UI.ThemedButton();
            this.cancelButton = new DataClean.UI.ThemedButton();
            this.pauseButton = new DataClean.UI.ThemedButton();
            this.retryButton = new DataClean.UI.ThemedButton();
            this.cleanupPage = new DataClean.UI.PagePanel();
            this.cleanupTitleLabel = new System.Windows.Forms.Label();
            this.cleanupSubtitleLabel = new System.Windows.Forms.Label();
            this.phasesCard = new DataClean.UI.CardPanel();
            this.phasesList = new DataClean.UI.ThemedListView();
            this.colPhaseStatus = new System.Windows.Forms.ColumnHeader();
            this.colPhaseName = new System.Windows.Forms.ColumnHeader();
            this.directoriesCard = new DataClean.UI.CardPanel();
            this.directoriesList = new DataClean.UI.ThemedListView();
            this.colDirectoryStatus = new System.Windows.Forms.ColumnHeader();
            this.colDirectoryName = new System.Windows.Forms.ColumnHeader();
            this.logPage = new DataClean.UI.PagePanel();
            this.logTitleLabel = new System.Windows.Forms.Label();
            this.logSubtitleLabel = new System.Windows.Forms.Label();
            this.logCard = new DataClean.UI.CardPanel();
            this.historyButton = new DataClean.UI.ThemedButton();
            this.detailTextBox = new DataClean.UI.ThemedTextBox();
            this.settingsPage = new DataClean.SettingsPage();
            this.footerPanel = new System.Windows.Forms.Panel();
            this.footerFlow = new System.Windows.Forms.FlowLayoutPanel();
            this.label3 = new System.Windows.Forms.Label();
            this.footerDivider1 = new System.Windows.Forms.Panel();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.footerDivider2 = new System.Windows.Forms.Panel();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.overviewPage.SuspendLayout();
            this.statsLayout.SuspendLayout();
            this.progressCard.SuspendLayout();
            this.activityCard.SuspendLayout();
            this.actionsPanel.SuspendLayout();
            this.cleanupPage.SuspendLayout();
            this.phasesCard.SuspendLayout();
            this.directoriesCard.SuspendLayout();
            this.logPage.SuspendLayout();
            this.logCard.SuspendLayout();
            this.footerPanel.SuspendLayout();
            this.footerFlow.SuspendLayout();
            this.SuspendLayout();
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Text = "DataClean";
            // 
            // overviewPage
            // 
            this.overviewPage.Controls.Add(this.headerTitleLabel);
            this.overviewPage.Controls.Add(this.heroLabel);
            this.overviewPage.Controls.Add(this.label1);
            this.overviewPage.Controls.Add(this.statsLayout);
            this.overviewPage.Controls.Add(this.progressCard);
            this.overviewPage.Controls.Add(this.activityCard);
            this.overviewPage.Controls.Add(this.detailsLink);
            this.overviewPage.Controls.Add(this.actionsPanel);
            this.overviewPage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.overviewPage.Location = new System.Drawing.Point(0, 0);
            this.overviewPage.Name = "overviewPage";
            this.overviewPage.Size = new System.Drawing.Size(802, 633);
            this.overviewPage.TabIndex = 0;
            // 
            // headerTitleLabel
            // 
            this.headerTitleLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.headerTitleLabel.AutoEllipsis = true;
            this.headerTitleLabel.Font = new System.Drawing.Font("Segoe UI Semibold", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.headerTitleLabel.Location = new System.Drawing.Point(20, 14);
            this.headerTitleLabel.Name = "headerTitleLabel";
            this.headerTitleLabel.Size = new System.Drawing.Size(767, 20);
            this.headerTitleLabel.TabIndex = 0;
            this.headerTitleLabel.Text = "Bereinigung läuft";
            this.headerTitleLabel.UseMnemonic = false;
            // 
            // heroLabel
            // 
            this.heroLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.heroLabel.AutoEllipsis = true;
            this.heroLabel.Font = new System.Drawing.Font("Segoe UI Semibold", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.heroLabel.Location = new System.Drawing.Point(16, 32);
            this.heroLabel.Name = "heroLabel";
            this.heroLabel.Size = new System.Drawing.Size(771, 48);
            this.heroLabel.TabIndex = 7;
            this.heroLabel.Text = "0 B freigegeben";
            this.heroLabel.UseMnemonic = false;
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoEllipsis = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(21, 80);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(766, 22);
            this.label1.TabIndex = 1;
            this.label1.Tag = "subtle";
            this.label1.Text = "In Arbeit …";
            this.label1.UseMnemonic = false;
            // 
            // statsLayout
            // 
            this.statsLayout.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.statsLayout.ColumnCount = 3;
            this.statsLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.statsLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.statsLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.statsLayout.Controls.Add(this.statRuntime, 0, 0);
            this.statsLayout.Controls.Add(this.statDeleted, 1, 0);
            this.statsLayout.Controls.Add(this.statSkipped, 2, 0);
            this.statsLayout.Location = new System.Drawing.Point(24, 112);
            this.statsLayout.Margin = new System.Windows.Forms.Padding(0);
            this.statsLayout.Name = "statsLayout";
            this.statsLayout.RowCount = 1;
            this.statsLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.statsLayout.Size = new System.Drawing.Size(763, 46);
            this.statsLayout.TabIndex = 2;
            // 
            // statRuntime
            // 
            this.statRuntime.Caption = "Laufzeit";
            this.statRuntime.Dock = System.Windows.Forms.DockStyle.Fill;
            this.statRuntime.Location = new System.Drawing.Point(0, 0);
            this.statRuntime.Margin = new System.Windows.Forms.Padding(0, 0, 9, 0);
            this.statRuntime.Name = "statRuntime";
            this.statRuntime.Size = new System.Drawing.Size(245, 46);
            this.statRuntime.TabIndex = 0;
            this.statRuntime.Value = "00:00";
            // 
            // statDeleted
            // 
            this.statDeleted.Caption = "Elemente gelöscht";
            this.statDeleted.Dock = System.Windows.Forms.DockStyle.Fill;
            this.statDeleted.LeadingRule = true;
            this.statDeleted.Location = new System.Drawing.Point(259, 0);
            this.statDeleted.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.statDeleted.Name = "statDeleted";
            this.statDeleted.Size = new System.Drawing.Size(244, 46);
            this.statDeleted.TabIndex = 1;
            // 
            // statSkipped
            // 
            this.statSkipped.Caption = "Übersprungen";
            this.statSkipped.Dock = System.Windows.Forms.DockStyle.Fill;
            this.statSkipped.LeadingRule = true;
            this.statSkipped.Location = new System.Drawing.Point(517, 0);
            this.statSkipped.Margin = new System.Windows.Forms.Padding(9, 0, 0, 0);
            this.statSkipped.Name = "statSkipped";
            this.statSkipped.Size = new System.Drawing.Size(246, 46);
            this.statSkipped.TabIndex = 2;
            // 
            // progressCard
            // 
            this.progressCard.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressCard.Controls.Add(this.phaseLabel);
            this.progressCard.Controls.Add(this.percentLabel);
            this.progressCard.Controls.Add(this.progressBar2);
            this.progressCard.Controls.Add(this.label2);
            this.progressCard.Controls.Add(this.directoryValueLabel);
            this.progressCard.Controls.Add(this.progressBar1);
            this.progressCard.Controls.Add(this.infoLine);
            this.progressCard.Location = new System.Drawing.Point(4, 168);
            this.progressCard.Name = "progressCard";
            this.progressCard.Size = new System.Drawing.Size(799, 102);
            this.progressCard.TabIndex = 3;
            this.progressCard.Title = "Gesamtfortschritt";
            // 
            // phaseLabel
            // 
            this.phaseLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.phaseLabel.AutoEllipsis = true;
            this.phaseLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.phaseLabel.Location = new System.Drawing.Point(146, 11);
            this.phaseLabel.Name = "phaseLabel";
            this.phaseLabel.Size = new System.Drawing.Size(506, 20);
            this.phaseLabel.TabIndex = 0;
            this.phaseLabel.Tag = "subtle";
            this.phaseLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.phaseLabel.UseMnemonic = false;
            // 
            // percentLabel
            // 
            this.percentLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.percentLabel.Font = new System.Drawing.Font("Segoe UI Semibold", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.percentLabel.Location = new System.Drawing.Point(670, 11);
            this.percentLabel.Name = "percentLabel";
            this.percentLabel.Size = new System.Drawing.Size(119, 20);
            this.percentLabel.TabIndex = 1;
            this.percentLabel.Text = "0 %";
            this.percentLabel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // progressBar2
            // 
            this.progressBar2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar2.Location = new System.Drawing.Point(20, 36);
            this.progressBar2.Maximum = 1000;
            this.progressBar2.Name = "progressBar2";
            this.progressBar2.Size = new System.Drawing.Size(763, 6);
            this.progressBar2.TabIndex = 2;
            this.progressBar2.TabStop = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(17, 49);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(115, 15);
            this.label2.TabIndex = 3;
            this.label2.Tag = "subtle";
            this.label2.Text = "Aktuelles Verzeichnis";
            // 
            // directoryValueLabel
            // 
            this.directoryValueLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.directoryValueLabel.AutoEllipsis = true;
            this.directoryValueLabel.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.directoryValueLabel.Location = new System.Drawing.Point(146, 49);
            this.directoryValueLabel.Name = "directoryValueLabel";
            this.directoryValueLabel.Size = new System.Drawing.Size(637, 18);
            this.directoryValueLabel.TabIndex = 4;
            this.directoryValueLabel.UseMnemonic = false;
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.Location = new System.Drawing.Point(20, 69);
            this.progressBar1.Muted = true;
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(763, 3);
            this.progressBar1.TabIndex = 5;
            this.progressBar1.TabStop = false;
            // 
            // infoLine
            // 
            this.infoLine.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.infoLine.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.infoLine.Location = new System.Drawing.Point(20, 76);
            this.infoLine.Name = "infoLine";
            this.infoLine.Size = new System.Drawing.Size(763, 22);
            this.infoLine.TabIndex = 6;
            this.infoLine.TabStop = false;
            this.infoLine.Visible = false;
            // 
            // activityCard
            // 
            this.activityCard.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.activityCard.Controls.Add(this.activityFilterCombo);
            this.activityCard.Controls.Add(this.activityList);
            this.activityCard.Location = new System.Drawing.Point(4, 278);
            this.activityCard.Name = "activityCard";
            this.activityCard.Size = new System.Drawing.Size(799, 291);
            this.activityCard.TabIndex = 4;
            this.activityCard.Title = "Aktivitätsprotokoll";
            // 
            // activityFilterCombo
            // 
            this.activityFilterCombo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.activityFilterCombo.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.activityFilterCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.activityFilterCombo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.activityFilterCombo.FormattingEnabled = true;
            this.activityFilterCombo.ItemHeight = 20;
            this.activityFilterCombo.Items.AddRange(new object[] {
            "Alle Einträge",
            "Gelöscht",
            "Übersprungen",
            "Fehler"});
            this.activityFilterCombo.Location = new System.Drawing.Point(617, 6);
            this.activityFilterCombo.Name = "activityFilterCombo";
            this.activityFilterCombo.Size = new System.Drawing.Size(166, 26);
            this.activityFilterCombo.TabIndex = 0;
            this.activityFilterCombo.SelectedIndexChanged += new System.EventHandler(this.ActivityFilterCombo_SelectedIndexChanged);
            // 
            // activityList
            // 
            this.activityList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.activityList.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.activityList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colActivityStatus,
            this.colActivityFile,
            this.colActivitySize});
            this.activityList.FillColumn = 1;
            this.activityList.FullRowSelect = true;
            this.activityList.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.activityList.Location = new System.Drawing.Point(20, 38);
            this.activityList.MultiSelect = false;
            this.activityList.Name = "activityList";
            this.activityList.OwnerDraw = true;
            this.activityList.PathColumn = 1;
            this.activityList.ShowSelection = false;
            this.activityList.Size = new System.Drawing.Size(763, 243);
            this.activityList.StatusColumn = 0;
            this.activityList.TabIndex = 1;
            this.activityList.UseCompatibleStateImageBehavior = false;
            this.activityList.View = System.Windows.Forms.View.Details;
            this.activityList.VirtualMode = true;
            this.activityList.RetrieveVirtualItem += new System.Windows.Forms.RetrieveVirtualItemEventHandler(this.ActivityList_RetrieveVirtualItem);
            // 
            // colActivityStatus
            // 
            this.colActivityStatus.Text = "Status";
            this.colActivityStatus.Width = 157;
            // 
            // colActivityFile
            // 
            this.colActivityFile.Text = "Datei";
            this.colActivityFile.Width = 508;
            // 
            // colActivitySize
            // 
            this.colActivitySize.Text = "Größe";
            this.colActivitySize.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.colActivitySize.Width = 98;
            // 
            // detailsLink
            // 
            this.detailsLink.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.detailsLink.AutoSize = true;
            this.detailsLink.Font = new System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.detailsLink.Location = new System.Drawing.Point(20, 587);
            this.detailsLink.Name = "detailsLink";
            this.detailsLink.Size = new System.Drawing.Size(96, 17);
            this.detailsLink.TabIndex = 5;
            this.detailsLink.TabStop = true;
            this.detailsLink.Text = "Details anzeigen";
            this.detailsLink.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.DetailsLink_LinkClicked);
            // 
            // actionsPanel
            // 
            this.actionsPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.actionsPanel.Controls.Add(this.finishButton);
            this.actionsPanel.Controls.Add(this.startButton);
            this.actionsPanel.Controls.Add(this.cancelButton);
            this.actionsPanel.Controls.Add(this.pauseButton);
            this.actionsPanel.Controls.Add(this.retryButton);
            this.actionsPanel.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.actionsPanel.Location = new System.Drawing.Point(160, 581);
            this.actionsPanel.Margin = new System.Windows.Forms.Padding(0);
            this.actionsPanel.Name = "actionsPanel";
            this.actionsPanel.Size = new System.Drawing.Size(627, 33);
            this.actionsPanel.TabIndex = 6;
            this.actionsPanel.WrapContents = false;
            // 
            // finishButton
            // 
            this.finishButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.finishButton.Kind = DataClean.UI.ButtonKind.Primary;
            this.finishButton.Location = new System.Drawing.Point(507, 0);
            this.finishButton.Margin = new System.Windows.Forms.Padding(14, 0, 0, 0);
            this.finishButton.Name = "finishButton";
            this.finishButton.Size = new System.Drawing.Size(120, 32);
            this.finishButton.TabIndex = 0;
            this.finishButton.Text = "Beenden";
            this.finishButton.UseVisualStyleBackColor = false;
            this.finishButton.Visible = false;
            this.finishButton.Click += new System.EventHandler(this.FinishButton_Click);
            // 
            // startButton
            // 
            this.startButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.startButton.Kind = DataClean.UI.ButtonKind.Primary;
            this.startButton.Location = new System.Drawing.Point(343, 0);
            this.startButton.Margin = new System.Windows.Forms.Padding(14, 0, 0, 0);
            this.startButton.Name = "startButton";
            this.startButton.Size = new System.Drawing.Size(150, 32);
            this.startButton.TabIndex = 4;
            this.startButton.Text = "Bereinigung starten";
            this.startButton.UseVisualStyleBackColor = false;
            this.startButton.Visible = false;
            this.startButton.Click += new System.EventHandler(this.StartButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cancelButton.Location = new System.Drawing.Point(373, 0);
            this.cancelButton.Margin = new System.Windows.Forms.Padding(14, 0, 0, 0);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(120, 32);
            this.cancelButton.TabIndex = 1;
            this.cancelButton.Text = "Abbrechen";
            this.cancelButton.UseVisualStyleBackColor = false;
            this.cancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // pauseButton
            // 
            this.pauseButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.pauseButton.Icon = DataClean.UI.IconKind.Pause;
            this.pauseButton.Location = new System.Drawing.Point(237, 0);
            this.pauseButton.Margin = new System.Windows.Forms.Padding(14, 0, 0, 0);
            this.pauseButton.Name = "pauseButton";
            this.pauseButton.Size = new System.Drawing.Size(122, 32);
            this.pauseButton.TabIndex = 2;
            this.pauseButton.Text = "Pausieren";
            this.pauseButton.UseVisualStyleBackColor = false;
            this.pauseButton.Click += new System.EventHandler(this.PauseButton_Click);
            // 
            // retryButton
            // 
            this.retryButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.retryButton.Location = new System.Drawing.Point(57, 0);
            this.retryButton.Margin = new System.Windows.Forms.Padding(14, 0, 0, 0);
            this.retryButton.Name = "retryButton";
            this.retryButton.Size = new System.Drawing.Size(166, 32);
            this.retryButton.TabIndex = 3;
            this.retryButton.Text = "Erweitertes Löschen";
            this.retryButton.UseVisualStyleBackColor = false;
            this.retryButton.Visible = false;
            this.retryButton.Click += new System.EventHandler(this.RetryButton_Click);
            // 
            // cleanupPage
            // 
            this.cleanupPage.Controls.Add(this.cleanupTitleLabel);
            this.cleanupPage.Controls.Add(this.cleanupSubtitleLabel);
            this.cleanupPage.Controls.Add(this.phasesCard);
            this.cleanupPage.Controls.Add(this.directoriesCard);
            this.cleanupPage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cleanupPage.Location = new System.Drawing.Point(0, 0);
            this.cleanupPage.Name = "cleanupPage";
            this.cleanupPage.Size = new System.Drawing.Size(802, 633);
            this.cleanupPage.TabIndex = 1;
            this.cleanupPage.Visible = false;
            // 
            // cleanupTitleLabel
            // 
            this.cleanupTitleLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cleanupTitleLabel.Font = new System.Drawing.Font("Segoe UI Semibold", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cleanupTitleLabel.Location = new System.Drawing.Point(20, 11);
            this.cleanupTitleLabel.Name = "cleanupTitleLabel";
            this.cleanupTitleLabel.Size = new System.Drawing.Size(767, 44);
            this.cleanupTitleLabel.TabIndex = 0;
            this.cleanupTitleLabel.Text = "Bereinigung";
            // 
            // cleanupSubtitleLabel
            // 
            this.cleanupSubtitleLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cleanupSubtitleLabel.Font = new System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cleanupSubtitleLabel.Location = new System.Drawing.Point(22, 54);
            this.cleanupSubtitleLabel.Name = "cleanupSubtitleLabel";
            this.cleanupSubtitleLabel.Size = new System.Drawing.Size(765, 26);
            this.cleanupSubtitleLabel.TabIndex = 1;
            this.cleanupSubtitleLabel.Tag = "subtle";
            this.cleanupSubtitleLabel.Text = "Phasen und Verzeichnisse dieses Laufs.";
            // 
            // phasesCard
            // 
            this.phasesCard.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.phasesCard.Controls.Add(this.phasesList);
            this.phasesCard.Location = new System.Drawing.Point(4, 93);
            this.phasesCard.Name = "phasesCard";
            this.phasesCard.Size = new System.Drawing.Size(799, 244);
            this.phasesCard.TabIndex = 2;
            this.phasesCard.Title = "Phasen";
            // 
            // phasesList
            // 
            this.phasesList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.phasesList.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.phasesList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colPhaseStatus,
            this.colPhaseName});
            this.phasesList.FullRowSelect = true;
            this.phasesList.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.phasesList.Location = new System.Drawing.Point(20, 38);
            this.phasesList.MultiSelect = false;
            this.phasesList.Name = "phasesList";
            this.phasesList.OwnerDraw = true;
            this.phasesList.ShowSelection = false;
            this.phasesList.Size = new System.Drawing.Size(763, 196);
            this.phasesList.StatusColumn = 0;
            this.phasesList.TabIndex = 0;
            this.phasesList.UseCompatibleStateImageBehavior = false;
            this.phasesList.View = System.Windows.Forms.View.Details;
            // 
            // colPhaseStatus
            // 
            this.colPhaseStatus.Text = "Status";
            this.colPhaseStatus.Width = 157;
            // 
            // colPhaseName
            // 
            this.colPhaseName.Text = "Phase";
            this.colPhaseName.Width = 606;
            // 
            // directoriesCard
            // 
            this.directoriesCard.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.directoriesCard.Controls.Add(this.directoriesList);
            this.directoriesCard.Location = new System.Drawing.Point(4, 351);
            this.directoriesCard.Name = "directoriesCard";
            this.directoriesCard.Size = new System.Drawing.Size(799, 263);
            this.directoriesCard.TabIndex = 3;
            this.directoriesCard.Title = "Verzeichnisse";
            // 
            // directoriesList
            // 
            this.directoriesList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.directoriesList.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.directoriesList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colDirectoryStatus,
            this.colDirectoryName});
            this.directoriesList.FullRowSelect = true;
            this.directoriesList.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.directoriesList.Location = new System.Drawing.Point(20, 38);
            this.directoriesList.MultiSelect = false;
            this.directoriesList.Name = "directoriesList";
            this.directoriesList.OwnerDraw = true;
            this.directoriesList.PathColumn = 1;
            this.directoriesList.ShowSelection = false;
            this.directoriesList.Size = new System.Drawing.Size(763, 215);
            this.directoriesList.StatusColumn = 0;
            this.directoriesList.TabIndex = 0;
            this.directoriesList.UseCompatibleStateImageBehavior = false;
            this.directoriesList.View = System.Windows.Forms.View.Details;
            // 
            // colDirectoryStatus
            // 
            this.colDirectoryStatus.Text = "Status";
            this.colDirectoryStatus.Width = 157;
            // 
            // colDirectoryName
            // 
            this.colDirectoryName.Text = "Verzeichnis";
            this.colDirectoryName.Width = 606;
            // 
            // logPage
            // 
            this.logPage.Controls.Add(this.logTitleLabel);
            this.logPage.Controls.Add(this.logSubtitleLabel);
            this.logPage.Controls.Add(this.logCard);
            this.logPage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.logPage.Location = new System.Drawing.Point(0, 0);
            this.logPage.Name = "logPage";
            this.logPage.Size = new System.Drawing.Size(802, 633);
            this.logPage.TabIndex = 2;
            this.logPage.Visible = false;
            // 
            // logTitleLabel
            // 
            this.logTitleLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.logTitleLabel.Font = new System.Drawing.Font("Segoe UI Semibold", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.logTitleLabel.Location = new System.Drawing.Point(20, 11);
            this.logTitleLabel.Name = "logTitleLabel";
            this.logTitleLabel.Size = new System.Drawing.Size(767, 44);
            this.logTitleLabel.TabIndex = 0;
            this.logTitleLabel.Text = "Protokoll";
            // 
            // logSubtitleLabel
            // 
            this.logSubtitleLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.logSubtitleLabel.Font = new System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.logSubtitleLabel.Location = new System.Drawing.Point(22, 54);
            this.logSubtitleLabel.Name = "logSubtitleLabel";
            this.logSubtitleLabel.Size = new System.Drawing.Size(765, 26);
            this.logSubtitleLabel.TabIndex = 1;
            this.logSubtitleLabel.Tag = "subtle";
            this.logSubtitleLabel.Text = "Ausführliches Protokoll dieses Laufs und frühere Läufe.";
            // 
            // logCard
            // 
            this.logCard.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.logCard.Controls.Add(this.historyButton);
            this.logCard.Controls.Add(this.detailTextBox);
            this.logCard.Location = new System.Drawing.Point(4, 93);
            this.logCard.Name = "logCard";
            this.logCard.Size = new System.Drawing.Size(799, 521);
            this.logCard.TabIndex = 2;
            this.logCard.Title = "Laufprotokoll";
            // 
            // historyButton
            // 
            this.historyButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.historyButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.historyButton.Location = new System.Drawing.Point(643, 6);
            this.historyButton.Name = "historyButton";
            this.historyButton.Size = new System.Drawing.Size(140, 26);
            this.historyButton.TabIndex = 0;
            this.historyButton.Text = "Lauf-Historie";
            this.historyButton.UseVisualStyleBackColor = false;
            this.historyButton.Click += new System.EventHandler(this.HistoryButton_Click);
            // 
            // detailTextBox
            // 
            this.detailTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.detailTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.detailTextBox.Font = new System.Drawing.Font("Consolas", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.detailTextBox.Location = new System.Drawing.Point(20, 38);
            this.detailTextBox.Multiline = true;
            this.detailTextBox.Name = "detailTextBox";
            this.detailTextBox.ReadOnly = true;
            this.detailTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.detailTextBox.Size = new System.Drawing.Size(763, 473);
            this.detailTextBox.TabIndex = 1;
            this.detailTextBox.WordWrap = false;
            // 
            // settingsPage
            // 
            this.settingsPage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.settingsPage.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.settingsPage.Location = new System.Drawing.Point(0, 0);
            this.settingsPage.Name = "settingsPage";
            this.settingsPage.Size = new System.Drawing.Size(802, 633);
            this.settingsPage.TabIndex = 4;
            this.settingsPage.Visible = false;
            // 
            // footerPanel
            // 
            this.footerPanel.Controls.Add(this.footerFlow);
            this.footerPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.footerPanel.Location = new System.Drawing.Point(0, 633);
            this.footerPanel.Name = "footerPanel";
            this.footerPanel.Size = new System.Drawing.Size(802, 44);
            this.footerPanel.TabIndex = 3;
            // 
            // footerFlow
            // 
            this.footerFlow.Controls.Add(this.label3);
            this.footerFlow.Controls.Add(this.footerDivider1);
            this.footerFlow.Controls.Add(this.linkLabel1);
            this.footerFlow.Controls.Add(this.footerDivider2);
            this.footerFlow.Controls.Add(this.linkLabel2);
            this.footerFlow.Dock = System.Windows.Forms.DockStyle.Fill;
            this.footerFlow.Location = new System.Drawing.Point(0, 0);
            this.footerFlow.Name = "footerFlow";
            this.footerFlow.Padding = new System.Windows.Forms.Padding(18, 13, 0, 0);
            this.footerFlow.Size = new System.Drawing.Size(802, 44);
            this.footerFlow.TabIndex = 0;
            this.footerFlow.WrapContents = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(18, 13);
            this.label3.Margin = new System.Windows.Forms.Padding(0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(147, 17);
            this.label3.TabIndex = 0;
            this.label3.Text = "Daniel Jörg Schuppelius";
            // 
            // footerDivider1
            // 
            this.footerDivider1.Location = new System.Drawing.Point(179, 13);
            this.footerDivider1.Margin = new System.Windows.Forms.Padding(14, 0, 14, 0);
            this.footerDivider1.Name = "footerDivider1";
            this.footerDivider1.Size = new System.Drawing.Size(1, 18);
            this.footerDivider1.TabIndex = 1;
            this.footerDivider1.Tag = "separator";
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabel1.Location = new System.Drawing.Point(194, 13);
            this.linkLabel1.Margin = new System.Windows.Forms.Padding(0);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(100, 17);
            this.linkLabel1.TabIndex = 2;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "schuppelius.org";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LinkLabel1_LinkClicked);
            // 
            // footerDivider2
            // 
            this.footerDivider2.Location = new System.Drawing.Point(308, 13);
            this.footerDivider2.Margin = new System.Windows.Forms.Padding(14, 0, 14, 0);
            this.footerDivider2.Name = "footerDivider2";
            this.footerDivider2.Size = new System.Drawing.Size(1, 18);
            this.footerDivider2.TabIndex = 3;
            this.footerDivider2.Tag = "separator";
            // 
            // linkLabel2
            // 
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabel2.Location = new System.Drawing.Point(323, 13);
            this.linkLabel2.Margin = new System.Windows.Forms.Padding(0);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(134, 17);
            this.linkLabel2.TabIndex = 4;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.Text = "info@schuppelius.org";
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LinkLabel2_LinkClicked);
            // 
            // DeleteForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(802, 677);
            this.Controls.Add(this.overviewPage);
            this.Controls.Add(this.cleanupPage);
            this.Controls.Add(this.logPage);
            this.Controls.Add(this.settingsPage);
            this.Controls.Add(this.footerPanel);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "DeleteForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Daniel Jörg Schuppelius - Informationstechnische Unternehmensdienstleistungen";
            this.TopMost = true;
            this.Shown += new System.EventHandler(this.DeleteForm_Shown);
            this.overviewPage.ResumeLayout(false);
            this.overviewPage.PerformLayout();
            this.statsLayout.ResumeLayout(false);
            this.progressCard.ResumeLayout(false);
            this.progressCard.PerformLayout();
            this.activityCard.ResumeLayout(false);
            this.actionsPanel.ResumeLayout(false);
            this.cleanupPage.ResumeLayout(false);
            this.phasesCard.ResumeLayout(false);
            this.directoriesCard.ResumeLayout(false);
            this.logPage.ResumeLayout(false);
            this.logCard.ResumeLayout(false);
            this.logCard.PerformLayout();
            this.footerPanel.ResumeLayout(false);
            this.footerFlow.ResumeLayout(false);
            this.footerFlow.PerformLayout();
            this.ResumeLayout(false);

        }
        #endregion

        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private DataClean.UI.PagePanel overviewPage;
        private System.Windows.Forms.Label headerTitleLabel;
        private System.Windows.Forms.Label heroLabel;
        private System.Windows.Forms.Label label1; // Statuszeile unter der Überschrift
        private System.Windows.Forms.TableLayoutPanel statsLayout;
        private DataClean.UI.StatCard statRuntime;
        private DataClean.UI.StatCard statDeleted;
        private DataClean.UI.StatCard statSkipped;
        private DataClean.UI.CardPanel progressCard;
        private System.Windows.Forms.Label phaseLabel;
        private System.Windows.Forms.Label percentLabel;
        private DataClean.UI.NeonProgressBar progressBar2; // Gesamtfortschritt
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label directoryValueLabel;
        private DataClean.UI.NeonProgressBar progressBar1; // Fortschritt des aktuellen Verzeichnisses
        private DataClean.UI.IconLabel infoLine;
        private DataClean.UI.CardPanel activityCard;
        private DataClean.UI.ThemedComboBox activityFilterCombo;
        private DataClean.UI.ThemedListView activityList;
        private System.Windows.Forms.ColumnHeader colActivityStatus;
        private System.Windows.Forms.ColumnHeader colActivityFile;
        private System.Windows.Forms.ColumnHeader colActivitySize;
        private System.Windows.Forms.LinkLabel detailsLink;
        private System.Windows.Forms.FlowLayoutPanel actionsPanel;
        private DataClean.UI.ThemedButton finishButton;
        private DataClean.UI.ThemedButton startButton;
        private DataClean.UI.ThemedButton cancelButton;
        private DataClean.UI.ThemedButton pauseButton;
        private DataClean.UI.ThemedButton retryButton; // Neustart mit Administratorrechten
        private DataClean.UI.PagePanel cleanupPage;
        private System.Windows.Forms.Label cleanupTitleLabel;
        private System.Windows.Forms.Label cleanupSubtitleLabel;
        private DataClean.UI.CardPanel phasesCard;
        private DataClean.UI.ThemedListView phasesList;
        private System.Windows.Forms.ColumnHeader colPhaseStatus;
        private System.Windows.Forms.ColumnHeader colPhaseName;
        private DataClean.UI.CardPanel directoriesCard;
        private DataClean.UI.ThemedListView directoriesList;
        private System.Windows.Forms.ColumnHeader colDirectoryStatus;
        private System.Windows.Forms.ColumnHeader colDirectoryName;
        private DataClean.UI.PagePanel logPage;
        private System.Windows.Forms.Label logTitleLabel;
        private System.Windows.Forms.Label logSubtitleLabel;
        private DataClean.UI.CardPanel logCard;
        private DataClean.UI.ThemedButton historyButton;
        private DataClean.UI.ThemedTextBox detailTextBox; // Ausführliches Laufprotokoll
        private DataClean.SettingsPage settingsPage;
        private System.Windows.Forms.Panel footerPanel;
        private System.Windows.Forms.FlowLayoutPanel footerFlow;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Panel footerDivider1;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.Panel footerDivider2;
        private System.Windows.Forms.LinkLabel linkLabel2;
    }
}
