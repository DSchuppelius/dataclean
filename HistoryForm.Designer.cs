namespace DataClean {
    partial class HistoryForm {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        private void InitializeComponent() {
            this.lblHeading = new System.Windows.Forms.Label();
            this.lblSubheading = new System.Windows.Forms.Label();
            this.cardHistory = new DataClean.UI.CardPanel();
            this.lstHistory = new DataClean.UI.ThemedListView();
            this.colStart = new System.Windows.Forms.ColumnHeader();
            this.colDuration = new System.Windows.Forms.ColumnHeader();
            this.colFreed = new System.Windows.Forms.ColumnHeader();
            this.colDeleted = new System.Windows.Forms.ColumnHeader();
            this.colSkipped = new System.Windows.Forms.ColumnHeader();
            this.colFailed = new System.Windows.Forms.ColumnHeader();
            this.colStatus = new System.Windows.Forms.ColumnHeader();
            this.btnClear = new DataClean.UI.ThemedButton();
            this.btnClose = new DataClean.UI.ThemedButton();
            this.cardHistory.SuspendLayout();
            this.SuspendLayout();
            //
            // lblHeading
            //
            this.lblHeading.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblHeading.Font = new System.Drawing.Font("Segoe UI", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHeading.Location = new System.Drawing.Point(20, 8);
            this.lblHeading.Name = "lblHeading";
            this.lblHeading.Size = new System.Drawing.Size(716, 44);
            this.lblHeading.TabIndex = 3;
            this.lblHeading.Text = "Lauf-Historie";
            //
            // lblSubheading
            //
            this.lblSubheading.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblSubheading.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSubheading.Location = new System.Drawing.Point(22, 50);
            this.lblSubheading.Name = "lblSubheading";
            this.lblSubheading.Size = new System.Drawing.Size(714, 26);
            this.lblSubheading.TabIndex = 4;
            this.lblSubheading.Tag = "subtle";
            this.lblSubheading.Text = "Abgeschlossene Bereinigungsläufe auf diesem Rechner.";
            //
            // cardHistory
            //
            this.cardHistory.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cardHistory.Controls.Add(this.lstHistory);
            this.cardHistory.Location = new System.Drawing.Point(24, 86);
            this.cardHistory.Name = "cardHistory";
            this.cardHistory.Size = new System.Drawing.Size(712, 346);
            this.cardHistory.TabIndex = 0;
            //
            // lstHistory
            //
            this.lstHistory.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lstHistory.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colStart,
            this.colDuration,
            this.colFreed,
            this.colDeleted,
            this.colSkipped,
            this.colFailed,
            this.colStatus});
            this.lstHistory.Location = new System.Drawing.Point(12, 12);
            this.lstHistory.Name = "lstHistory";
            this.lstHistory.ShowItemToolTips = true;
            this.lstHistory.Size = new System.Drawing.Size(688, 322);
            this.lstHistory.StatusColumn = 6;
            this.lstHistory.TabIndex = 0;
            //
            // colStart
            //
            this.colStart.Text = "Start";
            this.colStart.Width = 124;
            //
            // colDuration
            //
            this.colDuration.Text = "Dauer";
            this.colDuration.Width = 78;
            //
            // colFreed
            //
            this.colFreed.Text = "Freigegeben";
            this.colFreed.Width = 96;
            //
            // colDeleted
            //
            this.colDeleted.Text = "Gelöscht";
            this.colDeleted.Width = 78;
            //
            // colSkipped
            //
            this.colSkipped.Text = "Übersprungen";
            this.colSkipped.Width = 112;
            //
            // colFailed
            //
            this.colFailed.Text = "Fehler";
            this.colFailed.Width = 64;
            //
            // colStatus
            //
            this.colStatus.Text = "Status";
            this.colStatus.Width = 136;
            //
            // btnClear
            //
            this.btnClear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnClear.Icon = DataClean.UI.IconKind.Trash;
            this.btnClear.Kind = DataClean.UI.ButtonKind.Danger;
            this.btnClear.Location = new System.Drawing.Point(24, 448);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(160, 32);
            this.btnClear.TabIndex = 1;
            this.btnClear.Text = "Historie löschen";
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            //
            // btnClose
            //
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.Kind = DataClean.UI.ButtonKind.Primary;
            this.btnClose.Location = new System.Drawing.Point(626, 448);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(110, 32);
            this.btnClose.TabIndex = 2;
            this.btnClose.Text = "Schließen";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            //
            // HistoryForm
            //
            this.AcceptButton = this.btnClose;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnClose;
            this.ClientSize = new System.Drawing.Size(760, 500);
            this.Controls.Add(this.lblHeading);
            this.Controls.Add(this.lblSubheading);
            this.Controls.Add(this.cardHistory);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.btnClose);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "HistoryForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "DataClean - Lauf-Historie";
            this.cardHistory.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.Label lblHeading;
        private System.Windows.Forms.Label lblSubheading;
        private DataClean.UI.CardPanel cardHistory;
        private DataClean.UI.ThemedListView lstHistory;
        private System.Windows.Forms.ColumnHeader colStart;
        private System.Windows.Forms.ColumnHeader colDuration;
        private System.Windows.Forms.ColumnHeader colFreed;
        private System.Windows.Forms.ColumnHeader colDeleted;
        private System.Windows.Forms.ColumnHeader colSkipped;
        private System.Windows.Forms.ColumnHeader colFailed;
        private System.Windows.Forms.ColumnHeader colStatus;
        private DataClean.UI.ThemedButton btnClear;
        private DataClean.UI.ThemedButton btnClose;
    }
}
