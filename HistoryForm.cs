using System;
using System.Windows.Forms;

namespace DataClean {
    /// <summary>
    /// Zeigt die Historie abgeschlossener Bereinigungsläufe in einer Liste an.
    /// </summary>
    public partial class HistoryForm : UI.FuturisticForm {
        public HistoryForm() {
            EnableMinimizeButton = false;
            AllowResize = true;
            CaptionSubtitle = "Lauf-Historie";
            InitializeComponent();
            LoadHistory();
        }

        private void LoadHistory() {
            lstHistory.BeginUpdate();
            lstHistory.Items.Clear();

            foreach (var entry in Services.RunHistoryStore.LoadAll()) {
                var item = new ListViewItem(entry.StartTime.ToString("yyyy-MM-dd HH:mm"));
                TimeSpan duration = entry.EndTime - entry.StartTime;
                item.SubItems.Add(duration.ToString(@"hh\:mm\:ss"));
                item.SubItems.Add(FormatBytes(entry.FreedBytes));
                item.SubItems.Add(entry.DeletedItems.ToString());
                item.SubItems.Add(entry.SkippedItems.ToString());
                item.SubItems.Add(entry.FailedItems.ToString());
                item.SubItems.Add(BuildStatus(entry, out UI.StatusKind status));
                item.Tag = status;
                item.ToolTipText = entry.Tasks;
                lstHistory.Items.Add(item);
            }

            if (lstHistory.Items.Count == 0) {
                lstHistory.Items.Add(new ListViewItem("Keine Läufe aufgezeichnet."));
            }

            lstHistory.EndUpdate();
        }

        private static string BuildStatus(Services.RunHistoryEntry entry, out UI.StatusKind status) {
            if (entry.Canceled) {
                status = UI.StatusKind.Warning;
                return "Abgebrochen";
            }
            if (entry.DryRun) {
                status = UI.StatusKind.Preview;
                return "Vorschau";
            }
            if (entry.FailedItems > 0) {
                status = UI.StatusKind.Error;
                return "Mit Fehlern";
            }
            status = UI.StatusKind.Success;
            return entry.Unattended ? "Automatisch" : "Erfolgreich";
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

        private void btnClear_Click(object sender, EventArgs e) {
            if (UI.AppDialog.Show(this, "Gesamte Lauf-Historie löschen?", "Historie löschen",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) {
                return;
            }

            Services.RunHistoryStore.Clear();
            LoadHistory();
        }

        private void btnClose_Click(object sender, EventArgs e) {
            Close();
        }
    }
}
