using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace DataClean.UI {
    /// <summary>Statusdarstellung einer Tabellenzeile (Icon und Farbe).</summary>
    public enum StatusKind {
        None,
        Success,
        Skipped,
        Error,
        Running,
        Pending,
        Preview,
        Warning
    }

    /// <summary>
    /// Selbst gezeichnete Tabelle im Dashboard-Design. Zeilen, deren <see cref="ListViewItem.Tag"/>
    /// ein <see cref="StatusKind"/> ist, zeigen in der Statusspalte Icon und Farbe.
    /// </summary>
    public class ThemedListView : ListView, IThemedControl {
        private const int RowHeight = 26;
        private const int WM_PAINT = 0x000F;
        private const int WM_NCCALCSIZE = 0x0083;
        private const int LVM_GETHEADER = 0x101F;
        private readonly ImageList rowHeightList;
        private int statusColumn = -1;
        private int pathColumn = -1;
        private int fillColumn = -1;
        private bool showSelection = true;
        private bool fittingColumns;
        private int hoverIndex = -1;

        public ThemedListView() {
            DoubleBuffered = true;
            OwnerDraw = true;
            View = View.Details;
            FullRowSelect = true;
            HeaderStyle = ColumnHeaderStyle.Nonclickable;
            BorderStyle = BorderStyle.None;
            MultiSelect = false;
            UseCompatibleStateImageBehavior = false;

            // Die Zeilenhöhe einer ListView richtet sich nach der Bildliste.
            rowHeightList = new ImageList { ImageSize = new Size(1, RowHeight) };
            SmallImageList = rowHeightList;
            ApplyTheme();
        }

        /// <summary>Spalte, in der das Status-Icon gezeichnet wird (-1 = keine).</summary>
        [DefaultValue(-1)]
        public int StatusColumn {
            get { return statusColumn; }
            set { statusColumn = value; Invalidate(); }
        }

        /// <summary>Spalte mit Dateipfaden; gekürzt wird in der Pfadmitte (-1 = keine).</summary>
        [DefaultValue(-1)]
        public int PathColumn {
            get { return pathColumn; }
            set { pathColumn = value; Invalidate(); }
        }

        /// <summary>Spalte, die die verbleibende Breite ausfüllt (-1 = letzte Spalte).</summary>
        [DefaultValue(-1)]
        public int FillColumn {
            get { return fillColumn; }
            set { fillColumn = value; FitColumns(); }
        }

        [DefaultValue(true)]
        public bool ShowSelection {
            get { return showSelection; }
            set { showSelection = value; Invalidate(); }
        }

        public void ApplyTheme() {
            BackColor = AppTheme.Surface;
            ForeColor = AppTheme.Text;
            if (IsHandleCreated) {
                AppTheme.ApplyNativeTheme(this);
            }
            Invalidate();
        }

        public static IconKind IconFor(StatusKind status) {
            switch (status) {
                case StatusKind.Success: return IconKind.CheckCircle;
                case StatusKind.Skipped: return IconKind.Skip;
                case StatusKind.Error: return IconKind.ErrorCircle;
                case StatusKind.Running: return IconKind.Play;
                case StatusKind.Pending: return IconKind.Dot;
                case StatusKind.Preview: return IconKind.Info;
                case StatusKind.Warning: return IconKind.Warning;
                default: return IconKind.None;
            }
        }

        public static Color ColorFor(StatusKind status) {
            switch (status) {
                case StatusKind.Success: return AppTheme.Success;
                case StatusKind.Skipped:
                case StatusKind.Warning: return AppTheme.Warning;
                case StatusKind.Error: return AppTheme.Error;
                case StatusKind.Running:
                case StatusKind.Preview: return AppTheme.Accent;
                case StatusKind.Pending: return AppTheme.SubtleText;
                default: return AppTheme.Text;
            }
        }

        protected override void OnHandleCreated(EventArgs e) {
            base.OnHandleCreated(e);
            AppTheme.ApplyNativeTheme(this);
            FitColumns();
        }

        protected override void OnClientSizeChanged(EventArgs e) {
            base.OnClientSizeChanged(e);
            FitColumns();
        }

        /// <summary>Passt die Füllspalte an, sodass keine Lücke rechts im Tabellenkopf bleibt.</summary>
        public void FitColumns() {
            if (fittingColumns || Columns.Count == 0 || !IsHandleCreated) {
                return;
            }

            int fill = fillColumn >= 0 && fillColumn < Columns.Count ? fillColumn : Columns.Count - 1;
            int others = 0;
            for (int i = 0; i < Columns.Count; i++) {
                if (i != fill) {
                    others += Columns[i].Width;
                }
            }

            // Die vertikale Scrollleiste wird vorab eingerechnet: Die ListView aktualisiert sie erst
            // nach dem Setzen der Zeilenzahl und würde sonst kurzzeitig auch horizontal scrollen.
            int rowCount = VirtualMode ? VirtualListSize : Items.Count;
            bool needsScrollBar = HeaderHeight() + rowCount * RowHeight > Height;
            int available = Width - (needsScrollBar ? SystemInformation.VerticalScrollBarWidth : 0);
            int target = Math.Max(40, available - others);
            if (Columns[fill].Width != target) {
                fittingColumns = true;
                try {
                    Columns[fill].Width = target;
                } finally {
                    fittingColumns = false;
                }
            }
        }

        protected override void OnDrawColumnHeader(DrawListViewColumnHeaderEventArgs e) {
            using (var brush = new SolidBrush(AppTheme.TableHeader)) {
                e.Graphics.FillRectangle(brush, e.Bounds);
            }

            int last = Columns.Count - 1;
            bool beforeRightAligned = e.ColumnIndex == last - 1 && Columns[last].TextAlign == HorizontalAlignment.Right;
            if (e.ColumnIndex < last && !beforeRightAligned) {
                using (var pen = new Pen(AppTheme.CardBorder)) {
                    e.Graphics.DrawLine(pen, e.Bounds.Right - 1, e.Bounds.Y, e.Bounds.Right - 1, e.Bounds.Bottom - 1);
                }
            }

            var text = new Rectangle(e.Bounds.X + 10, e.Bounds.Y, Math.Max(0, e.Bounds.Width - 20), e.Bounds.Height);
            TextRenderer.DrawText(e.Graphics, e.Header.Text, Font, text, AppTheme.Text, CellFlags(e.Header.TextAlign, false));
        }

        protected override void OnDrawItem(DrawListViewItemEventArgs e) {
            e.DrawDefault = false;
        }

        protected override void OnDrawSubItem(DrawListViewSubItemEventArgs e) {
            int width = Columns[e.ColumnIndex].Width;
            var cell = new Rectangle(e.Bounds.X, e.Bounds.Y, width, e.Bounds.Height);

            bool selected = showSelection && e.Item.Selected;
            Color background = selected ? AppTheme.NavSelected : (e.ItemIndex == hoverIndex ? AppTheme.NavHover : AppTheme.Surface);
            using (var brush = new SolidBrush(background)) {
                e.Graphics.FillRectangle(brush, cell);
            }
            using (var pen = new Pen(AppTheme.TableLine)) {
                e.Graphics.DrawLine(pen, cell.X, cell.Bottom - 1, cell.Right, cell.Bottom - 1);
            }

            Color fore = AppTheme.Text;
            int left = cell.X + 10;

            if (e.ColumnIndex == statusColumn && e.Item.Tag is StatusKind status && status != StatusKind.None) {
                fore = ColorFor(status);
                const int iconSize = 16;
                Icons.Draw(e.Graphics, IconFor(status), new RectangleF(left, cell.Y + (cell.Height - iconSize) / 2f, iconSize, iconSize), fore);
                left += iconSize + 10;
            }

            var text = new Rectangle(left, cell.Y, Math.Max(0, cell.Right - left - 10), cell.Height);
            TextRenderer.DrawText(e.Graphics, e.SubItem.Text, Font, text, fore,
                CellFlags(Columns[e.ColumnIndex].TextAlign, e.ColumnIndex == pathColumn));
        }

        private static TextFormatFlags CellFlags(HorizontalAlignment align, bool isPath) {
            TextFormatFlags flags = TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine | TextFormatFlags.VerticalCenter;
            flags |= isPath ? TextFormatFlags.PathEllipsis : TextFormatFlags.EndEllipsis;
            if (align == HorizontalAlignment.Right) {
                flags |= TextFormatFlags.Right;
            } else if (align == HorizontalAlignment.Center) {
                flags |= TextFormatFlags.HorizontalCenter;
            }
            return flags;
        }

        protected override void WndProc(ref Message m) {
            // Die Spalten passen immer exakt in die Breite (FitColumns); ohne dieses Entfernen
            // bliebe nach dem Einblenden der vertikalen Leiste auch eine horizontale stehen.
            if (m.Msg == WM_NCCALCSIZE && IsHandleCreated) {
                int style = NativeMethods.GetWindowLong(Handle, NativeMethods.GWL_STYLE);
                if ((style & NativeMethods.WS_HSCROLL) != 0) {
                    NativeMethods.SetWindowLong(Handle, NativeMethods.GWL_STYLE, style & ~NativeMethods.WS_HSCROLL);
                }
            }

            base.WndProc(ref m);
            if (m.Msg == WM_PAINT && IsHandleCreated) {
                PaintEmptyArea();
            }
        }

        // Der Explorer-Stil zeichnet im Details-Modus Spaltenlinien in den leeren Bereich
        // unter der letzten Zeile; diese Fläche wird hier wieder einfarbig gefüllt.
        private void PaintEmptyArea() {
            int count = VirtualMode ? VirtualListSize : Items.Count;
            int top;
            if (count > 0) {
                top = GetItemRect(count - 1).Bottom;
            } else {
                top = HeaderHeight();
            }
            if (top >= ClientSize.Height) {
                return;
            }
            using (Graphics g = Graphics.FromHwnd(Handle))
            using (var brush = new SolidBrush(AppTheme.Surface)) {
                g.FillRectangle(brush, 0, Math.Max(0, top), ClientSize.Width, ClientSize.Height - Math.Max(0, top));
            }
        }

        private int HeaderHeight() {
            IntPtr header = NativeMethods.SendMessage(Handle, LVM_GETHEADER, IntPtr.Zero, IntPtr.Zero);
            NativeMethods.RECT rect;
            return header != IntPtr.Zero && NativeMethods.GetWindowRect(header, out rect) ? rect.Bottom - rect.Top : 0;
        }

        protected override void OnMouseMove(MouseEventArgs e) {
            base.OnMouseMove(e);
            ListViewItem item = VirtualListSize > 0 || Items.Count > 0 ? GetItemAt(e.X, e.Y) : null;
            SetHoverIndex(item != null ? item.Index : -1);
        }

        protected override void OnMouseLeave(EventArgs e) {
            base.OnMouseLeave(e);
            SetHoverIndex(-1);
        }

        private void SetHoverIndex(int index) {
            if (index == hoverIndex) {
                return;
            }
            InvalidateRow(hoverIndex);
            hoverIndex = index;
            InvalidateRow(hoverIndex);
        }

        private void InvalidateRow(int index) {
            int count = VirtualMode ? VirtualListSize : Items.Count;
            if (index >= 0 && index < count) {
                Invalidate(GetItemRect(index));
            }
        }

        protected override void Dispose(bool disposing) {
            if (disposing) {
                rowHeightList.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
