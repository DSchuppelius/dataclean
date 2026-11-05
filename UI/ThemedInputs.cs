using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace DataClean.UI {
    /// <summary>Flacher, abgerundeter Button mit optionalem Icon.</summary>
    public class ThemedButton : Button, IThemedControl {
        private ButtonKind kind = ButtonKind.Secondary;
        private IconKind icon = IconKind.None;
        private bool hovering;
        private bool pressed;

        public ThemedButton() {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            UseVisualStyleBackColor = false;
            Cursor = Cursors.Hand;
        }

        [DefaultValue(ButtonKind.Secondary)]
        public ButtonKind Kind {
            get { return kind; }
            set { kind = value; Invalidate(); }
        }

        [DefaultValue(IconKind.None)]
        public IconKind Icon {
            get { return icon; }
            set { icon = value; Invalidate(); }
        }

        public void ApplyTheme() {
            Invalidate();
        }

        protected override void OnMouseEnter(EventArgs e) {
            hovering = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e) {
            hovering = false;
            Invalidate();
            base.OnMouseLeave(e);
        }

        protected override void OnMouseDown(MouseEventArgs e) {
            pressed = e.Button == MouseButtons.Left;
            Invalidate();
            base.OnMouseDown(e);
        }

        protected override void OnMouseUp(MouseEventArgs e) {
            pressed = false;
            Invalidate();
            base.OnMouseUp(e);
        }

        protected override void OnEnabledChanged(EventArgs e) {
            Invalidate();
            base.OnEnabledChanged(e);
        }

        protected override void OnPaint(PaintEventArgs e) {
            Graphics g = e.Graphics;
            g.Clear(AppTheme.EffectiveBack(this));
            g.SmoothingMode = SmoothingMode.AntiAlias;

            Color fill, hover, border, fore;
            switch (kind) {
                case ButtonKind.Primary:
                    fill = AppTheme.Primary; hover = AppTheme.PrimaryHover;
                    border = AppTheme.Primary; fore = AppTheme.OnAccent;
                    break;
                case ButtonKind.Danger:
                    fill = AppTheme.Danger; hover = AppTheme.DangerHover;
                    border = AppTheme.InputBorder; fore = AppTheme.Error;
                    break;
                default:
                    fill = AppTheme.Secondary; hover = AppTheme.SecondaryHover;
                    border = AppTheme.SecondaryBorder; fore = AppTheme.Text;
                    break;
            }

            if (!Enabled) {
                fill = AppTheme.Secondary;
                border = Color.FromArgb(120, AppTheme.InputBorder);
                // TextRenderer ignoriert Transparenz, die Farbe muss daher deckend gemischt sein.
                fore = AppTheme.Blend(AppTheme.Text, fill, 0.45f);
            }

            Color background = Enabled && (hovering || pressed) ? hover : fill;
            bool focusCue = Focused && ShowFocusCues && Enabled;

            using (var path = AppTheme.RoundedPath(new Rectangle(0, 0, Width - 1, Height - 1), 3))
            using (var brush = new SolidBrush(background))
            using (var pen = new Pen(focusCue ? AppTheme.Accent : border, focusCue ? 2f : 1f)) {
                g.FillPath(brush, path);
                g.DrawPath(pen, path);
            }

            const int iconSize = 16;
            const int iconGap = 9;
            Size textSize = TextRenderer.MeasureText(g, Text, Font, Size, TextFormatFlags.NoPrefix);
            int contentWidth = textSize.Width + (icon != IconKind.None ? iconSize + iconGap : 0);
            int x = Math.Max(4, (Width - contentWidth) / 2);

            if (icon != IconKind.None) {
                Icons.Draw(g, icon, new RectangleF(x, (Height - iconSize) / 2f, iconSize, iconSize), fore);
                x += iconSize + iconGap;
            }

            TextRenderer.DrawText(g, Text, Font, new Rectangle(x, 0, Math.Max(0, Width - x - 2), Height), fore,
                TextFormatFlags.NoPrefix | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis);
        }
    }

    /// <summary>CheckBox mit gezeichnetem Kästchen (die native Glyphe bleibt sonst hell).</summary>
    public class ThemedCheckBox : CheckBox, IThemedControl {
        private const int BoxSize = 16;
        private const int TextGap = 8;
        private bool hovering;

        public ThemedCheckBox() {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            UseVisualStyleBackColor = false;
            Cursor = Cursors.Hand;
        }

        public void ApplyTheme() {
            Invalidate();
        }

        public override Size GetPreferredSize(Size proposedSize) {
            Size text = TextRenderer.MeasureText(Text, Font, Size.Empty, TextFormatFlags.NoPrefix);
            return new Size(BoxSize + TextGap + text.Width + 2, Math.Max(22, text.Height + 4));
        }

        /// <summary>Zeichnet das Kästchen; auch die Zeilen einer Tabelle verwenden es.</summary>
        internal static void PaintBox(Graphics g, Rectangle box, bool isChecked, bool enabled, bool hovering) {
            Color accent = enabled ? AppTheme.Accent : Color.FromArgb(110, AppTheme.Accent);
            Color border = isChecked || (hovering && enabled) ? accent : AppTheme.InputBorder;

            using (var path = AppTheme.RoundedPath(box, 2))
            using (var brush = new SolidBrush(isChecked ? accent : AppTheme.SurfaceAlt))
            using (var pen = new Pen(border)) {
                g.FillPath(brush, path);
                g.DrawPath(pen, path);
            }

            if (isChecked) {
                using (var pen = new Pen(AppTheme.OnAccent, 2f) {
                    StartCap = LineCap.Round,
                    EndCap = LineCap.Round,
                    LineJoin = LineJoin.Round
                }) {
                    g.DrawLines(pen, new[] {
                        new PointF(box.X + 3.6f, box.Y + 8f),
                        new PointF(box.X + 6.6f, box.Y + 11f),
                        new PointF(box.X + 11.8f, box.Y + 4.8f)
                    });
                }
            }
        }

        protected override void OnMouseEnter(EventArgs e) {
            hovering = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e) {
            hovering = false;
            Invalidate();
            base.OnMouseLeave(e);
        }

        protected override void OnEnabledChanged(EventArgs e) {
            Invalidate();
            base.OnEnabledChanged(e);
        }

        protected override void OnPaint(PaintEventArgs e) {
            Graphics g = e.Graphics;
            g.Clear(AppTheme.EffectiveBack(this));
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var box = new Rectangle(0, (Height - BoxSize) / 2, BoxSize - 1, BoxSize - 1);
            PaintBox(g, box, Checked, Enabled, hovering);

            Color fore = Enabled ? AppTheme.Text : Color.FromArgb(130, AppTheme.SubtleText);
            var textBounds = new Rectangle(BoxSize + TextGap, 0, Math.Max(0, Width - BoxSize - TextGap), Height);
            TextRenderer.DrawText(g, Text, Font, textBounds, fore,
                TextFormatFlags.NoPrefix | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis);

            if (Focused && ShowFocusCues) {
                ControlPaint.DrawFocusRectangle(g, new Rectangle(BoxSize + TextGap - 2, 1, Math.Max(0, Width - BoxSize - TextGap + 1), Height - 2), fore, AppTheme.EffectiveBack(this));
            }
        }
    }

    /// <summary>ComboBox mit dunklem Rahmen, Pfeil und Auswahlliste.</summary>
    public class ThemedComboBox : ComboBox, IThemedControl {
        private const int WM_PAINT = 0x000F;
        private const int WM_MOUSEWHEEL = 0x020A;

        public ThemedComboBox() {
            DrawMode = DrawMode.OwnerDrawFixed;
            FlatStyle = FlatStyle.Flat;
            ItemHeight = 20;
            ApplyTheme();
        }

        public void ApplyTheme() {
            BackColor = AppTheme.SurfaceAlt;
            ForeColor = AppTheme.Text;
            Invalidate();
        }

        protected override void OnHandleCreated(EventArgs e) {
            base.OnHandleCreated(e);
            AppTheme.ApplyNativeTheme(this);
        }

        protected override void OnDrawItem(DrawItemEventArgs e) {
            bool isEditPart = (e.State & DrawItemState.ComboBoxEdit) != 0;
            bool selected = !isEditPart && (e.State & DrawItemState.Selected) != 0;

            using (var brush = new SolidBrush(selected ? AppTheme.NavSelected : AppTheme.SurfaceAlt)) {
                e.Graphics.FillRectangle(brush, e.Bounds);
            }

            if (e.Index >= 0) {
                var textBounds = new Rectangle(e.Bounds.X + 6, e.Bounds.Y, Math.Max(0, e.Bounds.Width - 6), e.Bounds.Height);
                Color fore = Enabled ? AppTheme.Text : AppTheme.SubtleText;
                TextRenderer.DrawText(e.Graphics, GetItemText(Items[e.Index]), Font, textBounds, fore,
                    TextFormatFlags.NoPrefix | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis);
            }
        }

        protected override void OnGotFocus(EventArgs e) {
            base.OnGotFocus(e);
            Invalidate();
        }

        protected override void OnLostFocus(EventArgs e) {
            base.OnLostFocus(e);
            Invalidate();
        }

        protected override void WndProc(ref Message m) {
            // Beim Scrollen der Seite darf ein zufällig überfahrenes Feld seinen Wert nicht ändern.
            if (m.Msg == WM_MOUSEWHEEL && !Focused && !DroppedDown && Parent != null) {
                NativeMethods.SendMessage(Parent.Handle, m.Msg, m.WParam, m.LParam);
                return;
            }

            base.WndProc(ref m);
            if (m.Msg == WM_PAINT && IsHandleCreated && Width > 0 && Height > 0) {
                PaintChrome();
            }
        }

        // Zeichnet Rahmen und Pfeil über die nativen (hellen) Anteile der ComboBox.
        private void PaintChrome() {
            using (Graphics g = Graphics.FromHwnd(Handle)) {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                const int arrowWidth = 24;
                var arrowZone = new Rectangle(Width - arrowWidth, 1, arrowWidth - 1, Height - 2);
                using (var brush = new SolidBrush(AppTheme.SurfaceAlt)) {
                    g.FillRectangle(brush, arrowZone);
                }

                using (var pen = new Pen(AppTheme.SurfaceAlt, 2f)) {
                    g.DrawRectangle(pen, 1, 1, Width - 3, Height - 3);
                }

                Color arrowColor = Enabled ? AppTheme.Text : AppTheme.SubtleText;
                float arrowTop = (Height - 16) / 2f;
                Icons.Draw(g, IconKind.ChevronDown, new RectangleF(Width - arrowWidth + 3, arrowTop, 16f, 16f), arrowColor, 1.6f);

                using (var outline = AppTheme.RoundedPath(new Rectangle(0, 0, Width - 1, Height - 1), 3))
                using (var outside = new Region(new Rectangle(0, 0, Width, Height)))
                using (var parentBrush = new SolidBrush(AppTheme.EffectiveBack(this)))
                using (var pen = new Pen(Focused ? AppTheme.Accent : AppTheme.InputBorder)) {
                    outside.Exclude(outline);
                    g.FillRegion(parentBrush, outside);
                    g.DrawPath(pen, outline);
                }
            }
        }
    }

    /// <summary>TextBox mit themenkonformem 1-px-Rahmen und dunklen Scrollbars.</summary>
    public class ThemedTextBox : TextBox, IThemedControl {
        private const int WM_NCPAINT = 0x0085;

        public ThemedTextBox() {
            BorderStyle = BorderStyle.FixedSingle;
            ApplyTheme();
        }

        public void ApplyTheme() {
            bool flat = BorderStyle == BorderStyle.None;
            BackColor = flat ? AppTheme.Surface : AppTheme.SurfaceAlt;
            ForeColor = AppTheme.Text;
            RedrawFrame();
        }

        protected override void OnHandleCreated(EventArgs e) {
            base.OnHandleCreated(e);
            AppTheme.ApplyNativeTheme(this);
        }

        protected override void OnGotFocus(EventArgs e) {
            base.OnGotFocus(e);
            RedrawFrame();
        }

        protected override void OnLostFocus(EventArgs e) {
            base.OnLostFocus(e);
            RedrawFrame();
        }

        protected override void WndProc(ref Message m) {
            base.WndProc(ref m);
            if (m.Msg == WM_NCPAINT && BorderStyle == BorderStyle.FixedSingle && IsHandleCreated) {
                PaintBorder();
            }
        }

        private void RedrawFrame() {
            if (IsHandleCreated) {
                NativeMethods.RedrawWindow(Handle, IntPtr.Zero, IntPtr.Zero, NativeMethods.RDW_FRAME | NativeMethods.RDW_INVALIDATE);
            }
        }

        private void PaintBorder() {
            IntPtr hdc = NativeMethods.GetWindowDC(Handle);
            if (hdc == IntPtr.Zero) {
                return;
            }
            try {
                using (Graphics g = Graphics.FromHdc(hdc))
                using (var pen = new Pen(Focused ? AppTheme.Accent : AppTheme.InputBorder)) {
                    g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
                }
            } finally {
                NativeMethods.ReleaseDC(Handle, hdc);
            }
        }
    }

    /// <summary>Zahlenfeld mit gezeichneten Pfeilen und optional zweistelliger Anzeige.</summary>
    public class ThemedNumericUpDown : NumericUpDown, IThemedControl {
        private const int WM_MOUSEWHEEL = 0x020A;
        private bool twoDigits;

        public ThemedNumericUpDown() {
            BorderStyle = BorderStyle.FixedSingle;
            foreach (Control child in Controls) {
                if (!(child is TextBox)) {
                    child.Paint += Spinner_Paint;
                    child.Cursor = Cursors.Hand;
                }
            }
            ApplyTheme();
        }

        [DefaultValue(false)]
        public bool TwoDigits {
            get { return twoDigits; }
            set { twoDigits = value; UpdateEditText(); }
        }

        public void ApplyTheme() {
            BackColor = AppTheme.SurfaceAlt;
            ForeColor = AppTheme.Text;
            Invalidate(true);
        }

        protected override void WndProc(ref Message m) {
            // Beim Scrollen der Seite darf ein zufällig überfahrenes Feld seinen Wert nicht ändern.
            if (m.Msg == WM_MOUSEWHEEL && !ContainsFocus && Parent != null) {
                NativeMethods.SendMessage(Parent.Handle, m.Msg, m.WParam, m.LParam);
                return;
            }
            base.WndProc(ref m);
        }

        protected override void UpdateEditText() {
            if (twoDigits) {
                Text = ((int)Value).ToString("00");
            } else {
                base.UpdateEditText();
            }
        }

        protected override void OnEnabledChanged(EventArgs e) {
            base.OnEnabledChanged(e);
            Invalidate(true);
        }

        protected override void OnGotFocus(EventArgs e) {
            base.OnGotFocus(e);
            Invalidate();
        }

        protected override void OnLostFocus(EventArgs e) {
            base.OnLostFocus(e);
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e) {
            // Bewusst ohne base.OnPaint: die Basisklasse zeichnet einen hellen System-Rahmen.
            e.Graphics.Clear(AppTheme.SurfaceAlt);
            using (var pen = new Pen(ContainsFocus ? AppTheme.Accent : AppTheme.InputBorder)) {
                e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
            }
        }

        private void Spinner_Paint(object sender, PaintEventArgs e) {
            var spinner = (Control)sender;
            e.Graphics.Clear(AppTheme.SurfaceAlt);
            Color color = Enabled ? AppTheme.Text : AppTheme.SubtleText;
            float half = spinner.Height / 2f;
            float size = Math.Min(spinner.Width, half + 2f);
            float x = (spinner.Width - size) / 2f;
            Icons.Draw(e.Graphics, IconKind.ChevronUp, new RectangleF(x, (half - size) / 2f + 1f, size, size), color, 1.4f);
            Icons.Draw(e.Graphics, IconKind.ChevronDown, new RectangleF(x, half + (half - size) / 2f - 1f, size, size), color, 1.4f);
        }
    }

    /// <summary>Uhrzeit-Eingabe (Stunde : Minute); die native DateTimePicker-Darstellung lässt sich nicht einfärben.</summary>
    public class ThemedTimePicker : Panel, IThemedControl {
        private readonly ThemedNumericUpDown hours;
        private readonly ThemedNumericUpDown minutes;
        private readonly Label separator;

        public ThemedTimePicker() {
            hours = new ThemedNumericUpDown { Minimum = 0, Maximum = 23, TwoDigits = true, Location = new Point(0, 0), Width = 54 };
            separator = new Label { Text = ":", AutoSize = false, TextAlign = ContentAlignment.MiddleCenter, Location = new Point(54, 0), Size = new Size(12, 23), BackColor = Color.Transparent };
            minutes = new ThemedNumericUpDown { Minimum = 0, Maximum = 59, TwoDigits = true, Location = new Point(66, 0), Width = 54 };
            Controls.Add(hours);
            Controls.Add(separator);
            Controls.Add(minutes);
            Size = new Size(120, 23);
            BackColor = Color.Transparent;
            separator.ForeColor = AppTheme.Text;
            hours.ValueChanged += (s, e) => ValueChanged?.Invoke(this, EventArgs.Empty);
            minutes.ValueChanged += (s, e) => ValueChanged?.Invoke(this, EventArgs.Empty);
        }

        public event EventHandler ValueChanged;

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public DateTime Value {
            get { return DateTime.Today.AddHours((double)hours.Value).AddMinutes((double)minutes.Value); }
            set {
                hours.Value = value.Hour;
                minutes.Value = value.Minute;
            }
        }

        public void ApplyTheme() {
            hours.ApplyTheme();
            minutes.ApplyTheme();
            separator.ForeColor = AppTheme.Text;
        }
    }
}
