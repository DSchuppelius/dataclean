using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace DataClean.UI {
    /// <summary>Controls, die ihre Farben selbst aus der <see cref="AppTheme"/>-Palette beziehen.</summary>
    public interface IThemedControl {
        void ApplyTheme();
    }

    /// <summary>
    /// Panel, das Mausabfragen an den Fensterrändern an das randlose Fenster weiterreicht,
    /// damit die Resize-Griffe auch dort greifen, wo Inhalte bis an den Rand reichen.
    /// </summary>
    public class EdgeTransparentPanel : Panel {
        private const int WM_NCHITTEST = 0x84;
        private const int HTTRANSPARENT = -1;

        /// <summary>Reicht die Abfrage weiter, wenn der Mauszeiger auf einem Resize-Griff liegt.</summary>
        internal static bool HandleEdgeHitTest(Control control, ref Message m) {
            if (m.Msg == WM_NCHITTEST && control.FindForm() is FuturisticForm form && form.GetResizeHit(m.LParam) != 0) {
                m.Result = (IntPtr)HTTRANSPARENT;
                return true;
            }
            return false;
        }

        protected override void WndProc(ref Message m) {
            if (!HandleEdgeHitTest(this, ref m)) {
                base.WndProc(ref m);
            }
        }
    }

    /// <summary>Vollflächiger Seitenhintergrund im Inhaltsbereich.</summary>
    public class PagePanel : EdgeTransparentPanel, IThemedControl {
        public PagePanel() {
            DoubleBuffered = true;
            ApplyTheme();
        }

        public void ApplyTheme() {
            BackColor = AppTheme.Background;
        }
    }

    /// <summary>Scrollbarer Bereich mit dunklen Scrollleisten.</summary>
    public class ScrollHostPanel : EdgeTransparentPanel, IThemedControl {
        public ScrollHostPanel() {
            DoubleBuffered = true;
            ApplyTheme();
        }

        public void ApplyTheme() {
            BackColor = AppTheme.Background;
            if (IsHandleCreated) {
                AppTheme.ApplyNativeTheme(this);
            }
        }

        protected override void OnHandleCreated(EventArgs e) {
            base.OnHandleCreated(e);
            AppTheme.ApplyNativeTheme(this);
        }
    }

    /// <summary>Abgerundete Karte mit optionalem Titel und Icon – Ersatz für GroupBoxen.</summary>
    public class CardPanel : Panel, IThemedControl {
        private const int Radius = 7;
        private string title = string.Empty;
        private IconKind titleIcon = IconKind.None;

        public CardPanel() {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            ApplyTheme();
        }

        [DefaultValue("")]
        public string Title {
            get { return title; }
            set { title = value ?? string.Empty; Invalidate(); }
        }

        [DefaultValue(IconKind.None)]
        public IconKind TitleIcon {
            get { return titleIcon; }
            set { titleIcon = value; Invalidate(); }
        }

        public void ApplyTheme() {
            BackColor = AppTheme.Surface;
        }

        protected override void OnPaintBackground(PaintEventArgs e) {
            Graphics g = e.Graphics;
            g.Clear(AppTheme.EffectiveBack(this));
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = AppTheme.RoundedPath(new Rectangle(0, 0, Width - 1, Height - 1), Radius))
            using (var fill = new SolidBrush(AppTheme.Surface))
            using (var pen = new Pen(AppTheme.CardBorder)) {
                g.FillPath(fill, path);
                g.DrawPath(pen, path);
            }
        }

        protected override void OnPaint(PaintEventArgs e) {
            base.OnPaint(e);
            if (title.Length == 0) {
                return;
            }

            int x = 20;
            if (titleIcon != IconKind.None) {
                Icons.Draw(e.Graphics, titleIcon, new RectangleF(19f, 13f, 22f, 22f), AppTheme.Text);
                x = 48;
            }
            // Transparente Kind-Labels lassen die Karte verschoben neu zeichnen – die
            // Transformation muss erhalten bleiben, sonst erscheint der Titel in ihnen.
            TextRenderer.DrawText(e.Graphics, title, AppTheme.CardTitleFont, new Point(x, 13), AppTheme.Text,
                TextFormatFlags.NoPrefix | TextFormatFlags.PreserveGraphicsTranslateTransform | TextFormatFlags.PreserveGraphicsClipping);
        }
    }

    /// <summary>Kennzahl-Kachel (Icon, Beschriftung, großer Wert).</summary>
    public class StatCard : Control, IThemedControl {
        private IconKind icon = IconKind.File;
        private Tone tone = Tone.Accent;
        private string caption = string.Empty;
        private string value = "0";

        public StatCard() {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }

        [DefaultValue(IconKind.File)]
        public IconKind Icon {
            get { return icon; }
            set { icon = value; Invalidate(); }
        }

        [DefaultValue(Tone.Accent)]
        public Tone Tone {
            get { return tone; }
            set { tone = value; Invalidate(); }
        }

        [DefaultValue("")]
        public string Caption {
            get { return caption; }
            set { caption = value ?? string.Empty; Invalidate(); }
        }

        [DefaultValue("0")]
        public string Value {
            get { return value; }
            set {
                string next = value ?? string.Empty;
                if (next != this.value) {
                    this.value = next;
                    Invalidate();
                }
            }
        }

        public void ApplyTheme() {
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e) {
            Graphics g = e.Graphics;
            g.Clear(AppTheme.EffectiveBack(this));
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = AppTheme.RoundedPath(new Rectangle(0, 0, Width - 1, Height - 1), 7))
            using (var fill = new SolidBrush(AppTheme.Surface))
            using (var pen = new Pen(AppTheme.CardBorder)) {
                g.FillPath(fill, path);
                g.DrawPath(pen, path);
            }

            const int iconSize = 32;
            Icons.DrawGlowing(g, icon, new RectangleF(20f, (Height - iconSize) / 2f, iconSize, iconSize),
                AppTheme.ToneColor(tone), 1.9f);

            const TextFormatFlags flags = TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine;
            const int textLeft = 72;
            int textWidth = Math.Max(0, Width - textLeft - 10);
            TextRenderer.DrawText(g, caption, AppTheme.StatCaptionFont, new Rectangle(textLeft, 11, textWidth, 20), AppTheme.SubtleText, flags);
            TextRenderer.DrawText(g, value, AppTheme.StatValueFont, new Rectangle(textLeft, 29, textWidth, 34), AppTheme.Text, flags);
        }
    }

    /// <summary>Leuchtender Fortschrittsbalken; API kompatibel zur Standard-ProgressBar.</summary>
    public class NeonProgressBar : Control, IThemedControl {
        private int maximum = 100;
        private int value;
        private int glowSize = 4;

        public NeonProgressBar() {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            TabStop = false;
        }

        [DefaultValue(100)]
        public int Maximum {
            get { return maximum; }
            set {
                maximum = Math.Max(0, value);
                this.value = Math.Min(this.value, maximum);
                Invalidate();
            }
        }

        [DefaultValue(0)]
        public int Value {
            get { return value; }
            set {
                int next = Math.Max(0, Math.Min(value, maximum));
                if (next != this.value) {
                    this.value = next;
                    Invalidate();
                }
            }
        }

        /// <summary>Breite des Leuchtsaums; das Control ist um diesen Rand höher als der Balken.</summary>
        [DefaultValue(4)]
        public int GlowSize {
            get { return glowSize; }
            set { glowSize = Math.Max(0, value); Invalidate(); }
        }

        /// <summary>Anteil von 0.0 bis 1.0.</summary>
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public double Fraction {
            get { return maximum > 0 ? (double)value / maximum : 0d; }
        }

        public void ApplyTheme() {
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e) {
            Graphics g = e.Graphics;
            g.Clear(AppTheme.EffectiveBack(this));
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var track = new Rectangle(glowSize, glowSize, Width - 2 * glowSize - 1, Height - 2 * glowSize - 1);
            if (track.Width <= 0 || track.Height <= 0) {
                return;
            }

            int radius = Math.Min(track.Height / 2, 4);
            using (var path = AppTheme.RoundedPath(track, radius))
            using (var brush = new SolidBrush(AppTheme.Track)) {
                g.FillPath(brush, path);
            }

            int fillWidth = (int)Math.Round(track.Width * Fraction);
            if (fillWidth <= 0) {
                return;
            }

            var fill = new Rectangle(track.X, track.Y, Math.Max(fillWidth, track.Height), track.Height);
            for (int i = glowSize; i >= 1; i--) {
                var halo = Rectangle.Inflate(fill, i, i);
                int alpha = AppTheme.IsDark ? 44 / (i + 1) : 22 / (i + 1);
                using (var path = AppTheme.RoundedPath(halo, radius + i))
                using (var brush = new SolidBrush(Color.FromArgb(alpha, AppTheme.Accent))) {
                    g.FillPath(brush, path);
                }
            }

            using (var path = AppTheme.RoundedPath(fill, radius))
            using (var brush = new LinearGradientBrush(new Rectangle(fill.X, fill.Y - 1, fill.Width, fill.Height + 2),
                Lighten(AppTheme.Accent, 0.28f), AppTheme.Accent, LinearGradientMode.Vertical)) {
                g.FillPath(brush, path);
            }
        }

        private static Color Lighten(Color c, float amount) {
            return Color.FromArgb(
                c.R + (int)((255 - c.R) * amount),
                c.G + (int)((255 - c.G) * amount),
                c.B + (int)((255 - c.B) * amount));
        }
    }

    /// <summary>Einzeilige Hinweiszeile mit Icon.</summary>
    public class IconLabel : Control, IThemedControl {
        private IconKind icon = IconKind.Info;
        private Tone tone = Tone.Warning;

        public IconLabel() {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            TabStop = false;
        }

        [DefaultValue(IconKind.Info)]
        public IconKind Icon {
            get { return icon; }
            set { icon = value; Invalidate(); }
        }

        [DefaultValue(Tone.Warning)]
        public Tone Tone {
            get { return tone; }
            set { tone = value; Invalidate(); }
        }

        public void ApplyTheme() {
            Invalidate();
        }

        protected override void OnTextChanged(EventArgs e) {
            base.OnTextChanged(e);
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e) {
            Graphics g = e.Graphics;
            g.Clear(AppTheme.EffectiveBack(this));
            Color color = AppTheme.ToneColor(tone);
            const int iconSize = 18;
            Icons.Draw(g, icon, new RectangleF(0f, (Height - iconSize) / 2f, iconSize, iconSize), color);
            TextRenderer.DrawText(g, Text, Font, new Rectangle(iconSize + 8, 0, Math.Max(0, Width - iconSize - 8), Height), color,
                TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine | TextFormatFlags.VerticalCenter);
        }
    }

    /// <summary>Kontextmenü im App-Design (Infobereich).</summary>
    public class ThemedMenuRenderer : ToolStripProfessionalRenderer {
        public ThemedMenuRenderer() : base(new MenuColors()) {
            RoundedEdges = false;
        }

        protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e) {
            e.TextColor = e.Item.Enabled ? AppTheme.Text : AppTheme.SubtleText;
            base.OnRenderItemText(e);
        }

        private sealed class MenuColors : ProfessionalColorTable {
            public override Color ToolStripDropDownBackground { get { return AppTheme.Surface; } }
            public override Color ImageMarginGradientBegin { get { return AppTheme.Surface; } }
            public override Color ImageMarginGradientMiddle { get { return AppTheme.Surface; } }
            public override Color ImageMarginGradientEnd { get { return AppTheme.Surface; } }
            public override Color MenuBorder { get { return AppTheme.CardBorder; } }
            public override Color MenuItemBorder { get { return AppTheme.NavSelectedBorder; } }
            public override Color MenuItemSelected { get { return AppTheme.NavSelected; } }
            public override Color SeparatorDark { get { return AppTheme.Separator; } }
            public override Color SeparatorLight { get { return AppTheme.Surface; } }
        }
    }
}
