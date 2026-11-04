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

    /// <summary>Bereich mit optionalem Titel und Trennlinie – Ersatz für GroupBoxen.</summary>
    public class CardPanel : Panel, IThemedControl {
        private string title = string.Empty;

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

        public void ApplyTheme() {
            BackColor = AppTheme.Surface;
        }

        protected override void OnPaintBackground(PaintEventArgs e) {
            // Bereiche gliedern sich über Abstand und eine Trennlinie, nicht über Rahmen.
            e.Graphics.Clear(AppTheme.Surface);
            if (title.Length > 0) {
                using (var pen = new Pen(AppTheme.Separator)) {
                    e.Graphics.DrawLine(pen, 20, 0, Width - 17, 0);
                }
            }
        }

        protected override void OnPaint(PaintEventArgs e) {
            base.OnPaint(e);
            if (title.Length == 0) {
                return;
            }

            // Transparente Kind-Labels lassen die Karte verschoben neu zeichnen – die
            // Transformation muss erhalten bleiben, sonst erscheint der Titel in ihnen.
            TextRenderer.DrawText(e.Graphics, title, AppTheme.CardTitleFont, new Point(20, 12), AppTheme.Text,
                TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding | TextFormatFlags.PreserveGraphicsTranslateTransform | TextFormatFlags.PreserveGraphicsClipping);
        }
    }

    /// <summary>Kennzahl aus Beschriftung und Wert.</summary>
    public class StatCard : Control, IThemedControl {
        private string caption = string.Empty;
        private string value = "0";
        private bool leadingRule;

        public StatCard() {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }

        /// <summary>Trennlinie links, für alle Kennzahlen außer der ersten.</summary>
        [DefaultValue(false)]
        public bool LeadingRule {
            get { return leadingRule; }
            set { leadingRule = value; Invalidate(); }
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
            int left = 0;
            if (leadingRule) {
                using (var pen = new Pen(AppTheme.Separator)) {
                    g.DrawLine(pen, 0, 2, 0, Height - 3);
                }
                left = 16;
            }

            const TextFormatFlags flags = TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine;
            int width = Math.Max(0, Width - left);
            TextRenderer.DrawText(g, caption, AppTheme.StatCaptionFont, new Rectangle(left, 0, width, 18), AppTheme.SubtleText, flags);
            TextRenderer.DrawText(g, value, AppTheme.StatValueFont, new Rectangle(left, 17, width, 28), AppTheme.Text, flags);
        }
    }

    /// <summary>Flacher Fortschrittsbalken; API kompatibel zur Standard-ProgressBar.</summary>
    public class NeonProgressBar : Control, IThemedControl {
        private int maximum = 100;
        private int value;
        private bool marquee;
        private float marqueePosition;
        private Timer marqueeTimer;

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

        /// <summary>Zurückhaltender Grauton für den untergeordneten Balken.</summary>
        [DefaultValue(false)]
        public bool Muted { get; set; }

        /// <summary>Laufbalken für Schritte ohne messbaren Fortschritt. Nur vom UI-Thread setzen.</summary>
        [DefaultValue(false)]
        public bool Marquee {
            get { return marquee; }
            set {
                if (marquee == value) {
                    return;
                }
                marquee = value;
                marqueePosition = 0f;
                if (marquee) {
                    if (marqueeTimer == null) {
                        marqueeTimer = new Timer { Interval = 33 };
                        marqueeTimer.Tick += (s, e) => {
                            marqueePosition = marqueePosition >= 1f ? 0f : marqueePosition + 0.012f;
                            Invalidate();
                        };
                    }
                    marqueeTimer.Start();
                } else if (marqueeTimer != null) {
                    marqueeTimer.Stop();
                }
                Invalidate();
            }
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

            var track = new Rectangle(0, 0, Width - 1, Height - 1);
            if (track.Width <= 0 || track.Height <= 0) {
                return;
            }

            int radius = Math.Min(track.Height / 2, 2);
            using (var path = AppTheme.RoundedPath(track, radius))
            using (var brush = new SolidBrush(AppTheme.Track)) {
                g.FillPath(brush, path);
            }

            if (marquee) {
                int segment = Math.Max(track.Width / 5, 40);
                int start = track.X - segment + (int)((track.Width + segment) * marqueePosition);
                Rectangle moving = Rectangle.Intersect(track, new Rectangle(start, track.Y, segment, track.Height));
                if (moving.Width > 0) {
                    using (var path = AppTheme.RoundedPath(moving, radius))
                    using (var brush = new SolidBrush(AppTheme.Accent)) {
                        g.FillPath(brush, path);
                    }
                }
                return;
            }

            int fillWidth = (int)Math.Round(track.Width * Fraction);
            if (fillWidth <= 0) {
                return;
            }

            var fill = new Rectangle(track.X, track.Y, Math.Max(fillWidth, track.Height), track.Height);
            using (var path = AppTheme.RoundedPath(fill, radius))
            using (var brush = new SolidBrush(Muted ? AppTheme.SubtleText : AppTheme.Accent)) {
                g.FillPath(brush, path);
            }
        }

        protected override void Dispose(bool disposing) {
            if (disposing && marqueeTimer != null) {
                marqueeTimer.Dispose();
                marqueeTimer = null;
            }
            base.Dispose(disposing);
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
