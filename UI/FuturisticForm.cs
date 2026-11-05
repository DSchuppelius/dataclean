using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Microsoft.Win32;

namespace DataClean.UI {
    /// <summary>
    /// Basis-Form für das Dashboard-Design: randloses Fenster mit eigener Titelleiste
    /// (Marke, Untertitel, Fenster-Buttons), optionaler Navigations-Sidebar und Fußzeile.
    /// Die im Designer angelegten Controls landen zur Laufzeit im Inhaltsbereich.
    ///
    /// Reines WinForms/GDI+ – die abgerundeten Fensterecken unter Windows 11 sind optional.
    /// </summary>
    public class FuturisticForm : Form {
        private const int CaptionHeight = 36;
        private const int BorderThickness = 1;
        private const int ResizeGrip = 6;
        private const int SidebarWidth = 176;
        private const int NavTop = 26;
        private const int NavSpacing = 40;
        private const int SidebarLogoSize = 72;
        private const int SidebarLogoMargin = 18;

        private const int WM_NCHITTEST = 0x84;
        private const int HTLEFT = 10, HTRIGHT = 11, HTTOP = 12, HTTOPLEFT = 13,
            HTTOPRIGHT = 14, HTBOTTOM = 15, HTBOTTOMLEFT = 16, HTBOTTOMRIGHT = 17;
        private const int CS_DROPSHADOW = 0x00020000;
        private const int WS_MINIMIZEBOX = 0x00020000;
        private const int WS_MAXIMIZEBOX = 0x00010000;
        private const int WS_SYSMENU = 0x00080000;

        private EdgeTransparentPanel captionBar;
        private CaptionButton closeButton;
        private CaptionButton maxButton;
        private CaptionButton minButton;
        private EdgeTransparentPanel contentHost;
        private EdgeTransparentPanel sidebarPanel;
        private readonly List<NavItem> navItems = new List<NavItem>();
        private string captionSubtitle;
        private bool chromeBuilt;
        private bool appliedDark;

        public FuturisticForm() {
            AppTheme.Refresh();
            FormBorderStyle = FormBorderStyle.None;
            DoubleBuffered = true;
            BackColor = AppTheme.Background;
        }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool EnableMinimizeButton { get; set; } = true;

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool EnableMaximizeButton { get; set; } = false;

        /// <summary>Erlaubt das Ändern der Fenstergröße über die Ränder.</summary>
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool AllowResize { get; set; } = false;

        /// <summary>Beschränkt das Ändern der Fenstergröße auf die Breite (z. B. für kompakte Ansichten).</summary>
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool LockVerticalResize { get; set; } = false;

        /// <summary>Aktiviert die linke Navigations-Sidebar. Vor dem Laden setzen.</summary>
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool EnableSidebar { get; set; } = false;

        /// <summary>Fußzeile, die über die gesamte Fensterbreite läuft. Vor dem Laden setzen.</summary>
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Control FooterControl { get; set; }

        /// <summary>Bild unten in der Sidebar, z. B. das Firmenlogo. Vor dem Laden setzen.</summary>
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Image SidebarLogo { get; set; }

        /// <summary>Text rechts neben der Marke in der Titelleiste.</summary>
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string CaptionSubtitle {
            get { return captionSubtitle; }
            set {
                captionSubtitle = value;
                captionBar?.Invalidate();
            }
        }

        /// <summary>Inhaltsbereich, in den die Designer-Controls zur Laufzeit umgehängt werden.</summary>
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Panel ContentHost {
            get { return contentHost; }
        }

        /// <summary>Kleinste Fenstergröße bei <see cref="AllowResize"/>; leer = Designgröße.</summary>
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Size ResizeMinimum { get; set; } = Size.Empty;

        protected override CreateParams CreateParams {
            get {
                CreateParams cp = base.CreateParams;
                if (!DesignMode) {
                    cp.ClassStyle |= CS_DROPSHADOW;
                    cp.Style |= WS_SYSMENU;
                    if (EnableMinimizeButton) {
                        cp.Style |= WS_MINIMIZEBOX;
                    }
                    if (EnableMaximizeButton) {
                        cp.Style |= WS_MAXIMIZEBOX;
                    }
                }
                return cp;
            }
        }

        protected override void OnHandleCreated(EventArgs e) {
            base.OnHandleCreated(e);
            SetStyle(ControlStyles.ResizeRedraw, true);
            ApplyWindowFrame();
            if (!DesignMode) {
                SystemEvents.UserPreferenceChanged += SystemPreferenceChanged;
            }
        }

        protected override void OnHandleDestroyed(EventArgs e) {
            SystemEvents.UserPreferenceChanged -= SystemPreferenceChanged;
            base.OnHandleDestroyed(e);
        }

        // Steht die Einstellung auf "System", folgt das Fenster einem Wechsel in Windows sofort.
        // Das Ereignis kommt auch bei vielen anderen Änderungen, daher der Vergleich.
        private void SystemPreferenceChanged(object sender, UserPreferenceChangedEventArgs e) {
            if (e.Category == UserPreferenceCategory.General && chromeBuilt && AppTheme.ResolveDark() != appliedDark) {
                ReapplyTheme();
            }
        }

        protected override void OnLoad(EventArgs e) {
            // Chrome vor dem Basisaufruf aufbauen: base.OnLoad zentriert modale Fenster
            // und soll bereits die endgültige Fenstergröße verwenden.
            BuildChrome();
            base.OnLoad(e);
            if (!DesignMode) {
                ApplyThemeRecursive(this);
                UpdateChromeColors();
                if (!Modal && StartPosition == FormStartPosition.CenterScreen) {
                    CenterToScreen();
                }
            }
        }

        private void BuildChrome() {
            if (chromeBuilt || DesignMode) {
                return;
            }
            chromeBuilt = true;

            SuspendLayout();

            Size designSize = ClientSize;
            Control footer = FooterControl;
            int footerHeight = 0;
            if (footer != null) {
                Controls.Remove(footer);
                footerHeight = footer.Height;
            }

            // Vorhandene Controls in den Inhaltsbereich umhängen. Er hat zunächst die
            // Designgröße, damit Anker-Abstände der Controls unverändert bleiben.
            contentHost = new EdgeTransparentPanel {
                Bounds = new Rectangle(0, 0, designSize.Width, designSize.Height - footerHeight),
                BackColor = AppTheme.Background
            };
            contentHost.SuspendLayout();
            var existing = new Control[Controls.Count];
            Controls.CopyTo(existing, 0);
            foreach (Control control in existing) {
                Controls.Remove(control);
                contentHost.Controls.Add(control);
            }
            contentHost.ResumeLayout(false);
            contentHost.Dock = DockStyle.Fill;

            // Andock-Reihenfolge: zuletzt hinzugefügt wird zuerst angedockt
            // (Titelleiste > Fußzeile > Sidebar > Inhalt).
            Controls.Add(contentHost);

            if (EnableSidebar) {
                BuildSidebar();
                Controls.Add(sidebarPanel);
            }

            if (footer != null) {
                footer.Dock = DockStyle.Bottom;
                footer.Paint -= Footer_Paint;
                footer.Paint += Footer_Paint;
                Controls.Add(footer);
            }

            BuildCaptionBar();
            Controls.Add(captionBar);

            Padding = new Padding(BorderThickness);
            int width = designSize.Width + (EnableSidebar ? SidebarWidth : 0) + 2 * BorderThickness;
            int height = designSize.Height + CaptionHeight + 2 * BorderThickness;
            ClientSize = new Size(width, height);

            if (AllowResize) {
                Size minimum = ResizeMinimum.IsEmpty ? Size : ResizeMinimum;
                Rectangle workingArea = Screen.FromControl(this).WorkingArea;
                MinimumSize = new Size(Math.Min(minimum.Width, workingArea.Width), Math.Min(minimum.Height, workingArea.Height));
                Size = new Size(Math.Min(Width, workingArea.Width), Math.Min(Height, workingArea.Height));
            } else {
                // Nicht skalierbare Dialoge scrollen, falls der Bildschirm zu klein ist.
                contentHost.AutoScroll = true;
                Rectangle workingArea = Screen.FromControl(this).WorkingArea;
                if (Height > workingArea.Height) {
                    Height = workingArea.Height;
                }
            }

            ResumeLayout(true);
        }

        private void BuildCaptionBar() {
            captionBar = new EdgeTransparentPanel {
                Dock = DockStyle.Top,
                Height = CaptionHeight,
                BackColor = AppTheme.Caption
            };
            typeof(Panel).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                ?.SetValue(captionBar, true, null);
            captionBar.MouseDown += Caption_MouseDown;
            captionBar.Paint += Caption_Paint;

            // Zuletzt hinzugefügt dockt zuerst an: Schließen steht damit ganz rechts.
            if (EnableMinimizeButton) {
                minButton = new CaptionButton(IconKind.Minimize, false) { Dock = DockStyle.Right };
                minButton.Click += (s, e) => WindowState = FormWindowState.Minimized;
                captionBar.Controls.Add(minButton);
            }

            if (EnableMaximizeButton) {
                maxButton = new CaptionButton(IconKind.Maximize, false) { Dock = DockStyle.Right };
                maxButton.Click += (s, e) => ToggleMaximize();
                captionBar.Controls.Add(maxButton);
            }

            closeButton = new CaptionButton(IconKind.Close, true) { Dock = DockStyle.Right };
            closeButton.Click += (s, e) => Close();
            captionBar.Controls.Add(closeButton);
        }

        private void BuildSidebar() {
            sidebarPanel = new EdgeTransparentPanel {
                Dock = DockStyle.Left,
                Width = SidebarWidth,
                BackColor = AppTheme.Sidebar
            };
            sidebarPanel.Paint += (s, e) => {
                using (var pen = new Pen(AppTheme.Separator)) {
                    e.Graphics.DrawLine(pen, sidebarPanel.Width - 1, 0, sidebarPanel.Width - 1, sidebarPanel.Height);
                }
                PaintSidebarLogo(e.Graphics);
            };
            // Das Logo hängt am unteren Rand: Ändert sich die Höhe, muss die ganze Fläche neu gezeichnet werden.
            sidebarPanel.Resize += (s, e) => sidebarPanel.Invalidate();
        }

        // Das Logo steht als Marke über der Navigation; ohne Logo rückt die Navigation nach oben.
        private int NavFirstTop {
            get { return SidebarLogo != null ? SidebarLogoMargin + SidebarLogoSize + SidebarLogoMargin : NavTop; }
        }

        private void PaintSidebarLogo(Graphics g) {
            if (SidebarLogo == null) {
                return;
            }

            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.DrawImage(SidebarLogo, new Rectangle((sidebarPanel.Width - SidebarLogoSize) / 2, SidebarLogoMargin, SidebarLogoSize, SidebarLogoSize));
        }

        /// <summary>Fügt einen Navigationseintrag hinzu. Erst nach dem Laden aufrufen.</summary>
        public NavItem AddNavItem(IconKind icon, string text, EventHandler onClick) {
            if (sidebarPanel == null) {
                return null;
            }

            var item = new NavItem(icon, text) {
                Location = new Point(0, NavFirstTop + navItems.Count * NavSpacing),
                Size = new Size(SidebarWidth - 1, NavSpacing - 2)
            };
            if (onClick != null) {
                item.Click += onClick;
            }
            sidebarPanel.Controls.Add(item);
            navItems.Add(item);
            return item;
        }

        /// <summary>Hebt einen Navigationseintrag als aktiv hervor.</summary>
        public void SelectNavItem(NavItem selected) {
            foreach (NavItem item in navItems) {
                item.Selected = item == selected;
            }
        }

        private void Caption_Paint(object sender, PaintEventArgs e) {
            Graphics g = e.Graphics;
            using (var pen = new Pen(AppTheme.Separator)) {
                g.DrawLine(pen, 0, captionBar.Height - 1, captionBar.Width, captionBar.Height - 1);
            }

            int brandLeft = 14;
            // Mit Sidebar steht das Logo direkt darunter; zweimal untereinander wäre es nur Dekoration.
            if (AppTheme.BrandImage != null && !EnableSidebar) {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.DrawImage(AppTheme.BrandImage, new Rectangle(12, 8, 20, 20));
                brandLeft = 40;
            }

            const TextFormatFlags flags = TextFormatFlags.NoPrefix | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine;
            int barHeight = captionBar.Height - 1;
            Size brandSize = TextRenderer.MeasureText(g, "DataClean", AppTheme.BrandFont, Size.Empty, flags);
            TextRenderer.DrawText(g, "DataClean", AppTheme.BrandFont, new Rectangle(brandLeft, 0, brandSize.Width, barHeight), AppTheme.Text, flags);

            if (!string.IsNullOrEmpty(captionSubtitle)) {
                int x = brandLeft + brandSize.Width + 10;
                using (var pen = new Pen(AppTheme.Separator)) {
                    g.DrawLine(pen, x, 11, x, barHeight - 11);
                }
                int buttonsWidth = 0;
                foreach (Control control in captionBar.Controls) {
                    buttonsWidth += control.Width;
                }
                TextRenderer.DrawText(g, captionSubtitle, AppTheme.CaptionFont,
                    new Rectangle(x + 16, 0, Math.Max(0, captionBar.Width - x - 16 - buttonsWidth), barHeight),
                    AppTheme.SubtleText, flags | TextFormatFlags.EndEllipsis);
            }
        }

        private void Footer_Paint(object sender, PaintEventArgs e) {
            var footer = (Control)sender;
            using (var pen = new Pen(AppTheme.Separator)) {
                e.Graphics.DrawLine(pen, 0, 0, footer.Width, 0);
            }
        }

        private void Caption_MouseDown(object sender, MouseEventArgs e) {
            if (e.Button != MouseButtons.Left) {
                return;
            }
            if (e.Clicks == 2 && EnableMaximizeButton) {
                ToggleMaximize();
                return;
            }
            NativeMethods.ReleaseCapture();
            NativeMethods.SendMessage(Handle, NativeMethods.WM_NCLBUTTONDOWN,
                (IntPtr)NativeMethods.HT_CAPTION, IntPtr.Zero);
        }

        private void ToggleMaximize() {
            if (WindowState == FormWindowState.Maximized) {
                WindowState = FormWindowState.Normal;
                return;
            }

            // Randlose Fenster würden sonst die Taskleiste überdecken. MaximizedBounds
            // ist relativ zum jeweiligen Monitor anzugeben.
            Screen screen = Screen.FromHandle(Handle);
            MaximizedBounds = new Rectangle(
                screen.WorkingArea.X - screen.Bounds.X,
                screen.WorkingArea.Y - screen.Bounds.Y,
                screen.WorkingArea.Width,
                screen.WorkingArea.Height);
            WindowState = FormWindowState.Maximized;
        }

        private void UpdateChromeColors() {
            appliedDark = AppTheme.IsDark;
            BackColor = AppTheme.Background;
            if (captionBar != null) {
                captionBar.BackColor = AppTheme.Caption;
                captionBar.Invalidate(true);
            }
            if (contentHost != null) {
                contentHost.BackColor = AppTheme.Background;
                if (contentHost.AutoScroll) {
                    AppTheme.ApplyNativeTheme(contentHost);
                }
            }
            if (sidebarPanel != null) {
                sidebarPanel.BackColor = AppTheme.Sidebar;
                sidebarPanel.Invalidate(true);
            }
            if (FooterControl != null) {
                FooterControl.BackColor = AppTheme.Footer;
            }
            ApplyWindowFrame();
        }

        /// <summary>Wendet die aktuelle Palette erneut an, z. B. nach Änderung der Einstellungen.</summary>
        public virtual void ReapplyTheme() {
            AppTheme.Refresh();
            ApplyThemeRecursive(this);
            UpdateChromeColors();
            OnThemeReapplied();
            Invalidate(true);
        }

        /// <summary>Für formularspezifische Farben nach <see cref="ReapplyTheme"/>.</summary>
        protected virtual void OnThemeReapplied() {
        }

        /// <summary>Wendet die Palette rekursiv auf alle Kindcontrols an.</summary>
        protected void ApplyThemeRecursive(Control root) {
            foreach (Control control in root.Controls) {
                if (control == captionBar || control == sidebarPanel) {
                    continue;
                }
                if (control is IThemedControl themed) {
                    themed.ApplyTheme();
                } else {
                    StyleStandardControl(control);
                }
                if (control.HasChildren && !(control is NumericUpDown) && !(control is ThemedTimePicker)) {
                    ApplyThemeRecursive(control);
                }
            }
        }

        private static void StyleStandardControl(Control control) {
            string role = control.Tag as string;

            if (control is LinkLabel link) {
                link.LinkColor = AppTheme.Link;
                link.ActiveLinkColor = AppTheme.LinkActive;
                link.VisitedLinkColor = AppTheme.Link;
                link.LinkBehavior = LinkBehavior.AlwaysUnderline;
                link.BackColor = Color.Transparent;
            } else if (control is Label) {
                if (role == "nostyle") {
                    return;
                }
                control.ForeColor = role == "subtle" ? AppTheme.SubtleText : AppTheme.Text;
                control.BackColor = Color.Transparent;
            } else if (control is TreeView tree) {
                tree.BackColor = AppTheme.Surface;
                tree.ForeColor = AppTheme.Text;
                tree.LineColor = AppTheme.Separator;
                AppTheme.ApplyNativeTheme(tree);
            } else if (control is FlowLayoutPanel || control is TableLayoutPanel) {
                control.BackColor = Color.Transparent;
            } else if (control is Panel) {
                control.BackColor = role == "separator" ? AppTheme.CardBorder : AppTheme.Background;
            }
        }

        protected override void OnPaint(PaintEventArgs e) {
            base.OnPaint(e);
            if (WindowState == FormWindowState.Maximized) {
                return;
            }
            using (var pen = new Pen(AppTheme.WindowBorder)) {
                e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
            }
        }

        protected override void OnResize(EventArgs e) {
            base.OnResize(e);
            if (!chromeBuilt) {
                return;
            }
            Padding = WindowState == FormWindowState.Maximized ? Padding.Empty : new Padding(BorderThickness);
            maxButton?.SetIcon(WindowState == FormWindowState.Maximized ? IconKind.Restore : IconKind.Maximize);
            Invalidate();
        }

        // Unter Windows 11 rundet der Fenstermanager die Ecken kantengeglättet ab;
        // ältere Systeme behalten eckige Fenster, der Aufruf schlägt dort folgenlos fehl.
        private void ApplyWindowFrame() {
            if (DesignMode || !IsHandleCreated) {
                return;
            }
            try {
                int corner = NativeMethods.DWMWCP_ROUND;
                NativeMethods.DwmSetWindowAttribute(Handle, NativeMethods.DWMWA_WINDOW_CORNER_PREFERENCE, ref corner, sizeof(int));
                int border = ColorTranslator.ToWin32(AppTheme.WindowBorder);
                NativeMethods.DwmSetWindowAttribute(Handle, NativeMethods.DWMWA_BORDER_COLOR, ref border, sizeof(int));
            } catch (Exception ex) when (ex is DllNotFoundException || ex is EntryPointNotFoundException) {
                // Kein DWM verfügbar – eckiges Fenster ist ein vollwertiger Rückfall.
            }
        }

        /// <summary>Liefert den Hit-Test-Code des Resize-Griffs unter dem Mauszeiger oder 0.</summary>
        internal int GetResizeHit(IntPtr lParam) {
            if (!AllowResize || !chromeBuilt || WindowState != FormWindowState.Normal) {
                return 0;
            }

            // Bildschirmkoordinaten sind vorzeichenbehaftet (Monitore links/oberhalb des Hauptmonitors).
            int packed = unchecked((int)lParam.ToInt64());
            Point pos = PointToClient(new Point((short)(packed & 0xFFFF), (short)((packed >> 16) & 0xFFFF)));

            bool left = pos.X < ResizeGrip;
            bool right = pos.X >= ClientSize.Width - ResizeGrip;
            bool top = !LockVerticalResize && pos.Y < ResizeGrip;
            bool bottom = !LockVerticalResize && pos.Y >= ClientSize.Height - ResizeGrip;

            if (top && left) return HTTOPLEFT;
            if (top && right) return HTTOPRIGHT;
            if (bottom && left) return HTBOTTOMLEFT;
            if (bottom && right) return HTBOTTOMRIGHT;
            if (left) return HTLEFT;
            if (right) return HTRIGHT;
            if (top) return HTTOP;
            if (bottom) return HTBOTTOM;
            return 0;
        }

        protected override void WndProc(ref Message m) {
            if (m.Msg == WM_NCHITTEST) {
                int hit = GetResizeHit(m.LParam);
                if (hit != 0) {
                    m.Result = (IntPtr)hit;
                    return;
                }
            }
            base.WndProc(ref m);
        }

        /// <summary>Button der Titelleiste (Minimieren, Maximieren, Schließen).</summary>
        private sealed class CaptionButton : Control {
            private readonly bool isClose;
            private IconKind icon;
            private bool hovering;

            public CaptionButton(IconKind icon, bool isClose) {
                this.icon = icon;
                this.isClose = isClose;
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                         ControlStyles.OptimizedDoubleBuffer, true);
                Width = 46;
                TabStop = false;
            }

            public void SetIcon(IconKind next) {
                icon = next;
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

            protected override void OnPaint(PaintEventArgs e) {
                Color back = AppTheme.Caption;
                if (hovering) {
                    back = isClose ? Color.FromArgb(196, 43, 60) : AppTheme.SecondaryHover;
                }
                e.Graphics.Clear(back);
                using (var pen = new Pen(AppTheme.Separator)) {
                    e.Graphics.DrawLine(pen, 0, Height - 1, Width, Height - 1);
                }

                Color fore = hovering && isClose ? Color.White : AppTheme.Text;
                const float size = 16f;
                Icons.Draw(e.Graphics, icon, new RectangleF((Width - size) / 2f, (Height - size) / 2f, size, size), fore, 1.3f);
            }
        }
    }

    /// <summary>Navigationseintrag der Sidebar.</summary>
    public sealed class NavItem : Control {
        private readonly IconKind icon;
        private bool selected;
        private bool hovering;

        internal NavItem(IconKind icon, string text) {
            this.icon = icon;
            Text = text;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer | ControlStyles.Selectable, true);
            Cursor = Cursors.Hand;
            TabStop = true;
        }

        public bool Selected {
            get { return selected; }
            set {
                if (selected != value) {
                    selected = value;
                    Invalidate();
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

        protected override void OnGotFocus(EventArgs e) {
            Invalidate();
            base.OnGotFocus(e);
        }

        protected override void OnLostFocus(EventArgs e) {
            Invalidate();
            base.OnLostFocus(e);
        }

        protected override void OnKeyDown(KeyEventArgs e) {
            base.OnKeyDown(e);
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space) {
                OnClick(EventArgs.Empty);
                e.Handled = true;
            }
        }

        protected override void OnPaint(PaintEventArgs e) {
            Graphics g = e.Graphics;
            g.Clear(AppTheme.Sidebar);

            if (hovering && !selected) {
                using (var brush = new SolidBrush(AppTheme.NavHover)) {
                    g.FillRectangle(brush, 0, 0, Width, Height);
                }
            }
            if (selected) {
                using (var brush = new SolidBrush(AppTheme.Accent)) {
                    g.FillRectangle(brush, 0, 8, 2, Height - 16);
                }
            }
            if (Focused && ShowFocusCues) {
                using (var pen = new Pen(AppTheme.Accent)) {
                    g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
                }
            }

            Color fore = selected || hovering ? AppTheme.Text : AppTheme.SubtleText;
            Icons.Draw(g, icon, new RectangleF(18f, (Height - 18) / 2f, 18f, 18f), fore, 1.5f);
            TextRenderer.DrawText(g, Text, AppTheme.NavFont, new Rectangle(46, 0, Math.Max(0, Width - 48), Height), fore,
                TextFormatFlags.NoPrefix | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis);
        }
    }
}
