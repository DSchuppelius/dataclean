using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace DataClean.UI {
    /// <summary>Darstellungsvarianten der <see cref="ThemedButton"/>.</summary>
    public enum ButtonKind {
        Primary,
        Secondary,
        Danger
    }

    /// <summary>Semantische Farbtöne für Icons, Kennzahlen und Hinweiszeilen.</summary>
    public enum Tone {
        Accent,
        Success,
        Warning,
        Error,
        Muted
    }

    /// <summary>
    /// Zentrale, formübergreifende Farbpalette und Schriften des Dashboard-Designs
    /// (dunkles Navy mit Cyan-Akzent, siehe design.png).
    /// Reines GDI+/WinForms – lauffähig auch auf älteren Systemen (.NET Framework 4.8).
    /// </summary>
    internal static class AppTheme {
        public static Color Accent { get; private set; }
        public static Color Background { get; private set; }
        public static Color Surface { get; private set; }
        public static Color SurfaceAlt { get; private set; }
        public static Color Sidebar { get; private set; }
        public static Color Caption { get; private set; }
        public static Color Footer { get; private set; }
        public static Color Text { get; private set; }
        public static Color SubtleText { get; private set; }
        public static Color Separator { get; private set; }
        public static Color CardBorder { get; private set; }
        public static Color InputBorder { get; private set; }
        public static Color WindowBorder { get; private set; }
        public static Color Track { get; private set; }
        public static Color TableHeader { get; private set; }
        public static Color TableLine { get; private set; }
        public static Color NavSelected { get; private set; }
        public static Color NavSelectedBorder { get; private set; }
        public static Color NavHover { get; private set; }
        public static Color Primary { get; private set; }
        public static Color PrimaryHover { get; private set; }
        public static Color Secondary { get; private set; }
        public static Color SecondaryHover { get; private set; }
        public static Color SecondaryBorder { get; private set; }
        public static Color Danger { get; private set; }
        public static Color DangerHover { get; private set; }
        public static Color Link { get; private set; }
        public static Color LinkActive { get; private set; }
        public static Color Error { get; private set; }
        public static Color Success { get; private set; }
        public static Color Warning { get; private set; }
        public static Color OnAccent { get; private set; }

        public static bool IsDark { get; private set; }

        public static readonly Font HeadingFont = new Font("Segoe UI", 20f, FontStyle.Bold);
        public static readonly Font CardTitleFont = new Font("Segoe UI Semibold", 12f, FontStyle.Regular);
        public static readonly Font StatCaptionFont = new Font("Segoe UI", 10f, FontStyle.Regular);
        public static readonly Font StatValueFont = new Font("Segoe UI", 18f, FontStyle.Bold);
        public static readonly Font NavFont = new Font("Segoe UI", 10.5f, FontStyle.Regular);
        public static readonly Font BrandFont = new Font("Segoe UI", 11f, FontStyle.Bold);
        public static readonly Font CaptionFont = new Font("Segoe UI", 10f, FontStyle.Regular);

        static AppTheme() {
            Refresh();
        }

        /// <summary>
        /// Liest die Benutzereinstellung (System/Light/Dark) und baut die Palette neu auf.
        /// Standard bei "System"/leer ist dunkel, passend zum Dashboard-Design.
        /// </summary>
        public static void Refresh() {
            Apply(ResolveDark());
        }

        private static bool ResolveDark() {
            string mode = Properties.Settings.Default.themeOverride;
            return !string.Equals(mode, "Light", StringComparison.OrdinalIgnoreCase);
        }

        private static void Apply(bool dark) {
            IsDark = dark;
            if (dark) {
                Accent = Color.FromArgb(30, 240, 245);
                Background = Color.FromArgb(14, 22, 32);
                Surface = Color.FromArgb(18, 29, 40);
                SurfaceAlt = Color.FromArgb(28, 42, 55);
                Sidebar = Color.FromArgb(16, 28, 38);
                Caption = Color.FromArgb(23, 33, 44);
                Footer = Color.FromArgb(17, 28, 39);
                Text = Color.FromArgb(240, 245, 248);
                SubtleText = Color.FromArgb(190, 202, 212);
                Separator = Color.FromArgb(44, 56, 68);
                CardBorder = Color.FromArgb(74, 91, 104);
                InputBorder = Color.FromArgb(73, 88, 102);
                WindowBorder = Color.FromArgb(58, 72, 85);
                Track = Color.FromArgb(50, 66, 79);
                TableHeader = Color.FromArgb(36, 55, 65);
                TableLine = Color.FromArgb(44, 58, 69);
                NavSelected = Color.FromArgb(17, 78, 93);
                NavSelectedBorder = Color.FromArgb(38, 132, 146);
                NavHover = Color.FromArgb(26, 42, 56);
                Primary = Color.FromArgb(19, 32, 41);
                PrimaryHover = Color.FromArgb(19, 58, 70);
                Secondary = Color.FromArgb(29, 46, 58);
                SecondaryHover = Color.FromArgb(40, 63, 78);
                SecondaryBorder = Color.FromArgb(107, 127, 141);
                Danger = Color.FromArgb(58, 28, 36);
                DangerHover = Color.FromArgb(88, 36, 46);
                Link = Color.FromArgb(45, 228, 236);
                LinkActive = Color.FromArgb(150, 245, 250);
                Error = Color.FromArgb(255, 107, 127);
                Success = Color.FromArgb(104, 222, 112);
                Warning = Color.FromArgb(250, 196, 84);
                OnAccent = Color.FromArgb(10, 22, 30);
            } else {
                Accent = Color.FromArgb(0, 132, 160);
                Background = Color.FromArgb(242, 246, 250);
                Surface = Color.White;
                SurfaceAlt = Color.FromArgb(246, 249, 252);
                Sidebar = Color.FromArgb(231, 238, 244);
                Caption = Color.FromArgb(224, 233, 240);
                Footer = Color.FromArgb(231, 238, 244);
                Text = Color.FromArgb(20, 32, 44);
                SubtleText = Color.FromArgb(82, 98, 112);
                Separator = Color.FromArgb(208, 218, 226);
                CardBorder = Color.FromArgb(190, 203, 214);
                InputBorder = Color.FromArgb(160, 176, 190);
                WindowBorder = Color.FromArgb(160, 176, 190);
                Track = Color.FromArgb(214, 224, 232);
                TableHeader = Color.FromArgb(226, 235, 241);
                TableLine = Color.FromArgb(222, 230, 237);
                NavSelected = Color.FromArgb(200, 234, 240);
                NavSelectedBorder = Color.FromArgb(120, 196, 210);
                NavHover = Color.FromArgb(219, 229, 237);
                Primary = Color.FromArgb(234, 248, 251);
                PrimaryHover = Color.FromArgb(208, 238, 244);
                Secondary = Color.FromArgb(235, 241, 246);
                SecondaryHover = Color.FromArgb(220, 230, 238);
                SecondaryBorder = Color.FromArgb(150, 167, 181);
                Danger = Color.FromArgb(253, 236, 239);
                DangerHover = Color.FromArgb(250, 218, 224);
                Link = Color.FromArgb(0, 112, 140);
                LinkActive = Color.FromArgb(0, 82, 106);
                Error = Color.FromArgb(190, 40, 62);
                Success = Color.FromArgb(28, 134, 62);
                Warning = Color.FromArgb(166, 106, 0);
                OnAccent = Color.White;
            }
        }

        public static Color ToneColor(Tone tone) {
            switch (tone) {
                case Tone.Success: return Success;
                case Tone.Warning: return Warning;
                case Tone.Error: return Error;
                case Tone.Muted: return SubtleText;
                default: return Accent;
            }
        }

        /// <summary>Erster deckender Hintergrund in der Eltern-Kette eines Controls.</summary>
        public static Color EffectiveBack(Control control) {
            for (Control parent = control?.Parent; parent != null; parent = parent.Parent) {
                if (parent.BackColor.A == 255) {
                    return parent.BackColor;
                }
            }
            return Background;
        }

        /// <summary>Weist ein zum Modus passendes Fenster-Theme zu (z. B. dunkle Scrollbars).</summary>
        public static void ApplyNativeTheme(Control control) {
            try {
                NativeMethods.SetWindowTheme(control.Handle, IsDark ? "DarkMode_Explorer" : "Explorer", null);
            } catch (Exception ex) when (ex is DllNotFoundException || ex is EntryPointNotFoundException) {
                // Ältere Systeme ohne Theme-API behalten die Standarddarstellung.
            }
        }

        /// <summary>Erzeugt einen abgerundeten Pfad für Buttons/Panels.</summary>
        public static GraphicsPath RoundedPath(Rectangle bounds, int radius) {
            var path = new GraphicsPath();
            int d = Math.Min(radius * 2, Math.Min(bounds.Width, bounds.Height));
            if (d <= 0) {
                path.AddRectangle(bounds);
                path.CloseFigure();
                return path;
            }
            path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
            path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
            path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }
}
