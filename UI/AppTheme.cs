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
    /// Zentrale, formübergreifende Farbpalette und Schriften. Neutrale Grautöne mit dem Blau
    /// aus dem Firmenlogo als einziger Akzentfarbe.
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

        public static readonly Font CardTitleFont = new Font("Segoe UI Semibold", 10f, FontStyle.Regular);
        public static readonly Font StatCaptionFont = new Font("Segoe UI", 9f, FontStyle.Regular);
        public static readonly Font StatValueFont = new Font("Segoe UI Semibold", 14f, FontStyle.Regular);
        public static readonly Font NavFont = new Font("Segoe UI", 10f, FontStyle.Regular);
        public static readonly Font BrandFont = new Font("Segoe UI Semibold", 10f, FontStyle.Regular);
        public static readonly Font CaptionFont = new Font("Segoe UI", 9.5f, FontStyle.Regular);
        public static readonly Font MonoFont = new Font("Consolas", 9f, FontStyle.Regular);

        /// <summary>Firmenzeichen für die Titelleiste aller Fenster; wird einmalig beim Start gesetzt.</summary>
        public static Image BrandImage { get; set; }

        static AppTheme() {
            Refresh();
        }

        /// <summary>
        /// Liest die Benutzereinstellung (System/Light/Dark) und baut die Palette neu auf.
        /// </summary>
        public static void Refresh() {
            Apply(ResolveDark());
        }

        /// <summary>Der Modus, den die Einstellung und Windows gerade verlangen.</summary>
        public static bool ResolveDark() {
            return Services.ThemeMode.IsDark(Properties.Settings.Default.themeOverride, Services.ThemeMode.SystemUsesDarkMode());
        }

        private static void Apply(bool dark) {
            IsDark = dark;
            if (dark) {
                Accent = Color.FromArgb(40, 137, 233); // Blau aus dem Logo
                Background = Color.FromArgb(22, 24, 27);
                Surface = Background;
                SurfaceAlt = Color.FromArgb(32, 35, 39);
                Sidebar = Color.FromArgb(17, 19, 21);
                Caption = Sidebar;
                Footer = Sidebar;
                Text = Color.FromArgb(232, 234, 236);
                SubtleText = Color.FromArgb(150, 156, 163);
                Separator = Color.FromArgb(48, 52, 57);
                CardBorder = Separator;
                InputBorder = Color.FromArgb(70, 75, 82);
                WindowBorder = Color.FromArgb(52, 56, 61);
                Track = Separator;
                TableHeader = Background;
                TableLine = Color.FromArgb(38, 41, 45);
                NavSelected = Color.FromArgb(34, 52, 74);
                NavSelectedBorder = Accent;
                NavHover = Color.FromArgb(30, 33, 36);
                Primary = Accent;
                PrimaryHover = Color.FromArgb(64, 153, 240);
                Secondary = SurfaceAlt;
                SecondaryHover = Color.FromArgb(42, 46, 51);
                SecondaryBorder = InputBorder;
                Danger = SurfaceAlt;
                DangerHover = Color.FromArgb(58, 34, 36);
                Link = Color.FromArgb(96, 170, 242);
                LinkActive = Color.FromArgb(140, 195, 248);
                Error = Color.FromArgb(226, 110, 110);
                Success = Color.FromArgb(118, 190, 138);
                Warning = Color.FromArgb(214, 170, 92);
                OnAccent = Color.White;
            } else {
                Accent = Color.FromArgb(31, 120, 209); // Logo-Blau, für weißen Text leicht abgedunkelt
                Background = Color.FromArgb(250, 250, 249);
                Surface = Background;
                SurfaceAlt = Color.White;
                Sidebar = Color.FromArgb(241, 241, 239);
                Caption = Sidebar;
                Footer = Sidebar;
                Text = Color.FromArgb(28, 30, 33);
                SubtleText = Color.FromArgb(96, 102, 109);
                Separator = Color.FromArgb(222, 222, 219);
                CardBorder = Separator;
                InputBorder = Color.FromArgb(190, 192, 195);
                WindowBorder = Color.FromArgb(200, 200, 198);
                Track = Color.FromArgb(228, 228, 225);
                TableHeader = Background;
                TableLine = Color.FromArgb(236, 236, 233);
                NavSelected = Color.FromArgb(222, 235, 249);
                NavSelectedBorder = Accent;
                NavHover = Color.FromArgb(233, 233, 230);
                Primary = Accent;
                PrimaryHover = Color.FromArgb(22, 101, 180);
                Secondary = Color.White;
                SecondaryHover = Color.FromArgb(240, 240, 238);
                SecondaryBorder = InputBorder;
                Danger = Color.White;
                DangerHover = Color.FromArgb(252, 236, 236);
                Link = Color.FromArgb(24, 104, 184);
                LinkActive = Color.FromArgb(16, 80, 145);
                Error = Color.FromArgb(186, 50, 50);
                Success = Color.FromArgb(40, 130, 75);
                Warning = Color.FromArgb(150, 100, 10);
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

        /// <summary>Mischt zwei Farben deckend; <paramref name="amount"/> ist der Anteil von <paramref name="color"/>.</summary>
        public static Color Blend(Color color, Color background, float amount) {
            return Color.FromArgb(
                (int)(color.R * amount + background.R * (1f - amount)),
                (int)(color.G * amount + background.G * (1f - amount)),
                (int)(color.B * amount + background.B * (1f - amount)));
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
