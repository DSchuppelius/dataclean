using System;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace DataClean.UI {
    public enum IconKind {
        None,
        Skip,
        Home,
        Trash,
        Document,
        Gear,
        Info,
        Question,
        Warning,
        CheckCircle,
        ErrorCircle,
        Pause,
        Play,
        Minimize,
        Maximize,
        Restore,
        Close,
        ChevronDown,
        ChevronUp,
        Folder,
        Clock,
        Dot
    }

    /// <summary>
    /// Vektor-Icons im 24×24-Raster, per GDI+ gezeichnet. Dadurch sind sie unabhängig von
    /// Symbolschriften (Segoe Fluent/MDL2), die auf älteren Systemen fehlen.
    /// </summary>
    internal static class Icons {
        public static void Draw(Graphics g, IconKind kind, RectangleF bounds, Color color, float stroke = 1.7f) {
            if (kind == IconKind.None) {
                return;
            }

            GraphicsState state = g.Save();
            try {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                float size = Math.Min(bounds.Width, bounds.Height);
                g.TranslateTransform(bounds.X + (bounds.Width - size) / 2f, bounds.Y + (bounds.Height - size) / 2f);
                g.ScaleTransform(size / 24f, size / 24f);

                using (var pen = new Pen(color, stroke * 24f / size) {
                    StartCap = LineCap.Round,
                    EndCap = LineCap.Round,
                    LineJoin = LineJoin.Round
                })
                using (var brush = new SolidBrush(color)) {
                    DrawShape(g, kind, pen, brush, color);
                }
            } finally {
                g.Restore(state);
            }
        }

        private static void DrawShape(Graphics g, IconKind kind, Pen pen, Brush brush, Color color) {
            switch (kind) {
                case IconKind.Skip:
                    Polyline(g, pen, P(5f, 5f), P(12f, 12f), P(5f, 19f));
                    Polyline(g, pen, P(12.5f, 5f), P(19.5f, 12f), P(12.5f, 19f));
                    break;
                case IconKind.Home:
                    Polyline(g, pen, P(2.8f, 11.8f), P(12f, 3.5f), P(21.2f, 11.8f));
                    Polyline(g, pen, P(5.5f, 9.6f), P(5.5f, 20.5f), P(18.5f, 20.5f), P(18.5f, 9.6f));
                    Polyline(g, pen, P(9.7f, 20.5f), P(9.7f, 14f), P(14.3f, 14f), P(14.3f, 20.5f));
                    break;
                case IconKind.Trash:
                    g.DrawLine(pen, 4f, 6.5f, 20f, 6.5f);
                    Polyline(g, pen, P(9f, 6.5f), P(9f, 3.5f), P(15f, 3.5f), P(15f, 6.5f));
                    Polyline(g, pen, P(6f, 6.5f), P(7f, 21f), P(17f, 21f), P(18f, 6.5f));
                    break;
                case IconKind.Document:
                    using (var path = AppTheme.RoundedPath(new Rectangle(5, 3, 14, 18), 2)) {
                        g.DrawPath(pen, path);
                    }
                    g.DrawLine(pen, 8.5f, 8f, 15.5f, 8f);
                    g.DrawLine(pen, 8.5f, 12f, 15.5f, 12f);
                    g.DrawLine(pen, 8.5f, 16f, 15.5f, 16f);
                    break;
                case IconKind.Gear:
                    DrawGear(g, pen);
                    break;
                case IconKind.Info:
                    g.DrawEllipse(pen, 3f, 3f, 18f, 18f);
                    g.FillEllipse(brush, 10.8f, 6.6f, 2.4f, 2.4f);
                    g.DrawLine(pen, 12f, 11.5f, 12f, 16.8f);
                    break;
                case IconKind.Question:
                    g.DrawEllipse(pen, 3f, 3f, 18f, 18f);
                    g.DrawArc(pen, 9f, 7f, 6f, 6f, 180f, 250f);
                    g.DrawLine(pen, 13f, 12.6f, 12f, 14f);
                    g.FillEllipse(brush, 10.9f, 15.8f, 2.2f, 2.2f);
                    break;
                case IconKind.Warning:
                    Polygon(g, pen, P(12f, 3.5f), P(22f, 20.5f), P(2f, 20.5f));
                    g.DrawLine(pen, 12f, 9.8f, 12f, 14.5f);
                    g.FillEllipse(brush, 10.9f, 16.2f, 2.2f, 2.2f);
                    break;
                case IconKind.CheckCircle:
                    g.FillEllipse(brush, 2f, 2f, 20f, 20f);
                    using (var mark = MarkPen(pen)) {
                        Polyline(g, mark, P(7.3f, 12.4f), P(10.6f, 15.6f), P(16.7f, 8.9f));
                    }
                    break;
                case IconKind.ErrorCircle:
                    g.FillEllipse(brush, 2f, 2f, 20f, 20f);
                    using (var mark = MarkPen(pen)) {
                        g.DrawLine(mark, 8.3f, 8.3f, 15.7f, 15.7f);
                        g.DrawLine(mark, 15.7f, 8.3f, 8.3f, 15.7f);
                    }
                    break;
                case IconKind.Pause:
                    using (var bars = new Pen(color, 3.2f) { StartCap = LineCap.Round, EndCap = LineCap.Round }) {
                        g.DrawLine(bars, 8f, 4.5f, 8f, 19.5f);
                        g.DrawLine(bars, 16f, 4.5f, 16f, 19.5f);
                    }
                    break;
                case IconKind.Play:
                    g.FillPolygon(brush, new[] { P(7f, 4f), P(20f, 12f), P(7f, 20f) });
                    g.DrawPolygon(pen, new[] { P(7f, 4f), P(20f, 12f), P(7f, 20f) });
                    break;
                case IconKind.Minimize:
                    g.DrawLine(pen, 5f, 12f, 19f, 12f);
                    break;
                case IconKind.Maximize:
                    g.DrawRectangle(pen, 5.5f, 5.5f, 13f, 13f);
                    break;
                case IconKind.Restore:
                    g.DrawRectangle(pen, 4.5f, 8.5f, 11f, 11f);
                    Polyline(g, pen, P(8.5f, 8.5f), P(8.5f, 4.5f), P(19.5f, 4.5f), P(19.5f, 15.5f), P(15.5f, 15.5f));
                    break;
                case IconKind.Close:
                    g.DrawLine(pen, 5.5f, 5.5f, 18.5f, 18.5f);
                    g.DrawLine(pen, 18.5f, 5.5f, 5.5f, 18.5f);
                    break;
                case IconKind.ChevronDown:
                    Polyline(g, pen, P(6f, 9f), P(12f, 15f), P(18f, 9f));
                    break;
                case IconKind.ChevronUp:
                    Polyline(g, pen, P(6f, 15f), P(12f, 9f), P(18f, 15f));
                    break;
                case IconKind.Folder:
                    Polygon(g, pen, P(3f, 19.5f), P(3f, 5f), P(9.5f, 5f), P(11.5f, 8f), P(21f, 8f), P(21f, 19.5f));
                    break;
                case IconKind.Clock:
                    g.DrawEllipse(pen, 3f, 3f, 18f, 18f);
                    Polyline(g, pen, P(12f, 7f), P(12f, 12f), P(15.5f, 14f));
                    break;
                case IconKind.Dot:
                    g.FillEllipse(brush, 8f, 8f, 8f, 8f);
                    break;
            }
        }

        private static Pen MarkPen(Pen source) {
            return new Pen(AppTheme.OnAccent, Math.Max(source.Width, 2.3f)) {
                StartCap = LineCap.Round,
                EndCap = LineCap.Round,
                LineJoin = LineJoin.Round
            };
        }

        private static void DrawGear(Graphics g, Pen pen) {
            const int teeth = 8;
            const float outer = 10.4f, inner = 7.9f;
            var points = new PointF[teeth * 4];
            for (int i = 0; i < teeth; i++) {
                double center = i * (2 * Math.PI / teeth) - Math.PI / 2;
                double tip = 0.18, root = 0.29;
                points[i * 4] = Polar(inner, center - root);
                points[i * 4 + 1] = Polar(outer, center - tip);
                points[i * 4 + 2] = Polar(outer, center + tip);
                points[i * 4 + 3] = Polar(inner, center + root);
            }
            g.DrawPolygon(pen, points);
            g.DrawEllipse(pen, 8.6f, 8.6f, 6.8f, 6.8f);
        }

        private static PointF Polar(float radius, double angle) {
            return new PointF(12f + radius * (float)Math.Cos(angle), 12f + radius * (float)Math.Sin(angle));
        }

        private static PointF P(float x, float y) {
            return new PointF(x, y);
        }

        private static void Polyline(Graphics g, Pen pen, params PointF[] points) {
            g.DrawLines(pen, points);
        }

        private static void Polygon(Graphics g, Pen pen, params PointF[] points) {
            g.DrawPolygon(pen, points);
        }
    }
}
