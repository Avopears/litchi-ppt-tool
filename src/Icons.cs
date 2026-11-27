using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Windows.Forms;

namespace PptSelectionExport
{
    /// <summary>Ribbon icons, drawn in code so the add-in is a single DLL.</summary>
    internal static class Icons
    {
        private static float _scale;

        private static float Scale
        {
            get
            {
                if (_scale <= 0)
                {
                    try
                    {
                        using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) _scale = Math.Max(1f, Math.Min(4f, g.DpiX / 96f));
                    }
                    catch { _scale = 1f; }
                }
                return _scale;
            }
        }

        private static Bitmap Canvas(int logical, out Graphics g, out float s)
        {
            s = Scale;
            int px = (int)Math.Round(logical * s);
            var bmp = new Bitmap(px, px, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            g = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.ScaleTransform(s, s);
            return bmp;
        }

        /// <summary>A sheet of paper with a coloured label ("PDF" / "SVG"), 32 logical pixels.</summary>
        public static Bitmap Document(string label, Color accent)
        {
            Graphics g; float s;
            Bitmap bmp = Canvas(32, out g, out s);
            using (g)
            {
                // page with a folded corner
                using (var page = new GraphicsPath())
                {
                    page.AddPolygon(new[] { new PointF(6.5f, 1.5f), new PointF(19.5f, 1.5f), new PointF(26.5f, 8.5f), new PointF(26.5f, 30.5f), new PointF(6.5f, 30.5f) });
                    g.FillPath(Brushes.White, page);
                    using (var pen = new Pen(Color.FromArgb(120, 120, 120), 1f)) g.DrawPath(pen, page);
                }
                using (var fold = new GraphicsPath())
                {
                    fold.AddPolygon(new[] { new PointF(19.5f, 1.5f), new PointF(19.5f, 8.5f), new PointF(26.5f, 8.5f) });
                    using (var b = new SolidBrush(Color.FromArgb(215, 215, 215))) g.FillPath(b, fold);
                    using (var pen = new Pen(Color.FromArgb(120, 120, 120), 1f)) g.DrawPath(pen, fold);
                }
                // selection marquee hint
                using (var pen = new Pen(Color.FromArgb(150, 150, 150), 1f) { DashPattern = new[] { 2f, 1.5f } })
                    g.DrawRectangle(pen, 10.5f, 6.5f, 7f, 5f);
                // label
                using (var b = new SolidBrush(accent))
                using (var band = RoundRect(new RectangleF(1.5f, 14.5f, 25f, 12f), 2.2f))
                    g.FillPath(b, band);
                using (var font = new Font("Segoe UI", 7.6f, FontStyle.Bold, GraphicsUnit.Pixel))
                using (var fmt = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                    g.DrawString(label, font, Brushes.White, new RectangleF(1.5f, 14.5f, 25f, 12.6f), fmt);
            }
            return bmp;
        }

        /// <summary>Checkerboard = "no background".</summary>
        public static Bitmap Checker(int logical)
        {
            Graphics g; float s;
            Bitmap bmp = Canvas(logical, out g, out s);
            using (g)
            {
                float m = logical >= 32 ? 4f : 1.5f, size = logical - 2 * m;
                var box = new RectangleF(m, m, size, size);
                using (GraphicsPath clip = RoundRect(box, logical >= 32 ? 3f : 2f))
                {
                    g.SetClip(clip);
                    g.FillRectangle(Brushes.White, box);
                    int cells = logical >= 32 ? 4 : 3;
                    float c = size / cells;
                    using (var b = new SolidBrush(Color.FromArgb(176, 176, 176)))
                        for (int y = 0; y < cells; y++)
                            for (int x = 0; x < cells; x++)
                                if (((x + y) & 1) == 0) g.FillRectangle(b, m + x * c, m + y * c, c, c);
                    g.ResetClip();
                    using (var pen = new Pen(Color.FromArgb(110, 110, 110), 1f)) g.DrawPath(pen, clip);
                }
            }
            return bmp;
        }

        /// <summary>A swatch filled with the given colour.</summary>
        public static Bitmap Swatch(Color color, int logical)
        {
            Graphics g; float s;
            Bitmap bmp = Canvas(logical, out g, out s);
            using (g)
            {
                float m = logical >= 32 ? 4f : 1.5f, size = logical - 2 * m;
                using (GraphicsPath p = RoundRect(new RectangleF(m, m, size, size), logical >= 32 ? 3f : 2f))
                {
                    using (var b = new SolidBrush(Color.FromArgb(255, color))) g.FillPath(b, p);
                    using (var pen = new Pen(Color.FromArgb(110, 110, 110), 1f)) g.DrawPath(pen, p);
                }
            }
            return bmp;
        }

        /// <summary>Colour wheel for "more colours".</summary>
        public static Bitmap Palette(int logical)
        {
            Graphics g; float s;
            Bitmap bmp = Canvas(logical, out g, out s);
            using (g)
            {
                var box = new RectangleF(1.5f, 1.5f, logical - 3f, logical - 3f);
                Color[] cs = { Color.FromArgb(232, 65, 53), Color.FromArgb(251, 188, 5), Color.FromArgb(52, 168, 83), Color.FromArgb(66, 133, 244), Color.FromArgb(156, 39, 176), Color.FromArgb(255, 112, 67) };
                float sweep = 360f / cs.Length;
                for (int i = 0; i < cs.Length; i++)
                    using (var b = new SolidBrush(cs[i])) g.FillPie(b, box.X, box.Y, box.Width, box.Height, -90 + i * sweep, sweep + 0.5f);
                using (var pen = new Pen(Color.FromArgb(110, 110, 110), 1f)) g.DrawEllipse(pen, box);
            }
            return bmp;
        }

        private static GraphicsPath RoundRect(RectangleF r, float radius)
        {
            float d = radius * 2;
            var p = new GraphicsPath();
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        /// <summary>Office wants an IPictureDisp from getImage callbacks.</summary>
        public static object ToPictureDisp(Image image)
        {
            return PictureHost.Convert(image);
        }

        private sealed class PictureHost : AxHost
        {
            private PictureHost() : base("") { }
            public static object Convert(Image image) { return GetIPictureDispFromPicture(image); }
        }
    }
}
