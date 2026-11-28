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

        /// <summary>A slide with a measuring arrow along its width and one along its height: "slide size", 32 logical pixels.</summary>
        public static Bitmap SlideSize()
        {
            Graphics g; float s;
            Bitmap bmp = Canvas(32, out g, out s);
            using (g)
            {
                var slide = new RectangleF(2.5f, 3.5f, 20f, 15f);
                g.FillRectangle(Brushes.White, slide);
                using (var pen = new Pen(Color.FromArgb(120, 120, 120), 1f)) g.DrawRectangle(pen, slide.X, slide.Y, slide.Width, slide.Height);
                using (var b = new SolidBrush(Color.FromArgb(215, 215, 215))) g.FillRectangle(b, 5.5f, 6.5f, 14f, 3f);        // a title on it
                using (var b = new SolidBrush(Color.FromArgb(40, 110, 200)))
                {
                    // along the width, below the slide
                    g.FillRectangle(b, 5.5f, 24.75f, 14f, 1.5f);
                    g.FillPolygon(b, new[] { new PointF(2.5f, 25.5f), new PointF(6.5f, 22.5f), new PointF(6.5f, 28.5f) });
                    g.FillPolygon(b, new[] { new PointF(22.5f, 25.5f), new PointF(18.5f, 22.5f), new PointF(18.5f, 28.5f) });
                    // along the height, to the right of it
                    g.FillRectangle(b, 26.75f, 6.5f, 1.5f, 9f);
                    g.FillPolygon(b, new[] { new PointF(27.5f, 3.5f), new PointF(24.5f, 7.5f), new PointF(30.5f, 7.5f) });
                    g.FillPolygon(b, new[] { new PointF(27.5f, 18.5f), new PointF(24.5f, 14.5f), new PointF(30.5f, 14.5f) });
                }
            }
            return bmp;
        }

        /// <summary>A sheet of paper being pressed together from above and below: "slim down", 32 logical pixels.</summary>
        public static Bitmap Slim()
        {
            Graphics g; float s;
            Bitmap bmp = Canvas(32, out g, out s);
            using (g)
            {
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
                // two arrows pressing against a line in the middle
                using (var b = new SolidBrush(Color.FromArgb(40, 110, 200)))
                {
                    g.FillRectangle(b, 15f, 9.5f, 3f, 3.5f);
                    g.FillPolygon(b, new[] { new PointF(12f, 12.5f), new PointF(21f, 12.5f), new PointF(16.5f, 17.5f) });
                    g.FillRectangle(b, 10f, 18.75f, 13f, 1.5f);
                    g.FillPolygon(b, new[] { new PointF(12f, 26.5f), new PointF(21f, 26.5f), new PointF(16.5f, 21.5f) });
                    g.FillRectangle(b, 15f, 26f, 3f, 3f);
                }
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

        // ------------------------------------------------------------ alignment glyphs

        private static readonly Color AlignRed = Color.FromArgb(247, 52, 52);

        /// <summary>Glyph of an align or distribute button.</summary>
        public static Bitmap Align(AlignCommand cmd, int logical)
        {
            switch (cmd)
            {
                case AlignCommand.Left: return Glyph(logical, R(1, 1, 1, 14), R(3, 3, 11, 4), R(3, 9, 7, 4));
                case AlignCommand.Center: return Glyph(logical, R(7, 1, 1, 14), R(2, 3, 11, 4), R(4, 9, 7, 4));
                case AlignCommand.Right: return Glyph(logical, R(14, 1, 1, 14), R(2, 3, 11, 4), R(6, 9, 7, 4));
                case AlignCommand.Top: return Glyph(logical, R(1, 1, 14, 1), R(3, 3, 4, 7), R(9, 3, 4, 11));
                case AlignCommand.Middle: return Glyph(logical, R(1, 7, 14, 1), R(3, 2, 4, 11), R(9, 4, 4, 7));
                case AlignCommand.Bottom: return Glyph(logical, R(1, 14, 14, 1), R(3, 2, 4, 11), R(9, 6, 4, 7));
                case AlignCommand.DistributeHorizontally: return Glyph(logical, R(1, 1, 1, 14), R(14, 1, 1, 14), R(6, 3, 4, 10));
                default: return Glyph(logical, R(1, 1, 14, 1), R(1, 14, 14, 1), R(3, 6, 10, 4));   // distribute vertically
            }
        }

        /// <summary>Glyph of "align to object": small squares stepping towards a bigger one.</summary>
        public static Bitmap AlignToObject(int logical)
        {
            return Glyph(logical, R(1, 1, 4, 4), R(5, 5, 4, 4), R(9, 9, 6, 6));
        }

        /// <summary>Glyph of a rotate or flip button (after iSlide's: a frame with a turning arrow, mirrored triangles).</summary>
        public static Bitmap Transform(TransformCommand cmd, int logical)
        {
            int px = (int)Math.Round(logical * Scale);
            float u = px / 16f;
            var bmp = new Bitmap(px, px, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            using (var b = new SolidBrush(AlignRed))
            using (var pen = new Pen(AlignRed, Math.Max(1f, 1.35f * u)) { LineJoin = LineJoin.Miter })
            {
                if (cmd == TransformCommand.RotateRight)
                {
                    g.TranslateTransform(px, 0);       // the same picture, mirrored
                    g.ScaleTransform(-1, 1);
                }
                switch (cmd)
                {
                    case TransformCommand.RotateLeft:
                    case TransformCommand.RotateRight:
                        g.SmoothingMode = SmoothingMode.None;
                        OutlineUnits(g, b, R(1, 7, 11, 7), u);
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        float rad = 6.3f * u;
                        g.DrawArc(pen, 8.2f * u - rad, 10.2f * u - rad, 2 * rad, 2 * rad, 0, -84);
                        g.FillPolygon(b, new[] { P(4.6f, 3.9f, u), P(8.9f, 1.2f, u), P(8.9f, 6.6f, u) });
                        break;
                    case TransformCommand.FlipHorizontal:
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        g.FillPolygon(b, new[] { P(7, 1, u), P(7, 15, u), P(1, 15, u) });
                        g.DrawPolygon(pen, new[] { P(9.6f, 2.8f, u), P(9.6f, 14.4f, u), P(14.3f, 14.4f, u) });
                        break;
                    default:   // flip vertical
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        g.FillPolygon(b, new[] { P(1, 1, u), P(15, 7, u), P(1, 7, u) });
                        g.DrawPolygon(pen, new[] { P(1.6f, 9.6f, u), P(13.2f, 9.6f, u), P(1.6f, 14.3f, u) });
                        break;
                }
            }
            return bmp;
        }

        /// <summary>
        /// Glyph of a boolean button: two overlapping squares, drawn as the result looks
        /// (merged, overlap cut out, front square cut away, or split into pieces).
        /// </summary>
        public static Bitmap Merge(MergeCommand cmd, int logical)
        {
            int px = (int)Math.Round(logical * Scale);
            float u = px / 16f;
            var bmp = new Bitmap(px, px, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            using (var b = new SolidBrush(AlignRed))
            using (var light = new SolidBrush(Color.FromArgb(250, 168, 168)))
            using (var clear = new SolidBrush(Color.Transparent))
            {
                g.SmoothingMode = SmoothingMode.None;
                float[] back = R(1, 1, 9, 9), front = R(6, 6, 9, 9), overlap = R(6, 6, 4, 4);
                switch (cmd)
                {
                    case MergeCommand.Union:
                        FillUnits(g, b, back, u); FillUnits(g, b, front, u);
                        break;
                    case MergeCommand.Combine:
                        FillUnits(g, b, back, u); FillUnits(g, b, front, u);
                        g.CompositingMode = CompositingMode.SourceCopy;
                        FillUnits(g, clear, overlap, u);
                        break;
                    case MergeCommand.Subtract:
                        OutlineUnits(g, light, front, u);    // the cutter, faint
                        FillUnits(g, b, back, u);
                        g.CompositingMode = CompositingMode.SourceCopy;
                        FillUnits(g, clear, overlap, u);
                        break;
                    default:   // fragment: three separate pieces
                        foreach (float[] r in new[] { R(1, 1, 9, 4), R(1, 5, 4, 5), R(6, 6, 4, 4), R(11, 6, 4, 9), R(6, 11, 5, 4) })
                            FillUnits(g, b, r, u);
                        break;
                }
            }
            return bmp;
        }

        // ------------------------------------------------------------ stacking order

        /// <summary>
        /// Glyph of a stacking-order button: overlapping squares with the one that moves in full red.
        /// Three squares for "to front / to back" (all the way), two for "forward / backward" (one step).
        /// </summary>
        public static Bitmap Order(OrderCommand cmd, int logical)
        {
            int px = (int)Math.Round(logical * Scale);
            float u = px / 16f;
            var bmp = new Bitmap(px, px, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            using (var red = new SolidBrush(AlignRed))
            using (var light = new SolidBrush(Color.FromArgb(250, 168, 168)))
            {
                g.SmoothingMode = SmoothingMode.None;
                switch (cmd)      // drawn back to front; the square drawn last is on top
                {
                    case OrderCommand.BringToFront:
                        Layer(g, light, R(1, 1, 8, 8), u); Layer(g, light, R(4, 4, 8, 8), u); Layer(g, red, R(7, 7, 8, 8), u);
                        break;
                    case OrderCommand.SendToBack:
                        Layer(g, red, R(1, 1, 8, 8), u); Layer(g, light, R(4, 4, 8, 8), u); Layer(g, light, R(7, 7, 8, 8), u);
                        break;
                    case OrderCommand.BringForward:
                        Layer(g, light, R(1, 1, 9, 9), u); Layer(g, red, R(6, 6, 9, 9), u);
                        break;
                    default:
                        Layer(g, red, R(1, 1, 9, 9), u); Layer(g, light, R(6, 6, 9, 9), u);
                        break;
                }
            }
            return bmp;
        }

        /// <summary>A square that lies on top of what was drawn before: a thin gap is cut around it first.</summary>
        private static void Layer(Graphics g, Brush b, float[] r, float unit)
        {
            g.CompositingMode = CompositingMode.SourceCopy;
            using (var clear = new SolidBrush(Color.Transparent)) FillUnits(g, clear, R(r[0] - 1, r[1] - 1, r[2] + 2, r[3] + 2), unit);
            g.CompositingMode = CompositingMode.SourceOver;
            FillUnits(g, b, r, unit);
        }

        // ------------------------------------------------------------ shape style

        /// <summary>
        /// "Fill" (a solid square) or "outline" (a frame) above a bar in the colour that is
        /// current for it; the bar is empty with a red stroke for "none".
        /// </summary>
        public static Bitmap StyleColor(bool outline, Color? color, int logical)
        {
            int px = (int)Math.Round(logical * Scale);
            float u = px / 16f;
            var bmp = new Bitmap(px, px, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            using (var red = new SolidBrush(AlignRed))
            {
                g.SmoothingMode = SmoothingMode.None;
                float[] glyph = R(3, 0, 10, 10);
                if (outline)
                {
                    FillUnits(g, red, glyph, u);
                    g.CompositingMode = CompositingMode.SourceCopy;
                    using (var clear = new SolidBrush(Color.Transparent)) FillUnits(g, clear, R(5, 2, 6, 6), u);
                    g.CompositingMode = CompositingMode.SourceOver;
                }
                else FillUnits(g, red, glyph, u);

                // the colour bar; pale colours and "none" get a hairline so that they still show
                int x0, x1, y0, y1;
                SnapSpan(1, 14, u, out x0, out x1);
                SnapSpan(11.5f, 4.5f, u, out y0, out y1);
                var bar = new Rectangle(x0, y0, x1 - x0, y1 - y0);
                Color fill = color.HasValue ? Color.FromArgb(255, color.Value) : Color.White;
                using (var b = new SolidBrush(fill)) g.FillRectangle(b, bar);
                if (fill.GetBrightness() > 0.8f)
                    using (var pen = new Pen(Color.FromArgb(140, 140, 140), 1f)) g.DrawRectangle(pen, bar.X, bar.Y, bar.Width - 1, bar.Height - 1);
                if (!color.HasValue)
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    using (var pen = new Pen(AlignRed, Math.Max(1f, u))) g.DrawLine(pen, bar.Left + 1, bar.Bottom - 1, bar.Right - 1, bar.Top);
                }
            }
            return bmp;
        }

        /// <summary>Glyph of "weight": three lines, thin to thick.</summary>
        public static Bitmap Weight(int logical)
        {
            return Glyph(logical, R(1, 2, 14, 1), R(1, 6, 14, 2), R(1, 11, 14, 3));
        }

        /// <summary>Glyph of "dashes": a solid, a dashed and a dotted line.</summary>
        public static Bitmap Dashes(int logical)
        {
            return Glyph(logical, R(1, 2, 14, 2),
                R(1, 7, 4, 2), R(6, 7, 4, 2), R(11, 7, 4, 2),
                R(1, 12, 2, 2), R(5, 12, 2, 2), R(9, 12, 2, 2), R(13, 12, 2, 2));
        }

        // Cap types are drawn the way drawing programs show them: the stroke in a pale tone, and on
        // top of it the line itself with a dot on its end point. A flat cap stops at the dot, a
        // square or a round one reaches half the stroke's thickness past it.
        private static readonly Color CapStroke = Color.FromArgb(105, AlignRed);

        /// <summary>Glyph of "cap type": the end of a thick line with a round cap.</summary>
        public static Bitmap Cap(int logical)
        {
            int px = (int)Math.Round(logical * Scale);
            float u = px / 16f;
            var bmp = new Bitmap(px, px, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            using (var red = new SolidBrush(AlignRed))
            using (var pale = new SolidBrush(CapStroke))
            {
                int x0, x1, y0, y1;
                SnapSpan(0, 9, u, out x0, out x1);           // the line itself ends at x = 9
                SnapSpan(3, 10, u, out y0, out y1);
                FillCapShape(g, pale, x0, x1, y0, y1 - y0, CapKind.Round, false);
                g.SmoothingMode = SmoothingMode.None;
                g.PixelOffsetMode = PixelOffsetMode.Default;
                FillUnits(g, red, R(0, 7, 9, 2), u);
                FillUnits(g, red, R(7, 6, 4, 4), u);
            }
            return bmp;
        }

        /// <summary>A thick line with the given cap at both ends, for the cap list.</summary>
        public static Bitmap CapPreview(int logicalWidth, int logicalHeight, CapKind cap)
        {
            int w = (int)Math.Round(logicalWidth * Scale), h = (int)Math.Round(logicalHeight * Scale);
            var bmp = new Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            using (var red = new SolidBrush(AlignRed))
            using (var pale = new SolidBrush(CapStroke))
            {
                int half = Math.Max(3, (int)Math.Round(5 * Scale)), t = 2 * half;      // an even thickness: half of it is whole pixels
                int y = (h - t) / 2;
                int x0 = (int)Math.Round(24 * Scale), x1 = w - x0;                     // where the line itself starts and ends
                FillCapShape(g, pale, x0, x1, y, t, cap, true);

                g.SmoothingMode = SmoothingMode.None;
                g.PixelOffsetMode = PixelOffsetMode.Default;
                int core = 2 * Math.Max(1, (int)Math.Round(Scale)), dot = 2 * core;    // the line itself, and its end points
                g.FillRectangle(red, x0, y + half - core / 2, x1 - x0, core);
                g.FillRectangle(red, x0 - dot / 2, y + half - dot / 2, dot, dot);
                g.FillRectangle(red, x1 - dot / 2, y + half - dot / 2, dot, dot);
            }
            return bmp;
        }

        /// <summary>A stroke from x0 to x1, <paramref name="t"/> thick, with the cap at its right end and, if wanted, at its left end.</summary>
        private static void FillCapShape(Graphics g, Brush b, int x0, int x1, int y, int t, CapKind cap, bool bothEnds)
        {
            int half = t / 2;
            if (cap != CapKind.Round)
            {
                int extra = cap == CapKind.Square ? half : 0;
                g.SmoothingMode = SmoothingMode.None;
                g.FillRectangle(b, x0 - (bothEnds ? extra : 0), y, x1 - x0 + (bothEnds ? 2 : 1) * extra, t);
                return;
            }
            // one outline, so that the pale tone is not laid on twice where the parts would overlap
            using (var path = new GraphicsPath())
            {
                if (bothEnds) path.AddArc(x0 - half, y, t, t, 90, 180);
                else path.AddLine(x0, y + t, x0, y);
                path.AddArc(x1 - half, y, t, t, -90, 180);
                path.CloseFigure();
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.Half;
                g.FillPath(b, path);
            }
        }

        /// <summary>Glyph of "matrix layout": three by three squares.</summary>
        public static Bitmap MatrixLayoutGlyph(int logical)
        {
            var cells = new float[9][];
            for (int r = 0; r < 3; r++)
                for (int c = 0; c < 3; c++) cells[r * 3 + c] = R(c * 6, r * 6, 4, 4);
            return Glyph(logical, cells);
        }

        /// <summary>Glyph of "matrix cut": one square with a dashed cut across and one down.</summary>
        public static Bitmap MatrixCutGlyph(int logical)
        {
            int px = (int)Math.Round(logical * Scale);
            float u = px / 16f;
            var bmp = new Bitmap(px, px, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            using (var red = new SolidBrush(AlignRed))
            {
                g.SmoothingMode = SmoothingMode.None;
                OutlineUnits(g, red, R(1, 1, 14, 14), u);
                foreach (float at in new[] { 3.5f, 7f, 10.5f })
                {
                    FillUnits(g, red, R(7.5f, at, 1, 2), u);       // the cut running down
                    FillUnits(g, red, R(at, 7.5f, 2, 1), u);       // the cut running across
                }
            }
            return bmp;
        }

        /// <summary>One cell of a colour grid.</summary>
        public static Bitmap ColorCell(Color color, int logical)
        {
            int px = (int)Math.Round(logical * Scale);
            var bmp = new Bitmap(px, px, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                using (var b = new SolidBrush(Color.FromArgb(255, color))) g.FillRectangle(b, 0, 0, px, px);
                using (var pen = new Pen(Color.FromArgb(70, 0, 0, 0), 1f)) g.DrawRectangle(pen, 0, 0, px - 1, px - 1);
            }
            return bmp;
        }

        /// <summary>"No colour": an empty cell struck through in red.</summary>
        public static Bitmap NoColor(int logical)
        {
            int px = (int)Math.Round(logical * Scale);
            var bmp = new Bitmap(px, px, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                int m = Math.Max(1, (int)Math.Round(1.5f * Scale));
                var box = new Rectangle(m, m, px - 2 * m - 1, px - 2 * m - 1);
                g.FillRectangle(Brushes.White, box);
                using (var pen = new Pen(Color.FromArgb(120, 120, 120), 1f)) g.DrawRectangle(pen, box);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (var pen = new Pen(AlignRed, Math.Max(1f, 1.3f * Scale))) g.DrawLine(pen, box.Left + 1, box.Bottom - 1, box.Right - 1, box.Top + 1);
            }
            return bmp;
        }

        /// <summary>A horizontal line of the given weight (in points) and dash style, for the weight and dash lists.</summary>
        public static Bitmap LinePreview(int logicalWidth, int logicalHeight, float weightPoints, DashKind dash)
        {
            int w = (int)Math.Round(logicalWidth * Scale), h = (int)Math.Round(logicalHeight * Scale);
            var bmp = new Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                // whole pixels, so the lines are crisp; weights under one pixel are drawn fainter instead of thinner
                float exact = weightPoints * 96f / 72f * Scale;
                int t = Math.Max(1, (int)Math.Round(exact));
                int alpha = exact >= 1f ? 255 : (int)Math.Round(255 * Math.Max(0.35f, exact));
                int pad = (int)Math.Round(3 * Scale), y = (h - t) / 2, x = pad, right = w - pad;
                using (var b = new SolidBrush(Color.FromArgb(alpha, AlignRed)))
                {
                    float[] pattern = DashPattern(dash);
                    if (pattern == null) g.FillRectangle(b, x, y, right - x, t);
                    else
                    {
                        if (dash == DashKind.RoundDot)
                        {
                            g.SmoothingMode = SmoothingMode.AntiAlias;
                            g.PixelOffsetMode = PixelOffsetMode.Half;
                        }
                        for (int i = 0; x < right; i++)
                        {
                            int len = Math.Max(1, (int)Math.Round(pattern[i % pattern.Length] * t));
                            if (i % 2 == 0)      // even entries are dashes, odd ones gaps
                            {
                                if (dash == DashKind.RoundDot) g.FillEllipse(b, x, y, len, t);
                                else g.FillRectangle(b, x, y, Math.Min(len, right - x), t);
                            }
                            x += len;
                        }
                    }
                }
            }
            return bmp;
        }

        /// <summary>Dash and gap lengths in multiples of the line weight (Office's proportions); null = solid.</summary>
        private static float[] DashPattern(DashKind dash)
        {
            switch (dash)
            {
                case DashKind.RoundDot: return new[] { 1f, 1f };
                case DashKind.SquareDot: return new[] { 1f, 1f };
                case DashKind.Dash: return new[] { 4f, 3f };
                case DashKind.DashDot: return new[] { 4f, 3f, 1f, 3f };
                case DashKind.LongDash: return new[] { 8f, 3f };
                case DashKind.LongDashDot: return new[] { 8f, 3f, 1f, 3f };
                case DashKind.LongDashDotDot: return new[] { 8f, 3f, 1f, 3f, 1f, 3f };
                default: return null;
            }
        }

        private static PointF P(float x, float y, float u) { return new PointF(x * u, y * u); }

        private static void FillUnits(Graphics g, Brush b, float[] r, float unit)
        {
            int x0, x1, y0, y1;
            SnapSpan(r[0], r[2], unit, out x0, out x1);
            SnapSpan(r[1], r[3], unit, out y0, out y1);
            g.FillRectangle(b, x0, y0, x1 - x0, y1 - y0);
        }

        /// <summary>A rectangle's frame with whole-pixel edges of one guide-line thickness, so the corners meet cleanly.</summary>
        private static void OutlineUnits(Graphics g, Brush b, float[] r, float unit)
        {
            int x0, x1, y0, y1;
            SnapSpan(r[0], r[2], unit, out x0, out x1);
            SnapSpan(r[1], r[3], unit, out y0, out y1);
            int t = Math.Max(1, (int)Math.Floor(unit + 0.25f));
            g.FillRectangle(b, x0, y0, x1 - x0, t);
            g.FillRectangle(b, x0, y1 - t, x1 - x0, t);
            g.FillRectangle(b, x0, y0, t, y1 - y0);
            g.FillRectangle(b, x1 - t, y0, t, y1 - y0);
        }

        private static float[] R(float x, float y, float w, float h) { return new[] { x, y, w, h }; }

        /// <summary>
        /// Filled rectangles {x, y, width, height} on a 16 × 16 grid, snapped to whole device pixels
        /// so that the glyphs stay sharp at every display scaling.
        /// </summary>
        private static Bitmap Glyph(int logical, params float[][] rects)
        {
            int px = (int)Math.Round(logical * Scale);
            float unit = px / 16f;
            var bmp = new Bitmap(px, px, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            using (var b = new SolidBrush(AlignRed))
            {
                g.SmoothingMode = SmoothingMode.None;
                foreach (float[] r in rects) FillUnits(g, b, r, unit);
            }
            return bmp;
        }

        /// <summary>
        /// Device pixels covered by one grid span. One-unit guide lines get the same whole-pixel
        /// thickness everywhere (at 150 % some would be 1 pixel and some 2 otherwise).
        /// </summary>
        private static void SnapSpan(float pos, float size, float unit, out int start, out int end)
        {
            if (size <= 1f)
            {
                int t = Math.Max(1, (int)Math.Floor(unit + 0.25f));
                start = Snap((pos + size / 2) * unit - t / 2f);
                end = start + t;
            }
            else
            {
                start = Snap(pos * unit);
                end = Math.Max(start + 1, Snap((pos + size) * unit));
            }
        }

        private static int Snap(float v) { return (int)Math.Floor(v + 0.5f); }

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
