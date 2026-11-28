// Drives "matrix layout" and "matrix cut" against a real PowerPoint instance. The layout is checked cell by
// cell; a cut is checked by drawing: the pieces together must look like the object did before.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptSelectionExport
{
    internal static class MatrixTests
    {
        private static readonly Office.MsoTriState T = Office.MsoTriState.msoTrue, F = Office.MsoTriState.msoFalse;
        private static int passed, failed;
        private static string dir;

        public static void Run(PowerPoint.Application app)
        {
            dir = Path.Combine(Path.GetTempPath(), "PptSelectionExport", "matrixtests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            PowerPoint.Presentation pres = app.Presentations.Add(F);
            try
            {
                Layout(pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank));
                Cut(pres);
                Console.WriteLine("matrix tests: {0} passed, {1} failed", passed, failed);
            }
            finally
            {
                pres.Saved = T;
                pres.Close();
                try { Directory.Delete(dir, true); } catch { }
            }
        }

        // ------------------------------------------------------------------ layout

        private static void Layout(PowerPoint.Slide slide)
        {
            PowerPoint.Shapes sh = slide.Shapes;
            // six shapes of different sizes, scattered; f is turned by 90 degrees, so it covers 40 x 100
            float[][] frames =
            {
                new float[] { 300, 220, 100, 50 }, new float[] { 60, 300, 60, 80 }, new float[] { 500, 40, 120, 40 },
                new float[] { 200, 60, 50, 50 }, new float[] { 420, 330, 80, 30 }, new float[] { 90, 120, 100, 40 }
            };
            string[] names = { "a", "b", "c", "d", "e", "f" };
            for (int i = 0; i < names.Length; i++)
                sh.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, frames[i][0], frames[i][1], frames[i][2], frames[i][3]).Name = names[i];
            sh["f"].Rotation = 90;
            PowerPoint.Shape link = sh.AddConnector(Office.MsoConnectorType.msoConnectorStraight, 0, 0, 10, 10); link.Name = "link";
            link.ConnectorFormat.BeginConnect(sh["a"], 1); link.ConnectorFormat.EndConnect(sh["b"], 1);
            PowerPoint.Shape other = sh.AddShape(Office.MsoAutoShapeType.msoShapeOval, 700, 400, 40, 40); other.Name = "other";

            var before = new Dictionary<string, float[]>();
            foreach (string n in names) before[n] = Visual(sh[n]);
            double left = double.MaxValue, top = double.MaxValue;
            foreach (float[] b in before.Values) { left = Math.Min(left, b[0]); top = Math.Min(top, b[1]); }

            var layout = new GridLayout(sh.Range(new object[] { "a", "b", "c", "link", "d", "e", "f" }));
            Check("layout: glued connector is not a cell", layout.Count == 6, layout.Count + " shapes");
            Check("rows for columns", GridLayout.RowsFor(6, 3) == 2 && GridLayout.RowsFor(7, 3) == 3 && GridLayout.RowsFor(6, 1) == 6 && GridLayout.RowsFor(6, 9) == 1, "");

            // ---- 3 columns, gaps 10 and 20, in selection order: a b c / d e f
            int unmoved = layout.Apply(new GridOptions { Columns = 3, GapX = 10, GapY = 20 });
            string diff = Compare(sh, before, new[] { "a", "b", "c", "d", "e", "f" }, 3, left, top, 10, 20);
            Check("3 columns, selection order", unmoved == 0 && diff == "", diff);
            Check("the grid starts where the shapes were", Near(Visual(sh["a"])[0], left) && Near(Math.Min(Visual(sh["a"])[1], Math.Min(Visual(sh["b"])[1], Visual(sh["c"])[1])), top), "");
            Check("glued connector still glued, other shape untouched",
                  link.ConnectorFormat.BeginConnected == T && link.ConnectorFormat.EndConnected == T && Near(other.Left, 700) && Near(other.Top, 400), "");

            // ---- other values on the same object: computed from where the shapes were, not from the last layout
            layout.Apply(new GridOptions { Columns = 2, GapX = 0, GapY = 5 });
            diff = Compare(sh, before, new[] { "a", "b", "c", "d", "e", "f" }, 2, left, top, 0, 5);
            Check("2 columns after 3: as if applied first", diff == "", diff);
            layout.Apply(new GridOptions { Columns = 1, GapX = 30, GapY = -4 });
            diff = Compare(sh, before, new[] { "a", "b", "c", "d", "e", "f" }, 1, left, top, 30, -4);
            Check("1 column, negative gap", diff == "", diff);
            layout.Apply(new GridOptions { Columns = 99, GapX = 8, GapY = 8 });
            diff = Compare(sh, before, new[] { "a", "b", "c", "d", "e", "f" }, 6, left, top, 8, 8);
            Check("more columns than shapes: one row", diff == "", diff);

            // ---- by position: the three topmost shapes (c, d, f) are the first row, each row from left to right
            layout.Apply(new GridOptions { Columns = 3, GapX = 10, GapY = 20, ByPosition = true });
            diff = Compare(sh, before, new[] { "f", "d", "c", "b", "a", "e" }, 3, left, top, 10, 20);
            Check("3 columns, by position", diff == "", diff);

            layout.Restore();
            bool back = true;
            foreach (string n in names) { float[] now = Visual(sh[n]); for (int k = 0; k < 4; k++) if (!Near(now[k], before[n][k])) back = false; }
            Check("restore: everything where it was", back && link.ConnectorFormat.BeginConnected == T, "");
        }

        /// <summary>What a shape covers (left, top, width, height), a turned shape with its turned box.</summary>
        private static float[] Visual(PowerPoint.Shape s)
        {
            double w = s.Width, h = s.Height, cx = s.Left + w / 2, cy = s.Top + h / 2, a = s.Rotation * Math.PI / 180;
            double cos = Math.Abs(Math.Cos(a)), sin = Math.Abs(Math.Sin(a)), bw = w * cos + h * sin, bh = w * sin + h * cos;
            return new[] { (float)(cx - bw / 2), (float)(cy - bh / 2), (float)bw, (float)bh };
        }

        /// <summary>What differs from: the shapes, in this order, in a grid of this many columns, each centred in its cell.</summary>
        private static string Compare(PowerPoint.Shapes sh, Dictionary<string, float[]> before, string[] order, int columns, double left, double top, double gapX, double gapY)
        {
            int n = order.Length, rows = (n + columns - 1) / columns;
            var width = new double[columns];
            var height = new double[rows];
            for (int i = 0; i < n; i++)
            {
                width[i % columns] = Math.Max(width[i % columns], before[order[i]][2]);
                height[i / columns] = Math.Max(height[i / columns], before[order[i]][3]);
            }
            string diff = "";
            for (int i = 0; i < n; i++)
            {
                double x = left, y = top;
                for (int c = 0; c < i % columns; c++) x += width[c] + gapX;
                for (int r = 0; r < i / columns; r++) y += height[r] + gapY;
                double cx = x + width[i % columns] / 2, cy = y + height[i / columns] / 2;
                float[] now = Visual(sh[order[i]]);
                if (!Near(now[0] + now[2] / 2, cx) || !Near(now[1] + now[3] / 2, cy) || !Near(now[2], before[order[i]][2]) || !Near(now[3], before[order[i]][3]))
                    diff += string.Format("{0}: centre {1:0.##},{2:0.##}, expected {3:0.##},{4:0.##}  ", order[i], now[0] + now[2] / 2, now[1] + now[3] / 2, cx, cy);
            }
            return diff;
        }

        // ------------------------------------------------------------------ cut

        private static void Cut(PowerPoint.Presentation pres)
        {
            // a picture of four quarters: red | green over blue | yellow, with a mark that shows which way up it is
            string quarters = Path.Combine(dir, "quarters.png");
            using (var bmp = new Bitmap(600, 400))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                using (var font = new Font("Arial", 60, FontStyle.Bold, GraphicsUnit.Pixel))
                {
                    g.FillRectangle(Brushes.Red, 0, 0, 300, 200); g.FillRectangle(Brushes.Lime, 300, 0, 300, 200);
                    g.FillRectangle(Brushes.Blue, 0, 200, 300, 200); g.FillRectangle(Brushes.Yellow, 300, 200, 300, 200);
                    g.DrawString("A1", font, Brushes.White, 40, 50); g.FillEllipse(Brushes.Black, 470, 260, 80, 80);
                }
                bmp.Save(quarters, ImageFormat.Png);
            }

            // ---- a plain picture, 2 x 2, between a shape below it and one above it
            PowerPoint.Slide s = pres.Slides.Add(pres.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            PowerPoint.Shapes sh = s.Shapes;
            PowerPoint.Shape below = sh.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 80, 80, 100, 100); below.Name = "below";
            PowerPoint.Shape pic = sh.AddPicture(quarters, F, T, 100, 100, 300, 200); pic.Name = "pic";
            PowerPoint.Shape above = sh.AddShape(Office.MsoAutoShapeType.msoShapeOval, 350, 250, 100, 100); above.Name = "above";
            string before = Render(sh.Range(new object[] { "pic" }), "plain_before");
            List<PowerPoint.Shape> pieces;
            CutResult result = MatrixCut.Cut(sh, pic, 2, 2, out pieces);
            string colours = "";
            bool frames = pieces.Count == 4;
            for (int i = 0; i < pieces.Count; i++)
            {
                colours += ColourOf(pieces[i], "plain_piece" + i) + " ";
                int r = i / 2, c = i % 2;
                float[] f = FrameByName(sh, "pic " + (r + 1) + "-" + (c + 1));
                if (f == null || !Near(f[0], 100 + 150 * c) || !Near(f[1], 100 + 100 * r) || !Near(f[2], 150) || !Near(f[3], 100)) frames = false;
            }
            bool allPictures = true, ordered = true;
            foreach (PowerPoint.Shape p in pieces)
            {
                if (p.Type != Office.MsoShapeType.msoPicture) allPictures = false;
                if (p.ZOrderPosition <= below.ZOrderPosition || p.ZOrderPosition >= above.ZOrderPosition) ordered = false;
            }
            Check("picture 2 x 2: four pictures in their cells", result == CutResult.Done && frames && allPictures && colours.Trim() == "red green blue yellow", pieces.Count + " pieces: " + colours);
            Check("the original is gone, the pieces took its place", FrameByName(sh, "pic") == null && sh.Count == 6 && ordered, sh.Count + " shapes on the slide");
            Check("the pieces together look like the picture", Same(before, Render(RangeOf(sh, pieces), "plain_after"), 1.0), LastDifference);

            // ---- turned, mirrored, cropped by the user before: always the same picture afterwards
            CutAndCompare(pres, quarters, "picture turned by 30 degrees, 3 x 2", 3, 2, delegate(PowerPoint.Shape p) { p.Rotation = 30; }, 6, 3.0);
            CutAndCompare(pres, quarters, "picture mirrored, 2 x 2", 2, 2, delegate(PowerPoint.Shape p) { p.Flip(Office.MsoFlipCmd.msoFlipHorizontal); }, 4, 1.0);
            CutAndCompare(pres, quarters, "picture mirrored and turned, 2 x 3", 2, 3, delegate(PowerPoint.Shape p) { p.Flip(Office.MsoFlipCmd.msoFlipVertical); p.Rotation = 300; }, 6, 3.0);
            CutAndCompare(pres, quarters, "picture the user had cropped, 4 x 2", 4, 2, delegate(PowerPoint.Shape p)
            {
                p.PictureFormat.Crop.ShapeWidth = 240; p.PictureFormat.Crop.ShapeHeight = 150;
            }, 8, 1.0);

            // ---- a picture that sits in a placeholder of the layout
            try
            {
                PowerPoint.Slide held = pres.Slides.Add(pres.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutObject);
                int had = held.Shapes.Count;
                PowerPoint.Shape inside = held.Shapes.AddPicture(quarters, F, T, 0, 0);
                bool placeholder = inside.Type == Office.MsoShapeType.msoPlaceholder;
                inside.Name = "held";
                before = Render(held.Shapes.Range(new object[] { "held" }), "held_before");
                result = MatrixCut.Cut(held.Shapes, inside, 2, 2, out pieces);
                bool same = pieces.Count == 4 && Same(before, Render(RangeOf(held.Shapes, pieces), "held_after"), 1.0);
                Check("picture in a placeholder, 2 x 2", placeholder && result == CutResult.Done && same,
                      (placeholder ? "" : "the picture did not go into the placeholder; ") + pieces.Count + " pieces; " + held.Shapes.Count + " shapes, were " + had + "; " + LastDifference);
            }
            catch (Exception ex) { Check("picture in a placeholder, 2 x 2", false, ex.Message); }

            // ---- a shape: cut with Intersect
            s = pres.Slides.Add(pres.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            sh = s.Shapes;
            PowerPoint.Shape oval = sh.AddShape(Office.MsoAutoShapeType.msoShapeOval, 100, 100, 320, 200); oval.Name = "oval";
            oval.Fill.ForeColor.RGB = 0x2060E0; oval.Line.Visible = F; oval.TextFrame.TextRange.Text = "Label";
            bool hadText = MatrixCut.HasText(oval);
            oval.TextFrame.TextRange.Text = "";                 // for the picture of "before": the text is not part of what is cut
            before = Render(sh.Range(new object[] { "oval" }), "oval_before");
            oval.TextFrame.TextRange.Text = "Label";
            MatrixCut.Cut(sh, oval, 4, 4, out pieces);
            bool noText = true, filled = true;
            foreach (PowerPoint.Shape p in pieces)
            {
                if (p.HasTextFrame == T && p.TextFrame.HasText == T) noText = false;
                if (p.Fill.ForeColor.RGB != 0x2060E0) filled = false;
            }
            Check("oval 4 x 4: 16 pieces with its fill, text left out", pieces.Count == 16 && hadText && noText && filled && FrameByName(sh, "oval") == null, pieces.Count + " pieces");
            Check("the pieces together look like the oval", Same(before, Render(RangeOf(sh, pieces), "oval_after"), 2.0), LastDifference);

            PowerPoint.Shape diamond = sh.AddShape(Office.MsoAutoShapeType.msoShapeDiamond, 480, 100, 200, 160); diamond.Name = "diamond";
            MatrixCut.Cut(sh, diamond, 4, 4, out pieces);
            Check("diamond 4 x 4: its four empty corners give no piece", pieces.Count == 12 && FrameByName(sh, "diamond") == null, pieces.Count + " pieces");

            // ---- what cannot be cut is left alone
            PowerPoint.Shape line = sh.AddLine(100, 400, 300, 420);
            PowerPoint.Shape table = sh.AddTable(2, 2, 400, 380, 200, 80);
            sh.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 650, 380, 40, 40).Name = "g1";
            sh.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 700, 380, 40, 40).Name = "g2";
            PowerPoint.Shape group = sh.Range(new object[] { "g1", "g2" }).Group();
            int count = sh.Count;
            bool refused = true;
            foreach (PowerPoint.Shape no in new[] { line, table, group })
                if (MatrixCut.CanCut(no) || MatrixCut.Cut(sh, no, 2, 2, out pieces) != CutResult.Unsupported || pieces.Count != 0) refused = false;
            Check("line, table, group: refused, nothing changes", refused && sh.Count == count, sh.Count + " shapes, were " + count);
        }

        /// <summary>A picture prepared by <paramref name="prepare"/>, cut, and compared with how it looked before.</summary>
        private static void CutAndCompare(PowerPoint.Presentation pres, string file, string name, int columns, int rows, Action<PowerPoint.Shape> prepare, int expected, double tolerance)
        {
            try
            {
                PowerPoint.Shapes sh = pres.Slides.Add(pres.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank).Shapes;
                PowerPoint.Shape pic = sh.AddPicture(file, F, T, 200, 150, 300, 200); pic.Name = "pic";
                prepare(pic);
                string tag = name.Substring(0, name.IndexOf(',')).Replace(' ', '_');
                string before = Render(sh.Range(new object[] { "pic" }), tag + "_before");
                List<PowerPoint.Shape> pieces;
                MatrixCut.Cut(sh, pic, columns, rows, out pieces);
                bool same = Same(before, Render(RangeOf(sh, pieces), tag + "_after"), tolerance);
                Check(name, pieces.Count == expected && sh.Count == expected && same, pieces.Count + " pieces; " + LastDifference);
            }
            catch (Exception ex) { Check(name, false, ex.Message); }
        }

        private static PowerPoint.ShapeRange RangeOf(PowerPoint.Shapes sh, List<PowerPoint.Shape> pieces)
        {
            var indices = new object[pieces.Count];
            for (int i = 0; i < pieces.Count; i++) indices[i] = pieces[i].ZOrderPosition;
            return sh.Range(indices);
        }

        private static float[] FrameByName(PowerPoint.Shapes sh, string name)
        {
            for (int i = 1; i <= sh.Count; i++)
                if (sh[i].Name == name) return new[] { sh[i].Left, sh[i].Top, sh[i].Width, sh[i].Height };
            return null;
        }

        private static string Render(PowerPoint.ShapeRange range, string name)
        {
            string file = Path.Combine(dir, name + ".png");
            range.Export(file, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
            return file;
        }

        private static string ColourOf(PowerPoint.Shape shape, string name)
        {
            string file = Path.Combine(dir, name + ".png");
            shape.Export(file, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
            using (var bmp = new Bitmap(file))
            {
                Color c = bmp.GetPixel(bmp.Width / 8, bmp.Height * 7 / 8);       // a corner region free of the marks
                if (c.R > 200 && c.G < 60 && c.B < 60) return "red";
                if (c.G > 200 && c.R < 60 && c.B < 60) return "green";
                if (c.B > 200 && c.R < 60 && c.G < 60) return "blue";
                if (c.R > 200 && c.G > 200 && c.B < 60) return "yellow";
                return "(" + c.R + "," + c.G + "," + c.B + ")";
            }
        }

        private static string LastDifference = "";

        /// <summary>
        /// Whether two pictures are the same to within <paramref name="tolerance"/> (mean difference per
        /// colour channel, of 255). They may differ by a pixel in size: the bounds are rounded.
        /// </summary>
        private static bool Same(string fileA, string fileB, double tolerance)
        {
            using (var a = new Bitmap(fileA))
            using (var b = new Bitmap(fileB))
            {
                if (Math.Abs(a.Width - b.Width) > 2 || Math.Abs(a.Height - b.Height) > 2)
                {
                    LastDifference = "sizes " + a.Width + "x" + a.Height + " and " + b.Width + "x" + b.Height;
                    return false;
                }
                int w = Math.Min(a.Width, b.Width), h = Math.Min(a.Height, b.Height);
                var area = new Rectangle(0, 0, w, h);
                BitmapData da = a.LockBits(area, ImageLockMode.ReadOnly, PixelFormat.Format32bppPArgb), db = b.LockBits(area, ImageLockMode.ReadOnly, PixelFormat.Format32bppPArgb);
                long total = 0;
                try
                {
                    byte[] ra = new byte[w * 4], rb = new byte[w * 4];
                    for (int y = 0; y < h; y++)
                    {
                        System.Runtime.InteropServices.Marshal.Copy(new IntPtr(da.Scan0.ToInt64() + (long)y * da.Stride), ra, 0, ra.Length);
                        System.Runtime.InteropServices.Marshal.Copy(new IntPtr(db.Scan0.ToInt64() + (long)y * db.Stride), rb, 0, rb.Length);
                        for (int i = 0; i < ra.Length; i++) total += Math.Abs(ra[i] - rb[i]);
                    }
                }
                finally { a.UnlockBits(da); b.UnlockBits(db); }
                double mean = total / (4.0 * w * h);
                LastDifference = string.Format("mean difference {0:0.00} of 255 ({1}x{2} px)", mean, w, h);
                return mean <= tolerance;
            }
        }

        private static bool Near(double a, double b) { return Math.Abs(a - b) < 0.06; }

        private static void Check(string name, bool ok, string detail)
        {
            if (ok) passed++; else failed++;
            Console.WriteLine("{0,-48} {1,-8} {2}", name, ok ? "OK" : "MISMATCH", detail);
        }
    }
}
