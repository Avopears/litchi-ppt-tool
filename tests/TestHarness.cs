// Drives the export and alignment code against a real PowerPoint instance, without the ribbon.
// Usage: TestHarness.exe <output folder> [align]   ("align" runs only the alignment tests)
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptSelectionExport
{
    internal static class TestHarness
    {
        private static readonly Office.MsoTriState T = Office.MsoTriState.msoTrue, F = Office.MsoTriState.msoFalse;
        private static string outDir;
        private static PowerPoint.Application app;
        private static PowerPoint.Presentation pres;
        private static PowerPoint.Slide slide;

        private static void Run(string name, PowerPoint.ShapeRange range, bool child, BackgroundOptions bg)
        {
            try
            {
                SizeD s1 = Exporter.ExportSvg(range, Path.Combine(outDir, name + ".svg"), bg);
                SizeD s2 = Exporter.ExportPdf(app, pres, slide, range, child, Path.Combine(outDir, name + ".pdf"), bg);
                Console.WriteLine("{0,-22} svg {1:0.##} x {2:0.##} pt   pdf {3:0.##} x {4:0.##} pt   {5}",
                    name, s1.W, s1.H, s2.W, s2.H,
                    Math.Abs(s1.W - s2.W) < 0.2 && Math.Abs(s1.H - s2.H) < 0.2 ? "OK" : "SIZE MISMATCH");
            }
            catch (Exception ex) { Console.WriteLine("{0,-22} FAILED: {1}", name, ex); }
        }

        private static void RunImage(string name, PowerPoint.ShapeRange range, BackgroundOptions bg, int dpi, bool jpeg)
        {
            try
            {
                float slideW = pres.PageSetup.SlideWidth, slideH = pres.PageSetup.SlideHeight;
                string reference = Path.Combine(outDir, "_ref.png");
                Size natural = Exporter.ExportImage(range, slideW, slideH, reference, bg, 96, false);
                File.Delete(reference);
                string file = Path.Combine(outDir, name + (jpeg ? ".jpg" : ".png"));
                var sw = System.Diagnostics.Stopwatch.StartNew();
                Size px = Exporter.ExportImage(range, slideW, slideH, file, bg, dpi, jpeg);
                sw.Stop();
                double ew = natural.Width * dpi / 96.0, eh = natural.Height * dpi / 96.0;
                using (var bmp = new Bitmap(file))
                {
                    Color corner = bmp.GetPixel(0, bmp.Height - 1);
                    Console.WriteLine("{0,-26} {1} x {2} px (expected {3:0} x {4:0})  dpi tag {5:0.#}  corner ARGB {6},{7},{8},{9}  {10} ms  {11}",
                        name, px.Width, px.Height, ew, eh, bmp.HorizontalResolution, corner.A, corner.R, corner.G, corner.B, sw.ElapsedMilliseconds,
                        Math.Abs(px.Width - ew) <= 3 + ew * 0.01 && Math.Abs(px.Height - eh) <= 3 + eh * 0.01 ? "OK" : "SIZE MISMATCH");
                }
            }
            catch (Exception ex) { Console.WriteLine("{0,-26} FAILED: {1}", name, ex); }
        }

        private static PowerPoint.ShapeRange R(params PowerPoint.Shape[] shapes)
        {
            var names = new string[shapes.Length];
            for (int i = 0; i < shapes.Length; i++) names[i] = shapes[i].Name;
            return slide.Shapes.Range(names);
        }

        private static void Main(string[] args)
        {
            outDir = Path.GetFullPath(args[0]);
            Directory.CreateDirectory(outDir);
            foreach (string f in Directory.GetFiles(outDir)) File.Delete(f);

            if (args.Length > 1 && args[1] == "align")
            {
                app = new PowerPoint.Application();
                try
                {
                    if (args.Length > 2 && args[2] == "slim") SlimTests.Run(app);       // "align slim": only the slimming tests
                    else if (args.Length > 2 && args[2] == "size") ResizeTests.Run(app);        // "align size": only the slide size tests
                    else { AlignTests.Run(app); ShapeToolsTests.Run(app); StyleTests.Run(app); CapTests.Run(app); ResizeTests.Run(app); SlimTests.Run(app); }
                }
                finally { if (app.Presentations.Count == 0) app.Quit(); }
                return;
            }

            string png = Path.Combine(outDir, "_picture.png");
            using (var bmp = new Bitmap(240, 160))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.SteelBlue);
                    g.FillEllipse(Brushes.Gold, 60, 30, 120, 100);
                }
                bmp.Save(png, ImageFormat.Png);
            }

            app = new PowerPoint.Application();
            pres = app.Presentations.Add(F);
            try
            {
                // a dark slide background and master artwork that must NOT appear in the exports
                pres.SlideMaster.Background.Fill.ForeColor.RGB = 0x402010;
                PowerPoint.Shape logo = pres.SlideMaster.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, 960, 30);
                logo.Fill.ForeColor.RGB = 0x00C000;

                pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);            // so the tested slide is not the first one
                slide = pres.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutBlank);
                PowerPoint.Shapes sh = slide.Shapes;

                // unselected neighbour that overlaps the selection area: must not be exported
                PowerPoint.Shape noise = sh.AddShape(Office.MsoAutoShapeType.msoShapeOval, 150, 120, 380, 220);
                noise.Fill.ForeColor.RGB = 0xFF00FF;

                PowerPoint.Shape rect = sh.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, 100, 100, 200, 120);
                rect.Fill.ForeColor.RGB = 0x3080F0; rect.Line.Weight = 6; rect.Line.ForeColor.RGB = 0x202020;
                rect.Shadow.Visible = T; rect.Shadow.OffsetX = 10; rect.Shadow.OffsetY = 10; rect.Shadow.Blur = 8;
                PowerPoint.Shape text = sh.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 250, 190, 240, 50);
                text.TextFrame.TextRange.Text = "选区导出 Test 123"; text.TextFrame.TextRange.Font.Size = 26;
                text.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF;
                PowerPoint.Shape arrow = sh.AddLine(120, 270, 470, 310);
                arrow.Line.Weight = 5; arrow.Line.ForeColor.RGB = 0x00A0FF;
                arrow.Line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadTriangle;

                PowerPoint.Shape a = sh.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 560, 120, 90, 60);
                PowerPoint.Shape b = sh.AddShape(Office.MsoAutoShapeType.msoShapeOval, 780, 250, 90, 60);
                PowerPoint.Shape conn = sh.AddConnector(Office.MsoConnectorType.msoConnectorElbow, 0, 0, 10, 10);
                conn.ConnectorFormat.BeginConnect(a, 4); conn.ConnectorFormat.EndConnect(b, 2);
                conn.Line.Weight = 3; conn.Line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadStealth;

                PowerPoint.Shape g1 = sh.AddShape(Office.MsoAutoShapeType.msoShapeDiamond, 600, 360, 80, 80);
                PowerPoint.Shape g2 = sh.AddShape(Office.MsoAutoShapeType.msoShapeHexagon, 720, 390, 90, 80);
                PowerPoint.Shape g3 = sh.AddShape(Office.MsoAutoShapeType.msoShape5pointStar, 830, 350, 70, 70);
                PowerPoint.Shape grp = slide.Shapes.Range(new string[] { g1.Name, g2.Name, g3.Name }).Group();

                PowerPoint.Shape rot = sh.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 60, 400, 180, 50);
                rot.Rotation = 30; rot.Fill.ForeColor.RGB = 0x40C040;
                PowerPoint.Shape tiny = sh.AddShape(Office.MsoAutoShapeType.msoShapeOval, 330, 420, 24, 14);
                tiny.Fill.ForeColor.RGB = 0x0000C0; tiny.Line.Visible = F;
                PowerPoint.Shape off = sh.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, -70, 470, 150, 110);
                off.Fill.ForeColor.RGB = 0xC08000;
                PowerPoint.Shape pic = sh.AddPicture(png, F, T, 380, 400, 150, 100);
                PowerPoint.Shape tbl = sh.AddTable(2, 3, 380, 30, 300, 60);
                tbl.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "表头";
                tbl.Table.Cell(2, 2).Shape.TextFrame.TextRange.Text = "cell";

                var transparent = new BackgroundOptions();
                var yellow = new BackgroundOptions { Transparent = false, Color = Color.FromArgb(0xFF, 0xE9, 0xA8) };

                Run("t1_multi_transparent", R(rect, text, arrow), false, transparent);
                Run("t2_multi_yellow", R(rect, text, arrow), false, yellow);
                Run("t3_tiny", R(tiny), false, transparent);
                Run("t4_offslide", R(off), false, yellow);
                Run("t5_connector", R(a, b, conn), false, transparent);
                Run("t6_group_child", grp.GroupItems.Range(new object[] { 1, 3 }), true, transparent);
                Run("t7_group_whole", R(grp), false, transparent);
                Run("t8_rotated", R(rot), false, transparent);
                Run("t9_picture_table", R(pic, tbl), false, yellow);
                Run("t10_text_only", R(text), false, transparent);

                RunImage("i1_png_300_transparent", R(rect, text, arrow), transparent, 300, false);
                RunImage("i2_jpg_150_transparent", R(rect, text, arrow), transparent, 150, true);
                RunImage("i3_png_96_yellow", R(rect, text, arrow), yellow, 96, false);
                RunImage("i4_jpg_300_yellow", R(pic, tbl), yellow, 300, true);
                RunImage("i5_png_200_child", grp.GroupItems.Range(new object[] { 1, 3 }), transparent, 200, false);
                RunImage("i6_png_600_tiny", R(tiny), transparent, 600, false);
                RunImage("i7_png_600_everything", R(noise, rect, text, arrow, a, b, conn, grp, rot, pic, tbl), transparent, 600, false);
                RunImage("i8_png_300_offslide", R(off), transparent, 300, false);

                Console.WriteLine("source untouched: slides={0} shapes={1} saved={2} bgFollow={3} open presentations={4}",
                    pres.Slides.Count, slide.Shapes.Count, pres.Saved, slide.FollowMasterBackground, app.Presentations.Count);

                AlignTests.Run(app);
                ShapeToolsTests.Run(app);
                StyleTests.Run(app);
                CapTests.Run(app);
                ResizeTests.Run(app);
                SlimTests.Run(app);
            }
            finally
            {
                pres.Saved = T;
                pres.Close();
                if (app.Presentations.Count == 0) app.Quit();
                try { File.Delete(png); } catch { }
            }
        }
    }
}
