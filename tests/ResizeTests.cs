// Drives "slide size" against a real PowerPoint instance: for several new sizes, every object must end up
// scaled by one factor (the smaller of the two ratios) about the middle of the slide, with its text sizes,
// while masters and layouts follow the slide and line weights stay as they are.
using System;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptSelectionExport
{
    internal static class ResizeTests
    {
        private static readonly Office.MsoTriState T = Office.MsoTriState.msoTrue, F = Office.MsoTriState.msoFalse;
        private const float W0 = 960, H0 = 540;
        private static int passed, failed;

        /// <summary>Shapes that are scaled about the middle of the slide (the last one sizes its height to its text).</summary>
        private static readonly string[] Scaled = { "R", "T", "A", "B", "C", "FreeLine", "Locked", "MovedTitle", "Default" };
        /// <summary>Shapes that follow the slide: stretched with it in each direction.</summary>
        private static readonly string[] Stretched = { "TitleInPlace", "LayoutTitle", "LayoutLogo", "MasterTitle", "MasterLogo" };
        private static readonly string[] ScaledValues =
        {
            "font.R", "margin.R.left", "margin.R.top", "before.R", "within.R", "indent.R.left", "indent.R.first", "font.Default",
            "col.T", "row.T", "font.cell", "margin.cell", "font.MovedTitle", "font.TitleInPlace", "font.LayoutTitle", "font.MasterTitle", "font.MasterBody"
        };
        private static readonly string[] KeptValues = { "line.R", "glow.R", "border.cell", "rotation.Locked", "glued.C" };

        private sealed class Deck
        {
            public PowerPoint.Presentation Pres;
            public PowerPoint.Slide S1, S2;
            public PowerPoint.CustomLayout Layout;
            public Dictionary<string, float[]> Frames;
            public Dictionary<string, float> Values;
        }

        public static void Run(PowerPoint.Application app)
        {
            Check("fit factor: the smaller ratio", Near(SlideResizer.FitScale(960, 540, 480, 270), 0.5) && Near(SlideResizer.FitScale(960, 540, 1920, 540), 1)
                  && Near(SlideResizer.FitScale(960, 540, 480, 1080), 0.5) && Near(SlideResizer.FitScale(960, 540, 1440, 1080), 1.5)
                  && Near(SlideResizer.ToCm(SlideResizer.ToPoints(16)), 16) && Near(SlideResizer.ToPoints(2.54), 72), "");

            Case(app, "same proportions, half the size", 480, 270);
            Case(app, "same proportions, one and a half times", 1440, 810);
            Case(app, "16 cm x 9 cm", SlideResizer.ToPoints(16), SlideResizer.ToPoints(9));
            Case(app, "narrower only", 700, 540);
            Case(app, "smaller in both, by different ratios", 480, 405);
            Case(app, "narrower and taller", 480, 1080);
            Case(app, "wider only (nothing to scale)", 1920, 540);
            Case(app, "larger in both, by different ratios", 1440, 1080);

            // ---- there and back again: everything is where it was
            Deck d = Build(app);
            try
            {
                SlideResizer.Resize(d.Pres, SlideResizer.ToPoints(16), SlideResizer.ToPoints(9));
                float small = d.Pres.PageSetup.SlideWidth;
                double back = SlideResizer.Resize(d.Pres, W0, H0);
                string diff = Compare(d, 1, 1, 1, W0, H0, 0.15);
                Check("to 16 cm x 9 cm and back: as before", diff == "" && Near(back, W0 / small), diff);

                double same = SlideResizer.Resize(d.Pres, W0, H0);
                Check("the size it has already: nothing happens", Near(same, 1) && Compare(d, 1, 1, 1, W0, H0, 0.15) == "", "");

                bool refused = false;
                try { SlideResizer.Resize(d.Pres, 50, 540); } catch (ArgumentOutOfRangeException) { refused = true; }
                try { SlideResizer.Resize(d.Pres, 960, 5000); refused = false; } catch (ArgumentOutOfRangeException) { }
                Check("sizes PowerPoint cannot have are refused", refused && Near(d.Pres.PageSetup.SlideWidth, W0) && Near(d.Pres.PageSetup.SlideHeight, H0), "");
            }
            finally { Close(d); }

            Console.WriteLine("resize tests: {0} passed, {1} failed", passed, failed);
        }

        private static void Case(PowerPoint.Application app, string name, float width, float height)
        {
            Deck d = Build(app);
            try
            {
                var watch = System.Diagnostics.Stopwatch.StartNew();
                double scale = SlideResizer.Resize(d.Pres, width, height);
                watch.Stop();
                // PowerPoint keeps slide sizes in eighths of a point; the content goes by the size it really took
                float asked = width, askedHeight = height;
                width = d.Pres.PageSetup.SlideWidth; height = d.Pres.PageSetup.SlideHeight;
                double rw = width / W0, rh = height / H0, expected = Math.Min(rw, rh);
                string diff = Math.Abs(width - asked) <= 0.07 && Math.Abs(height - askedHeight) <= 0.07 ? "" : "slide is " + width + " x " + height + "  ";
                if (!Near(scale, expected)) diff += "factor " + scale + "  ";
                diff += Compare(d, expected, rw, rh, width, height, 0.15);
                Check(name, diff == "", diff == "" ? "content x " + expected.ToString("0.###") + ", " + watch.ElapsedMilliseconds + " ms" : diff);
            }
            catch (Exception ex) { Check(name, false, ex.Message); }
            finally { Close(d); }
        }

        /// <summary>What differs from "everything scaled by <paramref name="s"/>, masters and layouts stretched with the slide"; "" when nothing does.</summary>
        private static string Compare(Deck d, double s, double rw, double rh, float width, float height, double tolerance)
        {
            Dictionary<string, float[]> frames = Frames(d);
            Dictionary<string, float> values = Values(d);
            string diff = "";
            foreach (string name in Scaled)
            {
                float[] a = d.Frames[name], b = frames[name];
                double cx = width / 2.0 + (a[0] + a[2] / 2.0 - W0 / 2.0) * s, cy = height / 2.0 + (a[1] + a[3] / 2.0 - H0 / 2.0) * s;
                bool ownHeight = name == "Default";         // a text box that sizes itself to its text
                bool ok = Math.Abs(b[2] - a[2] * s) <= tolerance && Math.Abs(b[0] + b[2] / 2.0 - cx) <= tolerance && Math.Abs(b[1] + b[3] / 2.0 - cy) <= tolerance
                          && (ownHeight || Math.Abs(b[3] - a[3] * s) <= tolerance);
                if (!ok) diff += string.Format("{0}: {1:0.##},{2:0.##} {3:0.##}x{4:0.##}, expected centre {5:0.##},{6:0.##} size {7:0.##}x{8:0.##}  ", name, b[0], b[1], b[2], b[3], cx, cy, a[2] * s, a[3] * s);
            }
            foreach (string name in Stretched)
            {
                float[] a = d.Frames[name], b = frames[name];
                if (Math.Abs(b[0] - a[0] * rw) > tolerance || Math.Abs(b[1] - a[1] * rh) > tolerance || Math.Abs(b[2] - a[2] * rw) > tolerance || Math.Abs(b[3] - a[3] * rh) > tolerance)
                    diff += string.Format("{0}: {1:0.##},{2:0.##} {3:0.##}x{4:0.##}, expected {5:0.##},{6:0.##} {7:0.##}x{8:0.##}  ", name, b[0], b[1], b[2], b[3], a[0] * rw, a[1] * rh, a[2] * rw, a[3] * rh);
            }
            foreach (string name in ScaledValues)
                if (Math.Abs(values[name] - d.Values[name] * s) > 0.02 + 0.002 * Math.Abs(d.Values[name] * s))
                    diff += string.Format("{0}: {1:0.###}, expected {2:0.###}  ", name, values[name], d.Values[name] * s);
            foreach (string name in KeptValues)
                if (Math.Abs(values[name] - d.Values[name]) > 0.02)
                    diff += string.Format("{0}: {1:0.###}, was {2:0.###}  ", name, values[name], d.Values[name]);
            return diff;
        }

        // ------------------------------------------------------------------ the deck

        private static Deck Build(PowerPoint.Application app)
        {
            var d = new Deck { Pres = app.Presentations.Add(F) };
            d.Pres.PageSetup.SlideWidth = W0;
            d.Pres.PageSetup.SlideHeight = H0;
            d.Layout = d.Pres.SlideMaster.CustomLayouts[6];         // "Title Only" in the default template
            d.Pres.SlideMaster.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 20, 20, 100, 40).Name = "MasterLogo";
            d.Layout.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 840, 480, 100, 40).Name = "LayoutLogo";
            d.S1 = d.Pres.Slides.AddSlide(1, d.Layout);
            d.S2 = d.Pres.Slides.AddSlide(2, d.Layout);
            PowerPoint.Shape moved = d.S1.Shapes[1];
            moved.TextFrame.TextRange.Text = "Moved title";
            moved.Left = 200; moved.Top = 300; moved.Width = 400;
            d.S2.Shapes[1].TextFrame.TextRange.Text = "Title in place";

            PowerPoint.Shapes sh = d.S1.Shapes;
            PowerPoint.Shape r = sh.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 96, 108, 288, 108); r.Name = "R";
            r.TextFrame.TextRange.Text = "first\rsecond";
            r.TextFrame.TextRange.Font.Size = 20.5f;
            r.TextFrame.MarginLeft = 20; r.TextFrame.MarginTop = 10;
            Office.ParagraphFormat2 pf = r.TextFrame2.TextRange.ParagraphFormat;
            pf.SpaceBefore = 12; pf.LeftIndent = 30; pf.FirstLineIndent = -15; pf.LineRuleWithin = F; pf.SpaceWithin = 30;
            r.Glow.Radius = 10; r.Line.Weight = 3;
            PowerPoint.Shape plain = sh.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 500, 60, 200, 40); plain.Name = "Default";
            plain.TextFrame.TextRange.Text = "default size text";
            PowerPoint.Shape table = sh.AddTable(2, 2, 480, 270, 384, 162); table.Name = "T";
            table.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "cell";
            PowerPoint.Shape a = sh.AddShape(Office.MsoAutoShapeType.msoShapeOval, 60, 400, 80, 80); a.Name = "A";
            PowerPoint.Shape b = sh.AddShape(Office.MsoAutoShapeType.msoShapeOval, 300, 430, 80, 80); b.Name = "B";
            PowerPoint.Shape c = sh.AddConnector(Office.MsoConnectorType.msoConnectorStraight, 0, 0, 10, 10); c.Name = "C";
            c.ConnectorFormat.BeginConnect(a, 4); c.ConnectorFormat.EndConnect(b, 2);
            sh.AddLine(700, 400, 900, 500).Name = "FreeLine";
            PowerPoint.Shape locked = sh.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 720, 60, 200, 100); locked.Name = "Locked";
            locked.LockAspectRatio = T; locked.Rotation = 30;

            d.Frames = Frames(d);
            d.Values = Values(d);
            return d;
        }

        private static void Close(Deck d)
        {
            try { d.Pres.Saved = T; d.Pres.Close(); } catch { }
        }

        private static float[] Frame(PowerPoint.Shape s) { return new[] { s.Left, s.Top, s.Width, s.Height }; }

        private static Dictionary<string, float[]> Frames(Deck d)
        {
            var f = new Dictionary<string, float[]>();
            foreach (string name in new[] { "R", "T", "A", "B", "C", "FreeLine", "Locked", "Default" }) f[name] = Frame(d.S1.Shapes[name]);
            f["MovedTitle"] = Frame(d.S1.Shapes[1]);
            f["TitleInPlace"] = Frame(d.S2.Shapes[1]);
            f["LayoutTitle"] = Frame(d.Layout.Shapes[1]);
            f["LayoutLogo"] = Frame(d.Layout.Shapes["LayoutLogo"]);
            f["MasterTitle"] = Frame(d.Pres.SlideMaster.Shapes[1]);
            f["MasterLogo"] = Frame(d.Pres.SlideMaster.Shapes["MasterLogo"]);
            return f;
        }

        private static Dictionary<string, float> Values(Deck d)
        {
            var v = new Dictionary<string, float>();
            PowerPoint.Shape r = d.S1.Shapes["R"], t = d.S1.Shapes["T"], c = d.S1.Shapes["C"];
            Office.ParagraphFormat2 pf = r.TextFrame2.TextRange.ParagraphFormat;
            v["font.R"] = r.TextFrame.TextRange.Font.Size;
            v["margin.R.left"] = r.TextFrame.MarginLeft; v["margin.R.top"] = r.TextFrame.MarginTop;
            v["before.R"] = pf.SpaceBefore; v["within.R"] = pf.SpaceWithin;
            v["indent.R.left"] = pf.LeftIndent; v["indent.R.first"] = pf.FirstLineIndent;
            v["line.R"] = r.Line.Weight; v["glow.R"] = r.Glow.Radius;
            v["font.Default"] = d.S1.Shapes["Default"].TextFrame.TextRange.Font.Size;
            v["col.T"] = t.Table.Columns[1].Width; v["row.T"] = t.Table.Rows[1].Height;
            PowerPoint.Shape cell = t.Table.Cell(1, 1).Shape;
            v["font.cell"] = cell.TextFrame.TextRange.Font.Size; v["margin.cell"] = cell.TextFrame.MarginLeft;
            v["border.cell"] = t.Table.Cell(1, 1).Borders[PowerPoint.PpBorderType.ppBorderTop].Weight;
            v["font.MovedTitle"] = d.S1.Shapes[1].TextFrame.TextRange.Font.Size;
            v["font.TitleInPlace"] = d.S2.Shapes[1].TextFrame.TextRange.Font.Size;
            v["font.LayoutTitle"] = d.Layout.Shapes[1].TextFrame.TextRange.Font.Size;
            v["font.MasterTitle"] = d.Pres.SlideMaster.Shapes[1].TextFrame.TextRange.Font.Size;
            v["font.MasterBody"] = d.Pres.SlideMaster.Shapes[2].TextFrame.TextRange.Paragraphs(1, 1).Font.Size;
            v["rotation.Locked"] = d.S1.Shapes["Locked"].Rotation;
            v["glued.C"] = (c.ConnectorFormat.BeginConnected == T && c.ConnectorFormat.EndConnected == T) ? 1 : 0;
            return v;
        }

        private static bool Near(double a, double b) { return Math.Abs(a - b) < 0.011; }

        private static void Check(string name, bool ok, string detail)
        {
            if (ok) passed++; else failed++;
            Console.WriteLine("{0,-48} {1,-8} {2}", name, ok ? "OK" : "MISMATCH", detail);
        }
    }
}
