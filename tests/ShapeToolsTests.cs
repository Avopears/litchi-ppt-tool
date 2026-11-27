// Drives rotate / flip and the boolean operations against a real PowerPoint instance,
// and checks the colour picker's colour maths.
using System;
using System.Drawing;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptSelectionExport
{
    internal static class ShapeToolsTests
    {
        private static readonly Office.MsoTriState T = Office.MsoTriState.msoTrue, F = Office.MsoTriState.msoFalse;
        private static PowerPoint.Slide slide;
        private static int passed, failed;

        public static void Run(PowerPoint.Application app)
        {
            PowerPoint.Presentation pres = app.Presentations.Add(F);
            try
            {
                slide = pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                RotateAndFlip();
                foreach (MergeCommand cmd in (MergeCommand[])Enum.GetValues(typeof(MergeCommand))) MergeTwoRectangles(cmd);
                MergeRefusals();
                ColourMaths();
                Console.WriteLine("shape tool tests: {0} passed, {1} failed", passed, failed);
            }
            finally
            {
                pres.Saved = T;
                pres.Close();
            }
        }

        private static void RotateAndFlip()
        {
            PowerPoint.Shape a = Add(Office.MsoAutoShapeType.msoShapeRightTriangle, 100, 100, 120, 60, "ra");
            PowerPoint.Shape b = Add(Office.MsoAutoShapeType.msoShapeRightTriangle, 400, 200, 80, 100, "rb");
            b.Rotation = 30;
            PowerPoint.ShapeRange r = slide.Shapes.Range(new object[] { "ra", "rb" });

            ShapeTools.Transform(r, TransformCommand.RotateRight);
            Check("rotate right 90", Near(a.Rotation, 90) && Near(b.Rotation, 120) && Near(a.Left, 100) && Near(b.Left, 400),
                  a.Rotation + " / " + b.Rotation);
            ShapeTools.Transform(r, TransformCommand.RotateLeft);
            ShapeTools.Transform(r, TransformCommand.RotateLeft);
            Check("rotate left 90 (twice)", Near(a.Rotation, 270) && Near(b.Rotation, 300), a.Rotation + " / " + b.Rotation);
            ShapeTools.Transform(r, TransformCommand.RotateRight);

            ShapeTools.Transform(r, TransformCommand.FlipHorizontal);
            bool flipped = a.HorizontalFlip == T && Near(a.Left, 100) && Near(a.Top, 100);
            ShapeTools.Transform(r, TransformCommand.FlipHorizontal);
            Check("flip horizontal, in place, twice = back", flipped && a.HorizontalFlip == F, "");
            ShapeTools.Transform(r, TransformCommand.FlipVertical);
            Check("flip vertical", a.VerticalFlip == T && Near(a.Left, 100) && Near(a.Top, 100), "");
            Delete("ra", "rb");
        }

        /// <summary>Two overlapping rectangles; the first selected one is blue and decides the result's format.</summary>
        private static void MergeTwoRectangles(MergeCommand cmd)
        {
            PowerPoint.Shape p = Add(Office.MsoAutoShapeType.msoShapeRectangle, 100, 300, 100, 100, "mp");
            p.Fill.ForeColor.RGB = 0xFF0000;      // blue (BGR)
            PowerPoint.Shape q = Add(Office.MsoAutoShapeType.msoShapeRectangle, 150, 350, 100, 100, "mq");
            q.Fill.ForeColor.RGB = 0x00FF00;      // green
            int before = slide.Shapes.Count;

            MergeResult res = ShapeTools.Merge(slide.Shapes.Range(new object[] { "mp", "mq" }), false, cmd, slide.Shapes);
            int pieces = slide.Shapes.Count - before + 2;
            int expected = cmd == MergeCommand.Fragment ? 3 : 1;
            bool blue = true;
            double minLeft = double.MaxValue, maxRight = double.MinValue;
            for (int i = 1; i <= slide.Shapes.Count; i++)
            {
                PowerPoint.Shape s = slide.Shapes[i];
                if (s.Fill.ForeColor.RGB != 0xFF0000) blue = false;
                minLeft = Math.Min(minLeft, s.Left);
                maxRight = Math.Max(maxRight, s.Left + s.Width);
            }
            // subtract keeps only the part of the blue square outside the green one
            bool extent = cmd == MergeCommand.Subtract ? Near(minLeft, 100) && Near(maxRight, 200) : Near(minLeft, 100) && Near(maxRight, 250);
            Check("merge " + cmd, res == MergeResult.Done && pieces == expected && blue && extent,
                  pieces + " shape(s), first shape's colour " + blue + ", extent " + minLeft.ToString("0.#") + ".." + maxRight.ToString("0.#"));
            while (slide.Shapes.Count > 0) slide.Shapes[1].Delete();
        }

        private static void MergeRefusals()
        {
            Add(Office.MsoAutoShapeType.msoShapeRectangle, 100, 100, 50, 50, "one");
            MergeResult one = ShapeTools.Merge(slide.Shapes.Range(new object[] { "one" }), false, MergeCommand.Union, slide.Shapes);
            Add(Office.MsoAutoShapeType.msoShapeRectangle, 300, 100, 50, 50, "g1");
            Add(Office.MsoAutoShapeType.msoShapeRectangle, 320, 120, 50, 50, "g2");
            PowerPoint.Shape grp = slide.Shapes.Range(new object[] { "g1", "g2" }).Group();
            grp.Name = "grp";
            MergeResult withGroup = ShapeTools.Merge(slide.Shapes.Range(new object[] { "one", "grp" }), false, MergeCommand.Union, slide.Shapes);
            MergeResult inside = ShapeTools.Merge(grp.GroupItems.Range(new object[] { 1, 2 }), true, MergeCommand.Union, slide.Shapes);
            Check("merge refuses 1 shape / groups", one == MergeResult.NeedTwo && withGroup == MergeResult.Grouped && inside == MergeResult.Grouped
                  && slide.Shapes.Count == 2, one + " / " + withGroup + " / " + inside);
            while (slide.Shapes.Count > 0) slide.Shapes[1].Delete();
        }

        /// <summary>Every colour survives RGB -> HSV -> RGB, so the picker's handles land where the colour is.</summary>
        private static void ColourMaths()
        {
            int bad = 0;
            for (int r = 0; r < 256; r += 5)
                for (int g = 0; g < 256; g += 5)
                    for (int b = 0; b < 256; b += 5)
                    {
                        Color c = Color.FromArgb(r, g, b);
                        float h, s, v;
                        ColorPickerForm.ToHsv(c, out h, out s, out v);
                        if (ColorPickerForm.FromHsv(h, s, v).ToArgb() != c.ToArgb()) bad++;
                    }
            Check("colour maths: RGB -> HSV -> RGB", bad == 0, bad + " colours changed");
        }

        // ------------------------------------------------------------------ helpers

        private static PowerPoint.Shape Add(Office.MsoAutoShapeType type, float l, float t, float w, float h, string name)
        {
            PowerPoint.Shape s = slide.Shapes.AddShape(type, l, t, w, h);
            s.Name = name;
            return s;
        }

        private static void Delete(params string[] names) { foreach (string n in names) slide.Shapes[n].Delete(); }
        private static bool Near(double a, double b) { return Math.Abs(a - b) < 0.05; }

        private static void Check(string name, bool ok, string detail)
        {
            if (ok) passed++; else failed++;
            Console.WriteLine("{0,-48} {1,-8} {2}", name, ok ? "OK" : "MISMATCH", detail);
        }
    }
}
