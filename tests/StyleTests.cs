// Drives the shape style commands (fill, outline colour, weight, dashes) against a real PowerPoint instance.
using System;
using System.Drawing;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptSelectionExport
{
    internal static class StyleTests
    {
        private static readonly Office.MsoTriState T = Office.MsoTriState.msoTrue, F = Office.MsoTriState.msoFalse;
        private static int passed, failed;

        public static void Run(PowerPoint.Application app)
        {
            PowerPoint.Presentation pres = app.Presentations.Add(F);
            try
            {
                PowerPoint.Slide slide = pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                PowerPoint.Shapes sh = slide.Shapes;

                PowerPoint.Shape rect = sh.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 50, 50, 120, 80); rect.Name = "rect";
                PowerPoint.Shape line = sh.AddLine(50, 200, 300, 240); line.Name = "line";
                PowerPoint.Shape conn = sh.AddConnector(Office.MsoConnectorType.msoConnectorElbow, 320, 60, 420, 160); conn.Name = "conn";
                PowerPoint.Shape g1 = sh.AddShape(Office.MsoAutoShapeType.msoShapeOval, 450, 50, 80, 60); g1.Name = "g1";
                PowerPoint.Shape g2 = sh.AddShape(Office.MsoAutoShapeType.msoShapeDiamond, 560, 50, 80, 60); g2.Name = "g2";
                PowerPoint.Shape grp = sh.Range(new object[] { "g1", "g2" }).Group(); grp.Name = "grp";
                PowerPoint.Shape text = sh.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 300, 200, 40); text.Name = "text";
                text.TextFrame.TextRange.Text = "Litchi";
                PowerPoint.Shape table = sh.AddTable(2, 2, 300, 300, 200, 80); table.Name = "table";
                PowerPoint.Shape other = sh.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 700, 400, 60, 60); other.Name = "other";
                int otherFill = other.Fill.ForeColor.RGB, otherLine = other.Line.ForeColor.RGB;
                float otherWeight = other.Line.Weight;

                PowerPoint.ShapeRange all = sh.Range(new object[] { "rect", "line", "conn", "grp", "text" });
                Color teal = Color.FromArgb(0x14, 0xB8, 0xA6), pink = Color.FromArgb(0xEC, 0x48, 0x99);
                int tealOle = ColorTranslator.ToOle(teal), pinkOle = ColorTranslator.ToOle(pink);

                // ---- fill: lines and connectors are skipped, every shape in a group is changed
                int n = StyleTools.SetFill(all, teal);
                Check("fill colour: shapes, group members, text box", n == 4 && rect.Fill.ForeColor.RGB == tealOle && rect.Fill.Visible == T
                      && g1.Fill.ForeColor.RGB == tealOle && g2.Fill.ForeColor.RGB == tealOle && text.Fill.ForeColor.RGB == tealOle
                      && rect.Fill.Type == Office.MsoFillType.msoFillSolid, n + " changed");

                rect.Fill.TwoColorGradient(Office.MsoGradientStyle.msoGradientHorizontal, 1);
                StyleTools.SetFill(sh.Range(new object[] { "rect" }), pink);
                Check("fill colour replaces a gradient", rect.Fill.Type == Office.MsoFillType.msoFillSolid && rect.Fill.ForeColor.RGB == pinkOle, rect.Fill.Type.ToString());

                n = StyleTools.SetFill(sh.Range(new object[] { "line", "conn" }), teal);
                Check("fill on lines only: nothing to change", n == 0, n + " changed");

                n = StyleTools.SetFill(all, null);
                Check("no fill", n == 4 && rect.Fill.Visible == F && g1.Fill.Visible == F && text.Fill.Visible == F, n + " changed");

                // ---- outline colour: lines and connectors too
                n = StyleTools.SetOutlineColor(all, pink);
                Check("outline colour: shapes, lines, connectors", n == 6 && rect.Line.ForeColor.RGB == pinkOle && line.Line.ForeColor.RGB == pinkOle
                      && conn.Line.ForeColor.RGB == pinkOle && g2.Line.ForeColor.RGB == pinkOle && text.Line.Visible == T, n + " changed");

                n = StyleTools.SetOutlineColor(sh.Range(new object[] { "rect", "text" }), null);
                Check("no outline", n == 2 && rect.Line.Visible == F && text.Line.Visible == F && line.Line.Visible == T, n + " changed");

                // ---- weight: brings a hidden outline back, as PowerPoint does
                n = StyleTools.SetWeight(all, 2.25f);
                Check("weight 2.25 pt", n == 6 && Near(rect.Line.Weight, 2.25) && Near(line.Line.Weight, 2.25) && Near(g1.Line.Weight, 2.25)
                      && rect.Line.Visible == T && rect.Line.ForeColor.RGB == pinkOle, n + " changed, rect " + rect.Line.Weight);
                StyleTools.SetWeight(sh.Range(new object[] { "line" }), 0.001f);
                Check("weight is kept inside the limits", Near(line.Line.Weight, StyleTools.MinWeight), line.Line.Weight.ToString());

                // ---- dashes
                bool dashesOk = true;
                string detail = "";
                foreach (DashKind d in StyleTools.Dashes)
                {
                    StyleTools.SetDash(all, d);
                    if (line.Line.DashStyle != StyleTools.ToMso(d) || rect.Line.DashStyle != StyleTools.ToMso(d) || g2.Line.DashStyle != StyleTools.ToMso(d))
                    {
                        dashesOk = false;
                        detail += d + " ";
                    }
                }
                Check("all 8 dash styles on line, shape, group member", dashesOk, detail);

                // ---- shapes inside a group, selected on their own (what Selection.ChildShapeRange gives)
                n = StyleTools.SetFill(grp.GroupItems.Range(new object[] { 1 }), pink);
                Check("one shape inside a group", n == 1 && g1.Fill.ForeColor.RGB == pinkOle && g2.Fill.Visible == F, n + " changed");

                // ---- objects that cannot take the setting are skipped, not fatal
                int tableOnly = StyleTools.SetOutlineColor(sh.Range(new object[] { "table" }), pink);
                int mixed = StyleTools.SetWeight(sh.Range(new object[] { "table", "rect" }), 3f);
                Check("table is skipped, the rest still changes", mixed >= 1 && Near(rect.Line.Weight, 3), "table alone: " + tableOnly + ", table + rect: " + mixed);

                Check("unselected shape untouched", other.Fill.ForeColor.RGB == otherFill && other.Line.ForeColor.RGB == otherLine && Near(other.Line.Weight, otherWeight), "");

                // ---- typed weights
                float w;
                bool parse = StyleTools.TryParseWeight("1.25", out w) && Near(w, 1.25) && StyleTools.TryParseWeight(" 2,5 磅 ", out w) && Near(w, 2.5)
                             && StyleTools.TryParseWeight("0.75pt", out w) && Near(w, 0.75)
                             && !StyleTools.TryParseWeight("", out w) && !StyleTools.TryParseWeight("abc", out w)
                             && !StyleTools.TryParseWeight("0", out w) && !StyleTools.TryParseWeight("500", out w) && !StyleTools.TryParseWeight("-1", out w);
                Check("typed weights: accepted and refused", parse, "");

                Console.WriteLine("style tests: {0} passed, {1} failed", passed, failed);
            }
            finally
            {
                pres.Saved = T;
                pres.Close();
            }
        }

        private static bool Near(double a, double b) { return Math.Abs(a - b) < 0.011; }

        private static void Check(string name, bool ok, string detail)
        {
            if (ok) passed++; else failed++;
            Console.WriteLine("{0,-48} {1,-8} {2}", name, ok ? "OK" : "MISMATCH", detail);
        }
    }
}
