// Drives the line cap command against a real PowerPoint instance. The object model cannot read a
// cap type either, so every result is checked in a saved copy of the presentation.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Xml;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptSelectionExport
{
    internal static class CapTests
    {
        private static readonly Office.MsoTriState T = Office.MsoTriState.msoTrue, F = Office.MsoTriState.msoFalse;
        private static int passed, failed;
        private static string dir;

        public static void Run(PowerPoint.Application app)
        {
            string tempRoot = Path.Combine(Path.GetTempPath(), "PptSelectionExport");
            dir = Path.Combine(tempRoot, "captests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            int tempDirsBefore = Directory.GetDirectories(tempRoot).Length;

            PowerPoint.Presentation pres = app.Presentations.Add(F);
            try
            {
                // Three slides whose first shapes share an id. The slide under test is then moved to
                // the front, so that its place in the presentation is not the order it was made in.
                var twins = new List<PowerPoint.Shape>();
                for (int i = 1; i <= 3; i++)
                {
                    PowerPoint.Shape twin = pres.Slides.Add(i, PowerPoint.PpSlideLayout.ppLayoutBlank).Shapes.AddLine(20, 20, 200, 20);
                    twin.Name = "twin";
                    twin.Line.Weight = 5;
                    twins.Add(twin);
                }
                PowerPoint.Slide slide = pres.Slides[3];
                PowerPoint.Shapes sh = slide.Shapes;

                PowerPoint.Shape line = sh.AddLine(50, 60, 300, 100); line.Name = "line";
                line.Line.Weight = 6; line.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash; line.Line.ForeColor.RGB = 0x0000FF;
                line.Line.BeginArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadOval; line.Line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadTriangle;

                PowerPoint.Shape rect = sh.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, 50, 130, 260, 90); rect.Name = "rect";
                rect.Fill.TwoColorGradient(Office.MsoGradientStyle.msoGradientHorizontal, 1);
                rect.Shadow.Visible = T; rect.Line.Weight = 3; rect.Line.DashStyle = Office.MsoLineDashStyle.msoLineLongDashDot; rect.Rotation = 10;
                PowerPoint.TextRange tr = rect.TextFrame.TextRange;
                tr.Text = "AAA BBB CCC";
                tr.Font.Size = 20;
                tr.Characters(5, 3).Font.Bold = T; tr.Characters(5, 3).Font.Size = 32;
                tr.Characters(9, 3).Font.Color.RGB = 0x0000FF; tr.Characters(9, 3).Font.Italic = T;
                rect.TextFrame.MarginLeft = 20; rect.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop;

                PowerPoint.Shape conn = sh.AddConnector(Office.MsoConnectorType.msoConnectorElbow, 340, 60, 460, 160); conn.Name = "conn";
                conn.Line.Weight = 4;
                PowerPoint.Shape g1 = sh.AddShape(Office.MsoAutoShapeType.msoShapeArc, 480, 50, 90, 70); g1.Name = "g1"; g1.Line.Weight = 8;
                PowerPoint.Shape g2 = sh.AddLine(600, 50, 700, 120); g2.Name = "g2"; g2.Line.Weight = 7;
                PowerPoint.Shape grp = sh.Range(new object[] { "g1", "g2" }).Group(); grp.Name = "grp";

                PowerPoint.Shape text = sh.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 250, 240, 80); text.Name = "text";
                text.TextFrame.TextRange.Text = "first line\rsecond line";
                text.TextFrame.TextRange.Paragraphs(2, 1).ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                text.TextFrame.TextRange.Paragraphs(2, 1).ParagraphFormat.Bullet.Visible = T;
                text.Line.Visible = T; text.Line.Weight = 4; text.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;

                string png = Path.Combine(dir, "picture.png");
                using (var bmp = new Bitmap(60, 40))
                {
                    using (Graphics g = Graphics.FromImage(bmp)) g.Clear(Color.SteelBlue);
                    bmp.Save(png, ImageFormat.Png);
                }
                PowerPoint.Shape pic = sh.AddPicture(png, F, T, 340, 250, 90, 60); pic.Name = "pic";
                pic.Line.Visible = T; pic.Line.Weight = 5; pic.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;

                PowerPoint.FreeformBuilder fb = sh.BuildFreeform(Office.MsoEditingType.msoEditingAuto, 480, 250);
                fb.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingAuto, 540, 320);
                fb.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingAuto, 600, 250);
                PowerPoint.Shape free = fb.ConvertToShape(); free.Name = "free"; free.Line.Weight = 9;

                PowerPoint.Shape table = sh.AddTable(2, 2, 50, 380, 200, 80); table.Name = "table";
                PowerPoint.Shape other = sh.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 700, 400, 60, 60); other.Name = "other";

                slide.MoveTo(1);
                Check("the twins share one shape id", twins[0].Id == twins[1].Id && twins[1].Id == twins[2].Id, twins[0].Id + " " + twins[1].Id + " " + twins[2].Id);

                string[] names = { "twin", "line", "rect", "conn", "grp", "text", "pic", "free" };
                PowerPoint.Shape[] leaves = { twins[2], line, rect, conn, g1, g2, text, pic, free };
                PowerPoint.ShapeRange all = sh.Range(names);

                // ---- round caps on everything; nothing else in the file may change
                Dictionary<int, XmlElement> before = Snapshot(pres, slide);
                int presentations = app.Presentations.Count;
                var watch = System.Diagnostics.Stopwatch.StartNew();
                int n = LineCaps.SetCap(app, all, CapKind.Round);
                watch.Stop();
                Dictionary<int, XmlElement> after = Snapshot(pres, slide);
                Check("round: lines, shapes, group members, picture", n == leaves.Length && AllHave(after, leaves, "rnd"), n + " changed, " + Caps(after, leaves) + ", " + watch.ElapsedMilliseconds + " ms");
                Check("nothing but the cap changed in the file", SameExceptCap(before, after, leaves, other), Differences(before, after, leaves, other));
                Check("unselected shape has no cap", CapOf(after, other) == "", CapOf(after, other));
                Check("same shape id on the other slides: untouched",
                      CapOf(Snapshot(pres, pres.Slides[2]), twins[0]) == "" && CapOf(Snapshot(pres, pres.Slides[3]), twins[1]) == "", "");

                tr = rect.TextFrame.TextRange;
                Check("mixed text formats survive",
                      tr.Text == "AAA BBB CCC" && Near(tr.Characters(1, 3).Font.Size, 20) && tr.Characters(1, 3).Font.Bold == F
                      && Near(tr.Characters(5, 3).Font.Size, 32) && tr.Characters(5, 3).Font.Bold == T
                      && tr.Characters(9, 3).Font.Color.RGB == 0x0000FF && tr.Characters(9, 3).Font.Italic == T
                      && text.TextFrame.TextRange.Paragraphs(2, 1).ParagraphFormat.Bullet.Visible == T
                      && text.TextFrame.TextRange.Paragraphs(1, 1).ParagraphFormat.Bullet.Visible == F,
                      "BBB " + tr.Characters(5, 3).Font.Size + "/" + tr.Characters(5, 3).Font.Bold);
                Check("line and fill formats survive",
                      Near(line.Line.Weight, 6) && line.Line.DashStyle == Office.MsoLineDashStyle.msoLineDash && line.Line.ForeColor.RGB == 0x0000FF
                      && line.Line.EndArrowheadStyle == Office.MsoArrowheadStyle.msoArrowheadTriangle
                      && rect.Fill.Type == Office.MsoFillType.msoFillGradient && rect.Shadow.Visible == T && Near(rect.Rotation, 10) && Near(rect.TextFrame.MarginLeft, 20)
                      && Near(pic.Line.Weight, 5) && Near(g1.Line.Weight, 8), "");
                Check("the temporary copy is closed again", app.Presentations.Count == presentations, app.Presentations.Count + " open");

                // ---- the other two, and the same one again
                n = LineCaps.SetCap(app, all, CapKind.Flat);
                after = Snapshot(pres, slide);
                Check("flat", n == leaves.Length && AllHave(after, leaves, "flat"), n + " changed, " + Caps(after, leaves));
                n = LineCaps.SetCap(app, all, CapKind.Square);
                after = Snapshot(pres, slide);
                Check("square", n == leaves.Length && AllHave(after, leaves, "sq") && SameExceptCap(before, after, leaves, other), n + " changed, " + Caps(after, leaves));
                watch.Restart();
                n = LineCaps.SetCap(app, all, CapKind.Square);
                watch.Stop();
                Check("the cap they already have: nothing to do", n == leaves.Length && AllHave(Snapshot(pres, slide), leaves, "sq"), n + " have it, " + watch.ElapsedMilliseconds + " ms");

                // ---- a shape inside a group, selected on its own (what Selection.ChildShapeRange gives)
                n = LineCaps.SetCap(app, grp.GroupItems.Range(new object[] { 1 }), CapKind.Round);
                after = Snapshot(pres, slide);
                Check("one shape inside a group", n == 1 && CapOf(after, g1) == "rnd" && CapOf(after, g2) == "sq", n + " changed, g1 " + CapOf(after, g1) + ", g2 " + CapOf(after, g2));

                // ---- objects without a line of their own are skipped, not fatal
                int tableOnly = LineCaps.SetCap(app, sh.Range(new object[] { "table" }), CapKind.Round);
                int mixed = LineCaps.SetCap(app, sh.Range(new object[] { "table", "rect" }), CapKind.Round);
                after = Snapshot(pres, slide);
                Check("table is skipped, the rest still changes", tableOnly == 0 && mixed == 1 && CapOf(after, rect) == "rnd", "table alone: " + tableOnly + ", table + rect: " + mixed);

                // ---- a title placeholder with mixed text
                PowerPoint.Slide titled = pres.Slides.Add(pres.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
                PowerPoint.Shape title = titled.Shapes[1];
                PowerPoint.TextRange tt = title.TextFrame.TextRange;
                tt.Text = "Big title here";
                tt.Characters(1, 3).Font.Italic = T; tt.Characters(5, 5).Font.Color.RGB = 0x00A000;
                title.Line.Visible = T; title.Line.Weight = 6; title.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;
                Dictionary<int, XmlElement> titleBefore = Snapshot(pres, titled);
                n = LineCaps.SetCap(app, titled.Shapes.Range(new object[] { 1 }), CapKind.Round);
                Dictionary<int, XmlElement> titleAfter = Snapshot(pres, titled);
                Check("placeholder with mixed text", n == 1 && CapOf(titleAfter, title) == "rnd" && SameExceptCap(titleBefore, titleAfter, new[] { title }, null)
                      && tt.Characters(1, 3).Font.Italic == T && tt.Characters(5, 5).Font.Color.RGB == 0x00A000 && tt.Characters(11, 4).Font.Italic == F,
                      n + " changed, cap " + CapOf(titleAfter, title) + " " + Differences(titleBefore, titleAfter, new[] { title }, null));

                // ---- masters, layouts and notes pages are refused, and nothing is changed there
                PowerPoint.Shape onMaster = pres.SlideMaster.Shapes.AddLine(10, 10, 100, 10); onMaster.Name = "onMaster";
                PowerPoint.CustomLayout layout = pres.SlideMaster.CustomLayouts[1];
                PowerPoint.Shape onLayout = layout.Shapes.AddLine(10, 30, 100, 30); onLayout.Name = "onLayout";
                PowerPoint.ShapeRange onNotes = slide.NotesPage.Shapes.Range(new object[] { 1 });
                Check("master, layout, notes page: refused",
                      Refused(app, pres.SlideMaster.Shapes.Range(new object[] { "onMaster" })) && Refused(app, layout.Shapes.Range(new object[] { "onLayout" })) && Refused(app, onNotes),
                      "master " + Refused(app, pres.SlideMaster.Shapes.Range(new object[] { "onMaster" })) + ", layout " + Refused(app, layout.Shapes.Range(new object[] { "onLayout" })) + ", notes " + Refused(app, onNotes));

                // ---- the trap: a shape on the notes page with the id and the name of a shape on the slide
                PowerPoint.Shape noteLine = slide.NotesPage.Shapes.AddLine(40, 40, 200, 40);
                PowerPoint.Shape sameId = null;
                foreach (PowerPoint.Shape s in new[] { twins[2], line, rect, conn, text, pic, free, other })
                    if (s.Id == noteLine.Id) sameId = s;
                if (sameId == null) Check("notes page shape with a slide shape's id and name", false, "no shape on the slide has id " + noteLine.Id);
                else
                {
                    noteLine.Name = sameId.Name;
                    string capBefore = CapOf(Snapshot(pres, slide), sameId);
                    bool refused = Refused(app, slide.NotesPage.Shapes.Range(new object[] { noteLine.Name }));
                    string capAfter = CapOf(Snapshot(pres, slide), sameId);
                    Check("notes page shape with a slide shape's id and name", refused && capBefore == capAfter && capBefore != "rnd",
                          "id " + noteLine.Id + " '" + sameId.Name + "', refused " + refused + ", slide shape's cap '" + capBefore + "' -> '" + capAfter + "'");
                }

                Check("unselected shape still untouched", CapOf(Snapshot(pres, slide), other) == "", "");
                foreach (string f in Directory.GetFiles(dir)) File.Delete(f);
                Check("no temporary folders left behind", Directory.GetDirectories(tempRoot).Length == tempDirsBefore, Directory.GetDirectories(tempRoot).Length + " now, " + tempDirsBefore + " before");

                Console.WriteLine("cap tests: {0} passed, {1} failed", passed, failed);
            }
            finally
            {
                pres.Saved = T;
                pres.Close();
                try { Directory.Delete(dir, true); } catch { }
            }
        }

        private static bool Refused(PowerPoint.Application app, PowerPoint.ShapeRange range)
        {
            try { LineCaps.SetCap(app, range, CapKind.Round); return false; }
            catch (StyleNotice) { return true; }
        }

        /// <summary>The shapes of a slide as they are in the file right now, by shape id.</summary>
        private static Dictionary<int, XmlElement> Snapshot(PowerPoint.Presentation pres, PowerPoint.Slide slide)
        {
            string file = Path.Combine(dir, Guid.NewGuid().ToString("N") + ".pptx");
            pres.SaveCopyAs(file, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, F);
            var map = new Dictionary<int, XmlElement>();
            using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                XmlDocument doc = LineCaps.LoadXml(LineCaps.SlideEntry(zip, slide.SlideID));
                XmlNamespaceManager nm = LineCaps.Namespaces(doc);
                foreach (XmlNode node in doc.SelectNodes(LineCaps.ShapeXPath, nm))
                {
                    int id;
                    if (LineCaps.TryGetId((XmlElement)node, nm, out id) && !map.ContainsKey(id)) map[id] = (XmlElement)node;
                }
            }
            File.Delete(file);
            return map;
        }

        private static XmlElement LineOf(XmlElement shape)
        {
            return shape.SelectSingleNode("p:spPr/a:ln", LineCaps.Namespaces(shape.OwnerDocument)) as XmlElement;
        }

        private static string CapOf(Dictionary<int, XmlElement> shapes, PowerPoint.Shape s)
        {
            XmlElement shape;
            if (!shapes.TryGetValue(s.Id, out shape)) return "(not in the file)";
            XmlElement ln = LineOf(shape);
            return ln == null ? "" : ln.GetAttribute("cap");
        }

        private static bool AllHave(Dictionary<int, XmlElement> shapes, PowerPoint.Shape[] wanted, string cap)
        {
            foreach (PowerPoint.Shape s in wanted) if (CapOf(shapes, s) != cap) return false;
            return true;
        }

        private static string Caps(Dictionary<int, XmlElement> shapes, PowerPoint.Shape[] wanted)
        {
            var parts = new List<string>();
            foreach (PowerPoint.Shape s in wanted) parts.Add(CapOf(shapes, s));
            return string.Join(" ", parts.ToArray());
        }

        /// <summary>A shape's XML with the cap taken out (and the line element too, when the cap was all it held).</summary>
        private static string WithoutCap(Dictionary<int, XmlElement> shapes, PowerPoint.Shape s)
        {
            XmlElement shape;
            if (!shapes.TryGetValue(s.Id, out shape)) return "(not in the file)";
            shape = (XmlElement)shape.CloneNode(true);
            XmlElement ln = shape.SelectSingleNode("p:spPr/a:ln", LineCaps.Namespaces(shape.OwnerDocument)) as XmlElement;
            if (ln != null)
            {
                ln.RemoveAttribute("cap");
                if (ln.Attributes.Count == 0 && !ln.HasChildNodes) ln.ParentNode.RemoveChild(ln);
            }
            return shape.OuterXml;
        }

        private static bool SameExceptCap(Dictionary<int, XmlElement> a, Dictionary<int, XmlElement> b, PowerPoint.Shape[] shapes, PowerPoint.Shape other)
        {
            return Differences(a, b, shapes, other) == "";
        }

        private static string Differences(Dictionary<int, XmlElement> a, Dictionary<int, XmlElement> b, PowerPoint.Shape[] shapes, PowerPoint.Shape other)
        {
            string result = "";
            var list = new List<PowerPoint.Shape>(shapes);
            if (other != null) list.Add(other);
            foreach (PowerPoint.Shape s in list)
            {
                string x = WithoutCap(a, s), y = WithoutCap(b, s);
                if (x == y) continue;
                int i = 0;
                while (i < x.Length && i < y.Length && x[i] == y[i]) i++;
                int from = Math.Max(0, i - 30);
                result += s.Name + " differs at " + i + ": '" + x.Substring(from, Math.Min(90, x.Length - from)) + "' -> '" + y.Substring(from, Math.Min(90, y.Length - from)) + "'  ";
            }
            return result;
        }

        private static bool Near(double a, double b) { return Math.Abs(a - b) < 0.011; }

        private static void Check(string name, bool ok, string detail)
        {
            if (ok) passed++; else failed++;
            Console.WriteLine("{0,-48} {1,-8} {2}", name, ok ? "OK" : "MISMATCH", detail);
        }
    }
}
