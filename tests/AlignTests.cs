// Drives the alignment code against a real PowerPoint instance, without the ribbon.
// Where PowerPoint has the same command, the result must match PowerPoint's own result.
using System;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptSelectionExport
{
    internal static class AlignTests
    {
        private static readonly Office.MsoTriState T = Office.MsoTriState.msoTrue, F = Office.MsoTriState.msoFalse;
        private const double Tolerance = 0.02;      // points
        private static readonly AlignCommand[] Commands = (AlignCommand[])Enum.GetValues(typeof(AlignCommand));

        private static PowerPoint.Slide slide;
        private static float slideW, slideH;
        private static int passed, failed;

        public static void Run(PowerPoint.Application app)
        {
            PowerPoint.Presentation pres = app.Presentations.Add(F);
            try
            {
                slide = pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                slideW = pres.PageSetup.SlideWidth;
                slideH = pres.PageSetup.SlideHeight;

                foreach (AlignCommand cmd in Commands) SameAsPowerPoint(cmd);
                EvenGapsAroundTiltedShapes(true);
                EvenGapsAroundTiltedShapes(false);
                foreach (AlignCommand cmd in Commands) ToFirstSelected(cmd);
                foreach (AlignCommand cmd in Commands) GluedConnectorsFollow(cmd);
                foreach (AlignCommand cmd in Commands) OneShapeAndItsConnectorGoToSlide(cmd);
                TwoShapesCannotBeDistributed();
                ShapesInsideAGroup();
                Console.WriteLine("align tests: {0} passed, {1} failed", passed, failed);
            }
            finally
            {
                pres.Saved = T;
                pres.Close();
            }
        }

        // ------------------------------------------------------------------ cases

        /// <summary>
        /// Selection-relative align and distribute give exactly what PowerPoint's own commands give.
        /// (Not for distributing a shape turned by an odd angle: PowerPoint measures that one by its
        /// unrotated frame, we by what is visible. See <see cref="EvenGapsAroundTiltedShapes"/>.)
        /// </summary>
        private static void SameAsPowerPoint(AlignCommand cmd)
        {
            bool distribute = cmd == AlignCommand.DistributeHorizontally || cmd == AlignCommand.DistributeVertically;
            string[] ours = MakeMixedSet("a_", !distribute), theirs = MakeMixedSet("b_", !distribute);
            Native(slide.Shapes.Range(theirs), cmd, F);
            int unmoved;
            AlignResult res = Aligner.Apply(slide.Shapes.Range(ours), cmd, false, slideW, slideH, out unmoved);
            double worst = WorstDifference(ours, theirs);
            Check("same as PowerPoint: " + cmd, res == AlignResult.Done && unmoved == 0 && worst < Tolerance,
                  "max difference " + worst.ToString("0.###") + " pt");
            Delete(ours); Delete(theirs);
        }

        /// <summary>Distributing with a tilted shape among them: the visible gaps are all equal, the outer two stay put.</summary>
        private static void EvenGapsAroundTiltedShapes(bool horizontal)
        {
            string[] names = MakeMixedSet("t_", true);
            var before = new Dictionary<string, Bx>();
            foreach (string n in names) before[n] = Bx.Of(slide.Shapes[n]);
            int unmoved;
            Aligner.Apply(slide.Shapes.Range(names), horizontal ? AlignCommand.DistributeHorizontally : AlignCommand.DistributeVertically,
                          false, slideW, slideH, out unmoved);

            var order = new List<string>(names);
            order.Sort((p, q) => before[p].Center(horizontal).CompareTo(before[q].Center(horizontal)));
            var gaps = new List<double>();
            for (int i = 1; i < order.Count; i++)
                gaps.Add(Bx.Of(slide.Shapes[order[i]]).Start(horizontal) - Bx.Of(slide.Shapes[order[i - 1]]).End(horizontal));
            bool ok = unmoved == 0 && Same(Bx.Of(slide.Shapes[order[0]]), before[order[0]])
                      && Same(Bx.Of(slide.Shapes[order[order.Count - 1]]), before[order[order.Count - 1]]);
            foreach (double g in gaps)
                if (Math.Abs(g - gaps[0]) > Tolerance) ok = false;
            Check("even gaps around tilted shapes: " + (horizontal ? "horizontal" : "vertical"), ok,
                  "gaps " + string.Join(" / ", gaps.ConvertAll(g => g.ToString("0.##")).ToArray()));
            Delete(names);
        }

        /// <summary>
        /// "Align to object": the first shape in the selection (here: the top-most in z-order,
        /// so z-order cannot be what decides) stays put; the others line up with it or spread across it.
        /// </summary>
        private static void ToFirstSelected(AlignCommand cmd)
        {
            PowerPoint.Shape a = Rect(40, 40, 60, 40, "a");
            PowerPoint.Shape b = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, 650, 380, 90, 40);
            b.Name = "b"; b.Rotation = 45;
            PowerPoint.Shape c = Rect(100, 420, 30, 70, "c");
            PowerPoint.Shape key = Rect(300, 150, 260, 140, "key");
            Bx keyBefore = Bx.Of(key);
            var before = new[] { Bx.Of(a), Bx.Of(b), Bx.Of(c) };

            int unmoved;
            Aligner.Apply(slide.Shapes.Range(new object[] { "key", "a", "b", "c" }), cmd, true, slideW, slideH, out unmoved);
            Bx k = Bx.Of(key);
            var now = new[] { Bx.Of(a), Bx.Of(b), Bx.Of(c) };
            bool ok = unmoved == 0 && Same(k, keyBefore);
            string detail = "";
            bool horizontal = cmd == AlignCommand.Left || cmd == AlignCommand.Center || cmd == AlignCommand.Right
                              || cmd == AlignCommand.DistributeHorizontally;
            for (int i = 0; i < now.Length; i++)
            {
                // nothing moves along the other axis
                double other = horizontal ? now[i].T - before[i].T : now[i].L - before[i].L;
                if (Math.Abs(other) > Tolerance) { ok = false; detail += " moved across"; }
            }
            if (cmd == AlignCommand.DistributeHorizontally || cmd == AlignCommand.DistributeVertically)
            {
                // equal gaps: key start | x | x | x | key end
                var sorted = new List<Bx>(now);
                sorted.Sort((p, q) => p.Center(horizontal).CompareTo(q.Center(horizontal)));
                var gaps = new List<double> { sorted[0].Start(horizontal) - k.Start(horizontal) };
                for (int i = 1; i < sorted.Count; i++) gaps.Add(sorted[i].Start(horizontal) - sorted[i - 1].End(horizontal));
                gaps.Add(k.End(horizontal) - sorted[sorted.Count - 1].End(horizontal));
                foreach (double g in gaps)
                    if (Math.Abs(g - gaps[0]) > Tolerance) { ok = false; }
                detail += " gaps " + string.Join(" / ", gaps.ConvertAll(g => g.ToString("0.##")).ToArray());
            }
            else
            {
                foreach (Bx o in now)
                    if (Math.Abs(Edge(o, cmd) - Edge(k, cmd)) > Tolerance) { ok = false; detail += " edge " + Edge(o, cmd).ToString("0.##"); }
                detail += " key edge " + Edge(k, cmd).ToString("0.##");
            }
            Check("align to first selected: " + cmd, ok, detail.Trim());
            Delete(new[] { "a", "b", "c", "key" });
        }

        /// <summary>Connectors glued to shapes are not aligned themselves; they stay glued and follow.</summary>
        private static void GluedConnectorsFollow(AlignCommand cmd)
        {
            string[] ours = { Rect(100, 100, 100, 60, "ca").Name, Rect(400, 260, 120, 50, "cb").Name, Rect(700, 160, 80, 80, "cc").Name };
            string[] theirs = { Rect(100, 100, 100, 60, "ta").Name, Rect(400, 260, 120, 50, "tb").Name, Rect(700, 160, 80, 80, "tc").Name };
            PowerPoint.Shape k = slide.Shapes.AddConnector(Office.MsoConnectorType.msoConnectorElbow, 0, 0, 10, 10);
            k.Name = "ck";
            k.ConnectorFormat.BeginConnect(slide.Shapes["ca"], 4);
            k.ConnectorFormat.EndConnect(slide.Shapes["cb"], 2);
            k.RerouteConnections();

            int unmoved;
            AlignResult res = Aligner.Apply(slide.Shapes.Range(new object[] { "ca", "cb", "cc", "ck" }), cmd, false, slideW, slideH, out unmoved);
            Native(slide.Shapes.Range(theirs), cmd, F);
            double worst = WorstDifference(ours, theirs);
            bool glued = k.ConnectorFormat.BeginConnected == T && k.ConnectorFormat.EndConnected == T;
            Check("glued connector follows: " + cmd, res == AlignResult.Done && worst < Tolerance && glued,
                  "max difference " + worst.ToString("0.###") + " pt, still glued " + glued);
            Delete(new[] { "ck" }); Delete(ours); Delete(theirs);
        }

        /// <summary>One shape plus the connector glued to it is "one object": it goes to the slide, like a single shape.</summary>
        private static void OneShapeAndItsConnectorGoToSlide(AlignCommand cmd)
        {
            PowerPoint.Shape x = Rect(130, 120, 90, 50, "x");
            x.Rotation = 20;
            PowerPoint.Shape y = Rect(500, 380, 60, 40, "y");
            PowerPoint.Shape k = slide.Shapes.AddConnector(Office.MsoConnectorType.msoConnectorStraight, 0, 0, 10, 10);
            k.Name = "k";
            k.ConnectorFormat.BeginConnect(x, 3);
            k.ConnectorFormat.EndConnect(y, 1);
            PowerPoint.Shape twin = Rect(130, 120, 90, 50, "twin");
            twin.Rotation = 20;

            int unmoved;
            Aligner.Apply(slide.Shapes.Range(new object[] { "x", "k" }), cmd, false, slideW, slideH, out unmoved);
            Native(slide.Shapes.Range(new object[] { "twin" }), cmd, T);
            double worst = WorstDifference(new[] { "x" }, new[] { "twin" });
            bool glued = k.ConnectorFormat.BeginConnected == T && k.ConnectorFormat.EndConnected == T;
            Check("shape + its connector to slide: " + cmd, worst < Tolerance && glued,
                  "max difference " + worst.ToString("0.###") + " pt, still glued " + glued);
            Delete(new[] { "k", "x", "y", "twin" });
        }

        private static void TwoShapesCannotBeDistributed()
        {
            Rect(100, 100, 50, 50, "p");
            Rect(300, 200, 50, 50, "q");
            int unmoved;
            AlignResult res = Aligner.Apply(slide.Shapes.Range(new object[] { "p", "q" }), AlignCommand.DistributeHorizontally, false, slideW, slideH, out unmoved);
            bool still = slide.Shapes["p"].Left == 100 && slide.Shapes["q"].Left == 300;
            Check("two shapes: distribute refused", res == AlignResult.NeedThree && still, res.ToString());
            Delete(new[] { "p", "q" });
        }

        /// <summary>Shapes selected inside a group, in both modes.</summary>
        private static void ShapesInsideAGroup()
        {
            PowerPoint.Shape g = Group("g", 0);
            PowerPoint.Shape h = Group("h", 0);
            int unmoved;
            Aligner.Apply(g.GroupItems.Range(new object[] { 1, 2, 3 }), AlignCommand.Top, false, slideW, slideH, out unmoved);
            h.GroupItems.Range(new object[] { 1, 2, 3 }).Align(Office.MsoAlignCmd.msoAlignTops, F);
            double worst = 0;
            for (int i = 1; i <= 3; i++)
            {
                worst = Math.Max(worst, Math.Abs(g.GroupItems[i].Left - h.GroupItems[i].Left));
                worst = Math.Max(worst, Math.Abs(g.GroupItems[i].Top - h.GroupItems[i].Top));
            }
            Check("inside a group: same as PowerPoint", worst < Tolerance, "max difference " + worst.ToString("0.###") + " pt");
            Delete(new[] { "g", "h" });

            // the third child is selected first, so it is the reference
            PowerPoint.Shape e = Group("e", 0);
            Bx keyBefore = Bx.Of(e.GroupItems[3]);
            Aligner.Apply(e.GroupItems.Range(new object[] { 3, 1 }), AlignCommand.Right, true, slideW, slideH, out unmoved);
            Bx key = Bx.Of(e.GroupItems[3]), moved = Bx.Of(e.GroupItems[1]);
            Check("inside a group: align to first selected", Same(key, keyBefore) && Math.Abs(moved.R - key.R) < Tolerance,
                  "right edges " + moved.R.ToString("0.##") + " / " + key.R.ToString("0.##"));
            Delete(new[] { "e" });
        }

        // ------------------------------------------------------------------ helpers

        /// <summary>
        /// Rotated shapes, a line, a text box and a group, with distinct centres in both directions.
        /// <paramref name="tilted"/> adds a shape turned by 30 degrees.
        /// </summary>
        private static string[] MakeMixedSet(string p, bool tilted)
        {
            PowerPoint.Shapes sh = slide.Shapes;
            var names = new List<string>();
            names.Add(Rect(60, 80, 120, 60, p + "rect").Name);
            PowerPoint.Shape s;
            if (tilted)
            {
                s = sh.AddShape(Office.MsoAutoShapeType.msoShapeOval, 300, 200, 80, 120);
                s.Name = p + "oval30"; s.Rotation = 30; names.Add(s.Name);
            }
            s = Rect(520, 95, 160, 40, p + "rect90"); s.Rotation = 90; names.Add(s.Name);
            names.Add(Rect(150, 380, 400, 30, p + "wide").Name);
            s = sh.AddLine(700, 300, 820, 420); s.Name = p + "line"; names.Add(s.Name);
            s = sh.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 230, 460, 200, 40);
            s.Name = p + "text"; s.TextFrame.TextRange.Text = "Text"; names.Add(s.Name);
            names.Add(Group(p + "group", 600).Name);
            return names.ToArray();
        }

        private static PowerPoint.Shape Rect(float l, float t, float w, float h, string name)
        {
            PowerPoint.Shape s = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, l, t, w, h);
            s.Name = name;
            return s;
        }

        private static PowerPoint.Shape Group(string name, float dx)
        {
            PowerPoint.Shape g1 = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeDiamond, 40 + dx, 40, 60, 60);
            g1.Name = name + "1";
            PowerPoint.Shape g2 = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeHexagon, 120 + dx, 90, 70, 50);
            g2.Name = name + "2";
            PowerPoint.Shape g3 = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, 210 + dx, 20, 40, 90);
            g3.Name = name + "3"; g3.Rotation = 30;
            PowerPoint.Shape g = slide.Shapes.Range(new object[] { g1.Name, g2.Name, g3.Name }).Group();
            g.Name = name;
            return g;
        }

        private static void Native(PowerPoint.ShapeRange r, AlignCommand cmd, Office.MsoTriState toSlide)
        {
            switch (cmd)
            {
                case AlignCommand.Left: r.Align(Office.MsoAlignCmd.msoAlignLefts, toSlide); break;
                case AlignCommand.Center: r.Align(Office.MsoAlignCmd.msoAlignCenters, toSlide); break;
                case AlignCommand.Right: r.Align(Office.MsoAlignCmd.msoAlignRights, toSlide); break;
                case AlignCommand.Top: r.Align(Office.MsoAlignCmd.msoAlignTops, toSlide); break;
                case AlignCommand.Middle: r.Align(Office.MsoAlignCmd.msoAlignMiddles, toSlide); break;
                case AlignCommand.Bottom: r.Align(Office.MsoAlignCmd.msoAlignBottoms, toSlide); break;
                case AlignCommand.DistributeHorizontally: r.Distribute(Office.MsoDistributeCmd.msoDistributeHorizontally, toSlide); break;
                case AlignCommand.DistributeVertically: r.Distribute(Office.MsoDistributeCmd.msoDistributeVertically, toSlide); break;
            }
        }

        private static double WorstDifference(string[] ours, string[] theirs)
        {
            double worst = 0;
            for (int i = 0; i < ours.Length; i++)
            {
                PowerPoint.Shape a = slide.Shapes[ours[i]], b = slide.Shapes[theirs[i]];
                worst = Math.Max(worst, Math.Max(Math.Abs(a.Left - b.Left), Math.Abs(a.Top - b.Top)));
            }
            return worst;
        }

        private static double Edge(Bx b, AlignCommand cmd)
        {
            switch (cmd)
            {
                case AlignCommand.Left: return b.L;
                case AlignCommand.Center: return (b.L + b.R) / 2;
                case AlignCommand.Right: return b.R;
                case AlignCommand.Top: return b.T;
                case AlignCommand.Middle: return (b.T + b.B) / 2;
                default: return b.B;
            }
        }

        private static bool Same(Bx a, Bx b)
        {
            return Math.Abs(a.L - b.L) < Tolerance && Math.Abs(a.T - b.T) < Tolerance && Math.Abs(a.R - b.R) < Tolerance && Math.Abs(a.B - b.B) < Tolerance;
        }

        private static void Delete(string[] names)
        {
            foreach (string n in names) slide.Shapes[n].Delete();
        }

        private static void Check(string name, bool ok, string detail)
        {
            if (ok) passed++; else failed++;
            Console.WriteLine("{0,-48} {1,-8} {2}", name, ok ? "OK" : "MISMATCH", detail);
        }

        /// <summary>Rotated bounding box, computed here independently of the code under test.</summary>
        private struct Bx
        {
            public double L, T, R, B;
            public double Start(bool h) { return h ? L : T; }
            public double End(bool h) { return h ? R : B; }
            public double Center(bool h) { return h ? (L + R) / 2 : (T + B) / 2; }

            public static Bx Of(PowerPoint.Shape s)
            {
                double rad = s.Rotation * Math.PI / 180;
                double c = Math.Abs(Math.Cos(rad)), n = Math.Abs(Math.Sin(rad));
                double w = s.Width * c + s.Height * n, h = s.Width * n + s.Height * c;
                double cx = s.Left + s.Width / 2.0, cy = s.Top + s.Height / 2.0;
                return new Bx { L = cx - w / 2, T = cy - h / 2, R = cx + w / 2, B = cy + h / 2 };
            }
        }
    }
}
