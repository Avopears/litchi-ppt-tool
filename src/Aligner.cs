using System;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptSelectionExport
{
    internal enum AlignCommand { Left, Center, Right, Top, Middle, Bottom, DistributeHorizontally, DistributeVertically }

    internal enum AlignResult
    {
        Done,
        /// <summary>Distributing between the selected objects needs at least three of them.</summary>
        NeedThree,
    }

    /// <summary>
    /// Aligns and distributes shapes the way PowerPoint's own commands do: a rotated shape counts
    /// with its rotated bounding box, distributing leaves equal gaps, and a single object is
    /// aligned to the slide. The differences:
    /// <list type="bullet">
    /// <item>"Align to object": the first selected shape is the reference and does not move.
    ///   Distributing then spreads the other shapes evenly across it, the way PowerPoint spreads
    ///   shapes across the slide.</item>
    /// <item>Connectors glued to a shape are not moved themselves; they follow the shapes they are
    ///   glued to. (PowerPoint's own commands move them too, which tears them off.)</item>
    /// <item>Distributing measures a rotated shape by its rotated bounding box, like aligning does,
    ///   so the gaps you see are equal. (PowerPoint's own distribute measures a shape turned by,
    ///   say, 30 degrees by its unrotated frame.)</item>
    /// </list>
    /// </summary>
    internal static class Aligner
    {
        private const Office.MsoTriState True = Office.MsoTriState.msoTrue;
        private const double Epsilon = 0.0001;

        /// <param name="range">The selected shapes, in the order they were selected.</param>
        /// <param name="cmd">What to do.</param>
        /// <param name="toFirstSelected">Use the first selected shape as the fixed reference.</param>
        /// <param name="slideW">Slide width in points.</param>
        /// <param name="slideH">Slide height in points.</param>
        /// <param name="unmoved">How many shapes refused to move (locked ones, for example).</param>
        public static AlignResult Apply(PowerPoint.ShapeRange range, AlignCommand cmd, bool toFirstSelected,
                                        float slideW, float slideH, out int unmoved)
        {
            unmoved = 0;
            int n = range.Count;
            if (n == 0) return AlignResult.Done;
            if (n == 1)
            {
                // One object goes to the slide. PowerPoint's own command knows the page it is on
                // (a notes page, for example), so let it do this.
                AlignToSlide(range, cmd);
                return AlignResult.Done;
            }

            var all = new List<Item>(n);
            for (int i = 1; i <= n; i++) all.Add(new Item(range[i], i));
            List<Item> items = all.FindAll(it => !it.Glued);
            if (items.Count == 0) items = all;      // nothing but connectors selected: move those

            bool distribute = cmd == AlignCommand.DistributeHorizontally || cmd == AlignCommand.DistributeVertically;
            bool horizontal = cmd == AlignCommand.Left || cmd == AlignCommand.Center || cmd == AlignCommand.Right
                              || cmd == AlignCommand.DistributeHorizontally;

            if (items.Count == 1)
            {
                // one shape plus connectors glued to it: that shape goes to the slide
                var slide = new Box(0, 0, slideW, slideH);
                if (distribute) SpreadInside(items, slide, horizontal);
                else AlignTo(items, slide, cmd);
            }
            else if (toFirstSelected)
            {
                Box key = items[0].Bounds;
                List<Item> others = items.GetRange(1, items.Count - 1);
                if (distribute) SpreadInside(others, key, horizontal);
                else AlignTo(others, key, cmd);
            }
            else if (distribute)
            {
                if (items.Count < 3) return AlignResult.NeedThree;
                SpreadBetween(items, horizontal);
            }
            else
            {
                AlignTo(items, Union(items), cmd);
            }

            foreach (Item it in items)
                if (!it.Move()) unmoved++;
            return AlignResult.Done;
        }

        // ------------------------------------------------------------------ layout rules

        /// <summary>Moves each item so that the chosen edge or centre line matches the target's.</summary>
        private static void AlignTo(List<Item> items, Box target, AlignCommand cmd)
        {
            foreach (Item it in items)
            {
                Box b = it.Bounds;
                switch (cmd)
                {
                    case AlignCommand.Left: it.Dx = target.L - b.L; break;
                    case AlignCommand.Center: it.Dx = target.CenterX - b.CenterX; break;
                    case AlignCommand.Right: it.Dx = target.R - b.R; break;
                    case AlignCommand.Top: it.Dy = target.T - b.T; break;
                    case AlignCommand.Middle: it.Dy = target.CenterY - b.CenterY; break;
                    case AlignCommand.Bottom: it.Dy = target.B - b.B; break;
                }
            }
        }

        /// <summary>
        /// PowerPoint's "distribute": the first and the last item (by centre) stay where they are,
        /// the ones between them are moved so that all gaps are equal.
        /// </summary>
        private static void SpreadBetween(List<Item> items, bool horizontal)
        {
            List<Item> sorted = SortedByCenter(items, horizontal);
            Item first = sorted[0], last = sorted[sorted.Count - 1];
            double span = last.Bounds.End(horizontal) - first.Bounds.Start(horizontal);
            double gap = (span - TotalSize(sorted, horizontal)) / (sorted.Count - 1);
            double pos = first.Bounds.End(horizontal) + gap;
            for (int i = 1; i < sorted.Count - 1; i++)
            {
                Item it = sorted[i];
                it.Shift(pos - it.Bounds.Start(horizontal), horizontal);
                pos += it.Bounds.Size(horizontal) + gap;
            }
        }

        /// <summary>
        /// Spreads the items across a container with equal gaps between them and at both ends;
        /// a single item ends up centred. This is what PowerPoint does relative to the slide.
        /// </summary>
        private static void SpreadInside(List<Item> items, Box container, bool horizontal)
        {
            List<Item> sorted = SortedByCenter(items, horizontal);
            double gap = (container.Size(horizontal) - TotalSize(sorted, horizontal)) / (sorted.Count + 1);
            double pos = container.Start(horizontal) + gap;
            foreach (Item it in sorted)
            {
                it.Shift(pos - it.Bounds.Start(horizontal), horizontal);
                pos += it.Bounds.Size(horizontal) + gap;
            }
        }

        private static List<Item> SortedByCenter(List<Item> items, bool horizontal)
        {
            var sorted = new List<Item>(items);
            sorted.Sort((a, b) =>
            {
                int c = a.Bounds.Center(horizontal).CompareTo(b.Bounds.Center(horizontal));
                return c != 0 ? c : a.Order.CompareTo(b.Order);      // keep the selection order for ties
            });
            return sorted;
        }

        private static double TotalSize(List<Item> items, bool horizontal)
        {
            double total = 0;
            foreach (Item it in items) total += it.Bounds.Size(horizontal);
            return total;
        }

        private static Box Union(List<Item> items)
        {
            double l = double.MaxValue, t = double.MaxValue, r = double.MinValue, b = double.MinValue;
            foreach (Item it in items)
            {
                l = Math.Min(l, it.Bounds.L); t = Math.Min(t, it.Bounds.T);
                r = Math.Max(r, it.Bounds.R); b = Math.Max(b, it.Bounds.B);
            }
            return new Box(l, t, r, b);
        }

        private static void AlignToSlide(PowerPoint.ShapeRange range, AlignCommand cmd)
        {
            switch (cmd)
            {
                case AlignCommand.Left: range.Align(Office.MsoAlignCmd.msoAlignLefts, True); break;
                case AlignCommand.Center: range.Align(Office.MsoAlignCmd.msoAlignCenters, True); break;
                case AlignCommand.Right: range.Align(Office.MsoAlignCmd.msoAlignRights, True); break;
                case AlignCommand.Top: range.Align(Office.MsoAlignCmd.msoAlignTops, True); break;
                case AlignCommand.Middle: range.Align(Office.MsoAlignCmd.msoAlignMiddles, True); break;
                case AlignCommand.Bottom: range.Align(Office.MsoAlignCmd.msoAlignBottoms, True); break;
                case AlignCommand.DistributeHorizontally: range.Distribute(Office.MsoDistributeCmd.msoDistributeHorizontally, True); break;
                case AlignCommand.DistributeVertically: range.Distribute(Office.MsoDistributeCmd.msoDistributeVertically, True); break;
            }
        }

        // ------------------------------------------------------------------ shapes

        private sealed class Item
        {
            public readonly PowerPoint.Shape Shape;
            /// <summary>1-based position in the selection.</summary>
            public readonly int Order;
            public readonly Box Bounds;
            /// <summary>A connector with at least one end glued to a shape.</summary>
            public readonly bool Glued;
            public double Dx, Dy;

            public Item(PowerPoint.Shape shape, int order)
            {
                Shape = shape;
                Order = order;
                Bounds = Box.Of(shape);
                Glued = IsGluedConnector(shape);
            }

            public void Shift(double d, bool horizontal)
            {
                if (horizontal) Dx = d; else Dy = d;
            }

            /// <summary>Applies the computed offset; false when PowerPoint refuses to move the shape.</summary>
            public bool Move()
            {
                try
                {
                    if (Math.Abs(Dx) > Epsilon) Shape.IncrementLeft((float)Dx);
                    if (Math.Abs(Dy) > Epsilon) Shape.IncrementTop((float)Dy);
                    return true;
                }
                catch { return false; }
            }

            private static bool IsGluedConnector(PowerPoint.Shape s)
            {
                try
                {
                    if (s.Connector != True) return false;
                    PowerPoint.ConnectorFormat f = s.ConnectorFormat;
                    return f.BeginConnected == True || f.EndConnected == True;
                }
                catch { return false; }
            }
        }

        /// <summary>An axis-aligned rectangle in slide points.</summary>
        private struct Box
        {
            public readonly double L, T, R, B;

            public Box(double l, double t, double r, double b) { L = l; T = t; R = r; B = b; }

            public double CenterX { get { return (L + R) / 2; } }
            public double CenterY { get { return (T + B) / 2; } }

            public double Start(bool horizontal) { return horizontal ? L : T; }
            public double End(bool horizontal) { return horizontal ? R : B; }
            public double Size(bool horizontal) { return horizontal ? R - L : B - T; }
            public double Center(bool horizontal) { return horizontal ? CenterX : CenterY; }

            /// <summary>
            /// The box PowerPoint aligns a shape by: its frame, turned with the shape's rotation,
            /// then the smallest upright rectangle around that.
            /// </summary>
            public static Box Of(PowerPoint.Shape s)
            {
                double w = s.Width, h = s.Height;
                double cx = s.Left + w / 2, cy = s.Top + h / 2;
                double a = s.Rotation * Math.PI / 180;
                double cos = Math.Abs(Math.Cos(a)), sin = Math.Abs(Math.Sin(a));
                double hw = (w * cos + h * sin) / 2, hh = (w * sin + h * cos) / 2;
                return new Box(cx - hw, cy - hh, cx + hw, cy + hh);
            }
        }
    }
}
