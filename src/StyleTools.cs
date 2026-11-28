using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptSelectionExport
{
    /// <summary>The dash styles offered, in the order of PowerPoint's own "Dashes" menu.</summary>
    internal enum DashKind { Solid, RoundDot, SquareDot, Dash, DashDot, LongDash, LongDashDot, LongDashDotDot }

    /// <summary>
    /// Fill colour, outline colour, outline weight and dash style of the selected shapes.
    /// Each method returns how many shapes it changed; shapes a setting does not apply to
    /// (a fill on a straight line, an outline on a table) are skipped.
    /// </summary>
    internal static class StyleTools
    {
        /// <summary>Weights offered in the list, in points (the same steps as PowerPoint's own list).</summary>
        public static readonly float[] Weights = { 0.25f, 0.5f, 0.75f, 1f, 1.5f, 2.25f, 3f, 4.5f, 6f };

        public static readonly DashKind[] Dashes =
        {
            DashKind.Solid, DashKind.RoundDot, DashKind.SquareDot, DashKind.Dash,
            DashKind.DashDot, DashKind.LongDash, DashKind.LongDashDot, DashKind.LongDashDotDot
        };

        public const float MinWeight = 0.05f, MaxWeight = 100f;

        private const Office.MsoTriState True = Office.MsoTriState.msoTrue, False = Office.MsoTriState.msoFalse;

        /// <summary>Solid fill in the given colour, or no fill when <paramref name="color"/> is null.</summary>
        public static int SetFill(PowerPoint.ShapeRange range, Color? color)
        {
            int changed = 0;
            foreach (PowerPoint.Shape s in Leaves(range))
            {
                try
                {
                    if (IsLine(s)) continue;            // a line has nothing to fill
                    PowerPoint.FillFormat f = s.Fill;
                    if (color.HasValue)
                    {
                        f.Visible = True;
                        f.Solid();
                        f.ForeColor.RGB = ColorTranslator.ToOle(color.Value);
                    }
                    else f.Visible = False;
                    changed++;
                }
                catch { }
            }
            return changed;
        }

        /// <summary>Outline (or line) in the given colour, or no outline when <paramref name="color"/> is null.</summary>
        public static int SetOutlineColor(PowerPoint.ShapeRange range, Color? color)
        {
            int changed = 0;
            foreach (PowerPoint.Shape s in Leaves(range))
            {
                try
                {
                    PowerPoint.LineFormat l = s.Line;
                    if (color.HasValue)
                    {
                        l.Visible = True;
                        l.ForeColor.RGB = ColorTranslator.ToOle(color.Value);
                    }
                    else l.Visible = False;
                    changed++;
                }
                catch { }
            }
            return changed;
        }

        /// <summary>Outline weight in points. A shape without an outline gets one, as in PowerPoint.</summary>
        public static int SetWeight(PowerPoint.ShapeRange range, float points)
        {
            points = Math.Max(MinWeight, Math.Min(MaxWeight, points));
            int changed = 0;
            foreach (PowerPoint.Shape s in Leaves(range))
            {
                try
                {
                    PowerPoint.LineFormat l = s.Line;
                    l.Visible = True;
                    l.Weight = points;
                    changed++;
                }
                catch { }
            }
            return changed;
        }

        public static int SetDash(PowerPoint.ShapeRange range, DashKind dash)
        {
            int changed = 0;
            foreach (PowerPoint.Shape s in Leaves(range))
            {
                try
                {
                    PowerPoint.LineFormat l = s.Line;
                    l.Visible = True;
                    l.DashStyle = ToMso(dash);
                    changed++;
                }
                catch { }
            }
            return changed;
        }

        /// <summary>The shapes themselves, with every group replaced by the shapes inside it.</summary>
        internal static List<PowerPoint.Shape> Leaves(PowerPoint.ShapeRange range)
        {
            var list = new List<PowerPoint.Shape>();
            int n = range.Count;
            for (int i = 1; i <= n; i++)
            {
                PowerPoint.Shape s = range[i];
                bool group = false;
                try { group = s.Type == Office.MsoShapeType.msoGroup; } catch { }
                if (!group) { list.Add(s); continue; }
                PowerPoint.GroupShapes items = s.GroupItems;      // already flat: nested groups are expanded
                for (int j = 1; j <= items.Count; j++) list.Add(items[j]);
            }
            return list;
        }

        private static bool IsLine(PowerPoint.Shape s)
        {
            return s.Type == Office.MsoShapeType.msoLine || s.Connector == True;
        }

        internal static Office.MsoLineDashStyle ToMso(DashKind dash)
        {
            switch (dash)
            {
                case DashKind.RoundDot: return Office.MsoLineDashStyle.msoLineRoundDot;
                case DashKind.SquareDot: return Office.MsoLineDashStyle.msoLineSquareDot;
                case DashKind.Dash: return Office.MsoLineDashStyle.msoLineDash;
                case DashKind.DashDot: return Office.MsoLineDashStyle.msoLineDashDot;
                case DashKind.LongDash: return Office.MsoLineDashStyle.msoLineLongDash;
                case DashKind.LongDashDot: return Office.MsoLineDashStyle.msoLineLongDashDot;
                case DashKind.LongDashDotDot: return Office.MsoLineDashStyle.msoLineLongDashDotDot;
                default: return Office.MsoLineDashStyle.msoLineSolid;
            }
        }

        /// <summary>"1.5", "1,5", "1.5 pt" or "1.5 磅" -> 1.5. False when it is not a weight between the limits.</summary>
        public static bool TryParseWeight(string text, out float points)
        {
            points = 0;
            if (text == null) return false;
            string t = text.Trim().ToLowerInvariant().Replace("磅", "").Replace("pt", "").Replace(',', '.').Trim();
            float v;
            if (!float.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out v)) return false;
            if (float.IsNaN(v) || v < MinWeight || v > MaxWeight) return false;
            points = (float)Math.Round(v, 2);
            return true;
        }

        public static string FormatWeight(float points)
        {
            return points.ToString("0.##", CultureInfo.InvariantCulture) + (Strings.Chinese ? " 磅" : " pt");
        }
    }
}
