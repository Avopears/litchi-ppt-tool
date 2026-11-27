using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace PptSelectionExport
{
    /// <summary>A rectangle in slide points (origin = top-left corner of the slide).</summary>
    internal struct RectD
    {
        public double X, Y, W, H;
        public RectD(double x, double y, double w, double h) { X = x; Y = y; W = w; H = h; }
        public double Right { get { return X + W; } }
        public double Bottom { get { return Y + H; } }
        public override string ToString()
        {
            return string.Format(CultureInfo.InvariantCulture, "[x={0:0.###} y={1:0.###} w={2:0.###} h={3:0.###}]", X, Y, W, H);
        }
    }

    /// <summary>
    /// Helpers around the SVG that PowerPoint writes for ShapeRange.Export.
    /// That SVG is in CSS pixels (96 per inch) and looks like
    ///   &lt;svg width="495" height="273" ...&gt;&lt;defs&gt;...&lt;/defs&gt;&lt;g transform="translate(-132 -132)"&gt;...
    /// i.e. the drawing keeps slide coordinates and the root group shifts the
    /// visual bounding box of the exported shapes to the origin.
    /// </summary>
    internal static class SvgTools
    {
        private const double PxToPt = 72.0 / 96.0;

        private static readonly Regex TranslateRx = new Regex(
            @"^\s*translate\(\s*(-?[0-9.eE+-]+)(?:[\s,]+(-?[0-9.eE+-]+))?\s*\)\s*$", RegexOptions.CultureInvariant);

        private static string Attr(string tag, string name)
        {
            Match m = Regex.Match(tag, @"\s" + name + "\\s*=\\s*\"([^\"]*)\"", RegexOptions.CultureInvariant);
            return m.Success ? m.Groups[1].Value : null;
        }

        private static bool Num(string s, out double v)
        {
            v = 0;
            if (s == null) return false;
            s = s.Trim();
            if (s.EndsWith("px", StringComparison.OrdinalIgnoreCase)) s = s.Substring(0, s.Length - 2);
            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v);
        }

        private static int SkipWs(string s, int p)
        {
            while (p < s.Length && char.IsWhiteSpace(s[p])) p++;
            return p;
        }

        /// <summary>
        /// Reads the visual bounding box (in slide points) from an SVG produced by ShapeRange.Export.
        /// Returns false when the file does not have the expected structure.
        /// </summary>
        public static bool TryReadBounds(string svg, out RectD bounds)
        {
            bounds = new RectD();
            int i = svg.IndexOf("<svg", StringComparison.Ordinal);
            if (i < 0) return false;
            int rootEnd = svg.IndexOf('>', i);
            if (rootEnd < 0) return false;
            string root = svg.Substring(i, rootEnd - i + 1);

            double w, h;
            if (!Num(Attr(root, "width"), out w) || !Num(Attr(root, "height"), out h)) return false;
            if (w <= 0 || h <= 0) return false;

            int p = SkipWs(svg, rootEnd + 1);
            if (string.CompareOrdinal(svg, p, "<defs", 0, 5) == 0)
            {
                int tagEnd = svg.IndexOf('>', p);
                if (tagEnd < 0) return false;
                if (svg[tagEnd - 1] == '/') p = tagEnd + 1;
                else
                {
                    int e = svg.IndexOf("</defs>", tagEnd, StringComparison.Ordinal);
                    if (e < 0) return false;
                    p = e + "</defs>".Length;
                }
                p = SkipWs(svg, p);
            }

            if (string.CompareOrdinal(svg, p, "<g", 0, 2) != 0) return false;
            int gEnd = svg.IndexOf('>', p);
            if (gEnd < 0) return false;
            string g = svg.Substring(p, gEnd - p + 1);

            double tx = 0, ty = 0;
            string tr = Attr(g, "transform");
            if (tr != null)
            {
                Match m = TranslateRx.Match(tr);
                if (!m.Success) return false;
                if (!Num(m.Groups[1].Value, out tx)) return false;
                if (m.Groups[2].Success && !Num(m.Groups[2].Value, out ty)) return false;
            }

            bounds = new RectD(-tx * PxToPt, -ty * PxToPt, w * PxToPt, h * PxToPt);
            return true;
        }

        /// <summary>
        /// Post-processes PowerPoint's SVG: adds a viewBox (so the picture scales cleanly)
        /// and, when a background colour is requested, a full-size rectangle behind everything.
        /// A null colour keeps the background transparent.
        /// </summary>
        public static string Finish(string svg, Color? background)
        {
            int i = svg.IndexOf("<svg", StringComparison.Ordinal);
            if (i < 0) throw new InvalidDataException(Strings.T("err.svg"));
            int rootEnd = svg.IndexOf('>', i);
            if (rootEnd < 0) throw new InvalidDataException(Strings.T("err.svg"));
            string root = svg.Substring(i, rootEnd - i + 1);

            string ws = Attr(root, "width"), hs = Attr(root, "height");
            double w = 0, h = 0;
            bool sized = Num(ws, out w) & Num(hs, out h);

            var sb = new StringBuilder(svg.Length + 200);
            sb.Append(svg, 0, rootEnd);
            if (sized && Attr(root, "viewBox") == null)
            {
                sb.AppendFormat(CultureInfo.InvariantCulture, " viewBox=\"0 0 {0} {1}\"", F(w), F(h));
            }
            sb.Append('>');
            if (background.HasValue)
            {
                Color c = background.Value;
                string fill = string.Format(CultureInfo.InvariantCulture, "#{0:X2}{1:X2}{2:X2}", c.R, c.G, c.B);
                if (sized)
                    sb.AppendFormat(CultureInfo.InvariantCulture, "<rect x=\"0\" y=\"0\" width=\"{0}\" height=\"{1}\" fill=\"{2}\"/>", F(w), F(h), fill);
                else
                    sb.AppendFormat(CultureInfo.InvariantCulture, "<rect x=\"0\" y=\"0\" width=\"100%\" height=\"100%\" fill=\"{0}\"/>", fill);
            }
            sb.Append(svg, rootEnd + 1, svg.Length - rootEnd - 1);
            return sb.ToString();
        }

        private static string F(double v)
        {
            return v.ToString("0.####", CultureInfo.InvariantCulture);
        }
    }
}
