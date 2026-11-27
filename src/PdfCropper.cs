using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace PptSelectionExport
{
    /// <summary>
    /// Crops a one-page PDF written by PowerPoint to a sub-rectangle of the slide.
    ///
    /// PowerPoint cannot make a page smaller than 1 inch and rounds page sizes, so instead of
    /// resizing the slide we export the whole slide and then append a standard PDF
    /// "incremental update": a new version of the page object whose MediaBox is the wanted
    /// rectangle (moved to the origin by a one-line content stream placed in front of the
    /// original page content). Nothing in the original bytes is modified.
    /// </summary>
    internal static class PdfCropper
    {
        private static readonly Encoding Latin1 = Encoding.GetEncoding(28591);

        private static string F(double v)
        {
            if (Math.Abs(v) < 0.00005) v = 0;
            return v.ToString("0.####", CultureInfo.InvariantCulture);
        }

        private static double D(string s)
        {
            return double.Parse(s, NumberStyles.Float, CultureInfo.InvariantCulture);
        }

        /// <param name="pdf">The PDF PowerPoint produced for one slide.</param>
        /// <param name="crop">Rectangle to keep, in slide points (origin top-left).</param>
        /// <param name="slideW">Slide width in points.</param>
        /// <param name="slideH">Slide height in points.</param>
        /// <param name="finalSize">Resulting page size in points.</param>
        public static byte[] Crop(byte[] pdf, RectD crop, double slideW, double slideH, out SizeD finalSize)
        {
            string s = Latin1.GetString(pdf);

            // ---- trailer of the newest cross-reference section
            MatchCollection sx = Regex.Matches(s, @"startxref\s+(\d+)\s+%%EOF");
            if (sx.Count == 0) throw new InvalidDataException(Strings.T("err.pdf", "startxref"));
            Match lastSx = sx[sx.Count - 1];
            string prev = lastSx.Groups[1].Value;

            int tr = s.LastIndexOf("trailer", lastSx.Index, StringComparison.Ordinal);
            if (tr < 0) throw new InvalidDataException(Strings.T("err.pdf", "trailer"));
            string trailer = s.Substring(tr, lastSx.Index - tr);

            Match mSize = Regex.Match(trailer, @"/Size\s+(\d+)");
            Match mRoot = Regex.Match(trailer, @"/Root\s+(\d+\s+\d+\s+R)");
            if (!mSize.Success || !mRoot.Success) throw new InvalidDataException(Strings.T("err.pdf", "trailer fields"));
            Match mInfo = Regex.Match(trailer, @"/Info\s+(\d+\s+\d+\s+R)");
            Match mId = Regex.Match(trailer, @"/ID\s*(\[\s*<[0-9A-Fa-f]*>\s*<[0-9A-Fa-f]*>\s*\])");
            int newNum = int.Parse(mSize.Groups[1].Value, CultureInfo.InvariantCulture);

            // ---- the (single) page object
            Match pm = Regex.Match(s, @"/Type\s*/Page(?![A-Za-z])");
            if (!pm.Success) throw new InvalidDataException(Strings.T("err.pdf", "page object"));
            if (pm.NextMatch().Success) throw new InvalidDataException(Strings.T("err.pdf", "more than one page"));

            Match om = new Regex(@"(\d+)\s+(\d+)\s+obj", RegexOptions.RightToLeft).Match(s, pm.Index);
            int objEnd = s.IndexOf("endobj", pm.Index, StringComparison.Ordinal);
            if (!om.Success || objEnd < 0) throw new InvalidDataException(Strings.T("err.pdf", "page object bounds"));
            int pageNum = int.Parse(om.Groups[1].Value, CultureInfo.InvariantCulture);
            int pageGen = int.Parse(om.Groups[2].Value, CultureInfo.InvariantCulture);
            string body = s.Substring(om.Index + om.Length, objEnd - (om.Index + om.Length)).Trim();
            if (body.IndexOf("stream", StringComparison.Ordinal) >= 0 || !body.StartsWith("<<", StringComparison.Ordinal))
                throw new InvalidDataException(Strings.T("err.pdf", "page dictionary"));

            // ---- current page box
            const string boxRx = @"\s*\[\s*(-?[0-9.]+)\s+(-?[0-9.]+)\s+(-?[0-9.]+)\s+(-?[0-9.]+)\s*\]";
            Match mb = Regex.Match(body, "/MediaBox" + boxRx);
            double mx0 = 0, my0 = 0, mx1 = slideW, my1 = slideH;
            if (mb.Success)
            {
                mx0 = D(mb.Groups[1].Value); my0 = D(mb.Groups[2].Value);
                mx1 = D(mb.Groups[3].Value); my1 = D(mb.Groups[4].Value);
            }
            double kx = (mx1 - mx0) / slideW, ky = (my1 - my0) / slideH;

            // slide coordinates (y down) -> PDF coordinates (y up), clamped to the page
            double cx0 = Math.Max(mx0, mx0 + crop.X * kx);
            double cx1 = Math.Min(mx1, mx0 + crop.Right * kx);
            double cy1 = Math.Min(my1, my1 - crop.Y * ky);
            double cy0 = Math.Max(my0, my1 - crop.Bottom * ky);
            double w = cx1 - cx0, h = cy1 - cy0;
            if (w <= 0.01 || h <= 0.01) throw new InvalidOperationException(Strings.T("err.outside"));
            finalSize = new SizeD(w, h);

            // ---- new version of the page dictionary
            string newBox = "/MediaBox[0 0 " + F(w) + " " + F(h) + "]";
            string nb = body;
            nb = Regex.Replace(nb, "/(?:CropBox|TrimBox|BleedBox|ArtBox)" + boxRx, "");
            if (mb.Success) nb = Regex.Replace(nb, "/MediaBox" + boxRx, newBox);
            else nb = "<<" + newBox + nb.Substring(2);
            // link rectangles would no longer line up with the moved content
            nb = Regex.Replace(nb, @"/Annots\s*\[[^\]]*\]", "");
            nb = Regex.Replace(nb, @"/Annots\s+\d+\s+\d+\s+R", "");

            string prefixRef = newNum.ToString(CultureInfo.InvariantCulture) + " 0 R";
            Match single = Regex.Match(nb, @"/Contents\s+(\d+\s+\d+\s+R)");
            bool shifted = true;
            if (single.Success)
                nb = nb.Substring(0, single.Index) + "/Contents[" + prefixRef + " " + single.Groups[1].Value + "]" + nb.Substring(single.Index + single.Length);
            else
            {
                Match arr = Regex.Match(nb, @"/Contents\s*\[");
                if (arr.Success) nb = nb.Substring(0, arr.Index + arr.Length) + prefixRef + " " + nb.Substring(arr.Index + arr.Length);
                else shifted = false;   // empty page: nothing to move
            }

            // ---- append the update
            const string nl = "\r\n";
            var sb = new StringBuilder();
            if (!s.EndsWith("\n", StringComparison.Ordinal) && !s.EndsWith("\r", StringComparison.Ordinal)) sb.Append(nl);

            long offPage = pdf.Length + sb.Length;
            sb.Append(pageNum).Append(' ').Append(pageGen).Append(" obj").Append(nl).Append(nb).Append(nl).Append("endobj").Append(nl);

            long offPrefix = 0;
            if (shifted)
            {
                offPrefix = pdf.Length + sb.Length;
                string content = "1 0 0 1 " + F(-cx0) + " " + F(-cy0) + " cm\n";
                sb.Append(newNum).Append(" 0 obj").Append(nl)
                  .Append("<</Length ").Append(content.Length).Append(">>").Append(nl)
                  .Append("stream").Append(nl).Append(content).Append(nl).Append("endstream").Append(nl)
                  .Append("endobj").Append(nl);
            }

            long offXref = pdf.Length + sb.Length;
            sb.Append("xref").Append(nl);
            sb.Append(pageNum).Append(" 1").Append(nl)
              .Append(offPage.ToString("D10", CultureInfo.InvariantCulture)).Append(' ')
              .Append(pageGen.ToString("D5", CultureInfo.InvariantCulture)).Append(" n").Append(nl);
            if (shifted)
            {
                sb.Append(newNum).Append(" 1").Append(nl)
                  .Append(offPrefix.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n").Append(nl);
            }
            sb.Append("trailer").Append(nl);
            sb.Append("<</Size ").Append(shifted ? newNum + 1 : newNum).Append("/Root ").Append(mRoot.Groups[1].Value);
            if (mInfo.Success) sb.Append("/Info ").Append(mInfo.Groups[1].Value);
            if (mId.Success) sb.Append("/ID").Append(mId.Groups[1].Value);
            sb.Append("/Prev ").Append(prev).Append(">>").Append(nl);
            sb.Append("startxref").Append(nl).Append(offXref).Append(nl).Append("%%EOF").Append(nl);

            byte[] tail = Latin1.GetBytes(sb.ToString());
            byte[] result = new byte[pdf.Length + tail.Length];
            Buffer.BlockCopy(pdf, 0, result, 0, pdf.Length);
            Buffer.BlockCopy(tail, 0, result, pdf.Length, tail.Length);
            return result;
        }
    }

    internal struct SizeD
    {
        public double W, H;
        public SizeD(double w, double h) { W = w; H = h; }
    }
}
