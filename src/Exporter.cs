using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptSelectionExport
{
    internal sealed class BackgroundOptions
    {
        /// <summary>True = no background at all (transparent).</summary>
        public bool Transparent = true;
        /// <summary>Background colour used when <see cref="Transparent"/> is false.</summary>
        public Color Color = Color.White;
    }

    /// <summary>
    /// Sets the physical resolution of a PNG file (its pHYs chunk) without touching the picture:
    /// the file is copied chunk by chunk, with the resolution put in front of the image data.
    /// </summary>
    internal static class PngResolution
    {
        private static readonly byte[] Signature = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        private static readonly uint[] CrcTable = MakeCrcTable();

        /// <summary>Copies <paramref name="source"/> to <paramref name="target"/> with the resolution set. InvalidDataException when it is not a PNG.</summary>
        public static void Copy(string source, string target, int dpi)
        {
            byte[] png = File.ReadAllBytes(source);
            if (png.Length < Signature.Length + 12) throw new InvalidDataException();
            for (int i = 0; i < Signature.Length; i++) if (png[i] != Signature[i]) throw new InvalidDataException();

            // pixels per metre, the same in both directions; unit 1 = metre
            uint perMetre = (uint)Math.Round(dpi / 0.0254);
            byte[] phys = new byte[4 + 4 + 9 + 4];
            WriteUInt(phys, 0, 9);
            phys[4] = (byte)'p'; phys[5] = (byte)'H'; phys[6] = (byte)'Y'; phys[7] = (byte)'s';
            WriteUInt(phys, 8, perMetre);
            WriteUInt(phys, 12, perMetre);
            phys[16] = 1;
            WriteUInt(phys, 17, Crc(phys, 4, 13));

            using (var result = new MemoryStream(png.Length + phys.Length))
            {
                result.Write(png, 0, Signature.Length);
                int pos = Signature.Length;
                bool stamped = false, ended = false;
                while (pos + 12 <= png.Length && !ended)
                {
                    long length = ((long)png[pos] << 24) | ((long)png[pos + 1] << 16) | ((long)png[pos + 2] << 8) | png[pos + 3];
                    if (length > png.Length - pos - 12) throw new InvalidDataException();
                    int total = 12 + (int)length;
                    bool isPhys = Is(png, pos + 4, "pHYs");
                    if (!stamped && Is(png, pos + 4, "IDAT"))       // the resolution has to come before the image data
                    {
                        result.Write(phys, 0, phys.Length);
                        stamped = true;
                    }
                    if (!isPhys) result.Write(png, pos, total);     // a resolution that was there is replaced by ours
                    ended = Is(png, pos + 4, "IEND");
                    pos += total;
                }
                if (!stamped || !ended) throw new InvalidDataException();
                File.WriteAllBytes(target, result.ToArray());
            }
        }

        private static bool Is(byte[] data, int offset, string type)
        {
            return data[offset] == type[0] && data[offset + 1] == type[1] && data[offset + 2] == type[2] && data[offset + 3] == type[3];
        }

        private static void WriteUInt(byte[] data, int offset, uint value)
        {
            data[offset] = (byte)(value >> 24); data[offset + 1] = (byte)(value >> 16); data[offset + 2] = (byte)(value >> 8); data[offset + 3] = (byte)value;
        }

        private static uint[] MakeCrcTable()
        {
            var table = new uint[256];
            for (uint n = 0; n < 256; n++)
            {
                uint c = n;
                for (int k = 0; k < 8; k++) c = (c & 1) != 0 ? 0xEDB88320u ^ (c >> 1) : c >> 1;
                table[n] = c;
            }
            return table;
        }

        private static uint Crc(byte[] data, int offset, int count)
        {
            uint c = 0xFFFFFFFFu;
            for (int i = 0; i < count; i++) c = CrcTable[(c ^ data[offset + i]) & 0xFF] ^ (c >> 8);
            return c ^ 0xFFFFFFFFu;
        }
    }

    /// <summary>
    /// Exports a range of shapes, cropped to the smallest rectangle that covers everything
    /// that is drawn for them (outlines, arrow heads and shadows included).
    /// </summary>
    internal static class Exporter
    {
        private const PowerPoint.PpShapeFormat SvgFormat = (PowerPoint.PpShapeFormat)6;   // ppShapeFormatSVG
        private const Office.MsoTriState True = Office.MsoTriState.msoTrue;
        private const Office.MsoTriState False = Office.MsoTriState.msoFalse;
        private const float MaxSlide = 4032f;   // PowerPoint's largest slide edge: 56 inches

        // ------------------------------------------------------------------ SVG

        /// <summary>Writes the shapes as an SVG file and returns its size in points.</summary>
        public static SizeD ExportSvg(PowerPoint.ShapeRange range, string path, BackgroundOptions bg)
        {
            string dir = NewTempDir();
            try
            {
                string tmp = Path.Combine(dir, "export.svg");
                range.Export(tmp, SvgFormat);
                string svg = File.ReadAllText(tmp, Encoding.UTF8);

                RectD b;
                SizeD size = SvgTools.TryReadBounds(svg, out b) ? new SizeD(b.W, b.H) : new SizeD(0, 0);
                svg = SvgTools.Finish(svg, bg.Transparent ? (Color?)null : bg.Color);
                File.WriteAllText(path, svg, new UTF8Encoding(false));
                return size;
            }
            finally { TryDeleteDir(dir); }
        }

        // ------------------------------------------------------------------ PNG / JPG

        private const int MaxImageSide = 20000;
        private const long MaxImagePixels = 120L * 1000 * 1000;

        /// <summary>
        /// Writes the shapes as a PNG or JPEG picture at the given resolution and returns its size in pixels.
        /// JPEG has no transparency, so a transparent background becomes white there.
        /// </summary>
        /// <param name="slideW">Slide width in points.</param>
        /// <param name="slideH">Slide height in points.</param>
        public static Size ExportImage(PowerPoint.ShapeRange range, float slideW, float slideH, string path,
                                       BackgroundOptions bg, int dpi, bool jpeg)
        {
            string dir = NewTempDir();
            try
            {
                // PowerPoint's own picture of the shapes: transparent, 96 pixels per inch, cropped to the visual bounds.
                string natural = Path.Combine(dir, "natural.png");
                range.Export(natural, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                Size n = ReadImageSize(natural);

                double scale = dpi / 96.0;
                int tw = Math.Max(1, (int)Math.Round(n.Width * scale)), th = Math.Max(1, (int)Math.Round(n.Height * scale));
                if (tw > MaxImageSide || th > MaxImageSide || (long)tw * th > MaxImagePixels)
                    throw new InvalidOperationException(Strings.T("err.toolarge", tw, th));

                string raw = natural;
                if (tw != n.Width || th != n.Height)
                {
                    // With ppScaleXY the two size arguments say how large the whole slide would be
                    // (the slide's own size in points = the natural 96 dpi picture), and the shapes scale with it.
                    raw = Path.Combine(dir, "scaled.png");
                    double pw = slideW * scale, ph = slideH * scale;
                    range.Export(raw, PowerPoint.PpShapeFormat.ppShapeFormatPNG,
                        Math.Max(1, (int)Math.Round(pw)), Math.Max(1, (int)Math.Round(ph)), PowerPoint.PpExportMode.ppScaleXY);

                    // Safety net in case another PowerPoint build measures those arguments differently:
                    // if the picture is clearly not the size we asked for, correct the arguments once.
                    Size got = ReadImageSize(raw);
                    double fx = (double)tw / got.Width, fy = (double)th / got.Height;
                    if (Math.Abs(fx - 1) > 0.03 || Math.Abs(fy - 1) > 0.03)
                    {
                        if (tw > 200 && th > 200 || got.Width > 200 && got.Height > 200)
                        {
                            range.Export(raw, PowerPoint.PpShapeFormat.ppShapeFormatPNG,
                                Math.Max(1, (int)Math.Round(pw * fx)), Math.Max(1, (int)Math.Round(ph * fy)), PowerPoint.PpExportMode.ppScaleXY);
                        }
                    }
                }

                // A transparent PNG is the picture PowerPoint just wrote, as it is. Only the resolution has to
                // be stamped into the file; decoding, redrawing and encoding it again would take as long as
                // PowerPoint took to make it.
                if (!jpeg && bg.Transparent)
                {
                    try
                    {
                        Size written = ReadImageSize(raw);
                        PngResolution.Copy(raw, path, dpi);
                        return written;
                    }
                    catch (InvalidDataException) { }        // not laid out like a PNG: take the long way
                }

                using (var fs = new FileStream(raw, FileMode.Open, FileAccess.Read))
                using (Image src = Image.FromStream(fs))
                {
                    bool opaque = jpeg || !bg.Transparent;
                    Color fillColor = bg.Transparent ? Color.White : Color.FromArgb(255, bg.Color);
                    using (var bmp = new Bitmap(src.Width, src.Height,
                        opaque ? System.Drawing.Imaging.PixelFormat.Format24bppRgb : System.Drawing.Imaging.PixelFormat.Format32bppArgb))
                    {
                        using (Graphics g = Graphics.FromImage(bmp))
                        {
                            if (opaque) g.Clear(fillColor);
                            else
                            {
                                g.Clear(Color.Transparent);
                                g.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                            }
                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
                            g.DrawImage(src, new Rectangle(0, 0, src.Width, src.Height), 0, 0, src.Width, src.Height, GraphicsUnit.Pixel);
                        }
                        // so that other programs place the picture at its true physical size
                        bmp.SetResolution(dpi, dpi);

                        if (jpeg)
                        {
                            System.Drawing.Imaging.ImageCodecInfo codec = null;
                            foreach (System.Drawing.Imaging.ImageCodecInfo c in System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders())
                                if (c.FormatID == System.Drawing.Imaging.ImageFormat.Jpeg.Guid) codec = c;
                            using (var ep = new System.Drawing.Imaging.EncoderParameters(1))
                            {
                                ep.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 95L);
                                bmp.Save(path, codec, ep);
                            }
                        }
                        else bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                        return bmp.Size;
                    }
                }
            }
            finally { TryDeleteDir(dir); }
        }

        private static Size ReadImageSize(string file)
        {
            if (!File.Exists(file)) throw new InvalidDataException(Strings.T("err.image"));
            using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            using (Image img = Image.FromStream(fs, false, false))
                return img.Size;
        }

        // ------------------------------------------------------------------ PDF

        /// <summary>
        /// Writes the shapes as a one-page vector PDF and returns the page size in points.
        /// Works on a temporary copy of the presentation, so the open document, its undo
        /// history and the clipboard are left untouched.
        /// </summary>
        /// <param name="isChildRange">True when <paramref name="range"/> holds shapes selected inside a group.</param>
        public static SizeD ExportPdf(PowerPoint.Application app, PowerPoint.Presentation pres, PowerPoint.Slide slide,
                                      PowerPoint.ShapeRange range, bool isChildRange, string path, BackgroundOptions bg)
        {
            // Shape ids survive the round trip through the copy; that is how we find the selection again.
            var ids = new HashSet<int>();
            int selected = range.Count;
            for (int i = 1; i <= selected; i++)
            {
                PowerPoint.Shape s = range[i];
                ids.Add(s.Id);
                if (isChildRange && s.Type == Office.MsoShapeType.msoGroup)
                {
                    PowerPoint.GroupShapes items = s.GroupItems;
                    for (int j = 1; j <= items.Count; j++) ids.Add(items[j].Id);
                }
            }
            int slideIndex = slide.SlideIndex;

            string dir = NewTempDir();
            PowerPoint.Presentation copy = null;
            try
            {
                string copyPath = Path.Combine(dir, "copy.pptx");
                pres.SaveCopyAs(copyPath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, False);
                // read-only, untitled (keeps it out of "recent files"), without a window
                copy = app.Presentations.Open(copyPath, True, True, False);
                PowerPoint.Slide cs = copy.Slides[slideIndex];

                // ---- hide everything that is not selected
                var keepTop = new List<object>();
                PowerPoint.Shape partialGroup = null;
                var keepLeaves = new List<object>();
                PowerPoint.Shapes shapes = cs.Shapes;
                int n = shapes.Count;
                for (int i = 1; i <= n; i++)
                {
                    PowerPoint.Shape sh = shapes[i];
                    if (ids.Contains(sh.Id)) { keepTop.Add(i); continue; }

                    if (isChildRange && partialGroup == null && sh.Type == Office.MsoShapeType.msoGroup)
                    {
                        PowerPoint.GroupShapes items = sh.GroupItems;
                        int cnt = items.Count;
                        var hit = new List<object>();
                        for (int j = 1; j <= cnt; j++) if (ids.Contains(items[j].Id)) hit.Add(j);
                        if (hit.Count > 0)
                        {
                            for (int j = 1; j <= cnt; j++) if (!ids.Contains(items[j].Id)) items[j].Visible = False;
                            partialGroup = sh;
                            keepLeaves = hit;
                            continue;
                        }
                    }
                    Hide(sh);
                }

                bool child = partialGroup != null;
                if (!child && keepTop.Count == 0)
                    throw new InvalidOperationException(Strings.T("err.notfound"));

                // ---- slide background and master artwork
                try { cs.DisplayMasterShapes = False; } catch { }
                cs.FollowMasterBackground = False;
                PowerPoint.FillFormat fill = cs.Background.Fill;
                if (bg.Transparent)
                {
                    fill.Visible = False;          // PowerPoint then paints no page background into the PDF
                }
                else
                {
                    fill.Visible = True;
                    fill.Solid();
                    fill.ForeColor.RGB = ColorTranslator.ToOle(bg.Color);
                    fill.Transparency = 0f;
                }

                // ---- where is the selection drawn?
                string measureSvg = Path.Combine(dir, "measure.svg");
                RectD box = Measure(cs, child, partialGroup, keepTop, keepLeaves, measureSvg);

                // PowerPoint clips the PDF to the slide. If the selection reaches outside the slide,
                // make room (enlarging a slide never rescales its content) and move it inside.
                float sw = copy.PageSetup.SlideWidth, sh2 = copy.PageSetup.SlideHeight;
                const double eps = 0.05;
                if (box.X < -eps || box.Y < -eps || box.Right > sw + eps || box.Bottom > sh2 + eps)
                {
                    float needW = Math.Min(MaxSlide, Math.Max(sw, (float)Math.Ceiling(box.W) + 4f));
                    float needH = Math.Min(MaxSlide, Math.Max(sh2, (float)Math.Ceiling(box.H) + 4f));
                    if (needW > sw) copy.PageSetup.SlideWidth = needW;
                    if (needH > sh2) copy.PageSetup.SlideHeight = needH;
                    if (needW > sw || needH > sh2)
                    {
                        sw = copy.PageSetup.SlideWidth; sh2 = copy.PageSetup.SlideHeight;
                        box = Measure(cs, child, partialGroup, keepTop, keepLeaves, measureSvg);   // content was re-centred
                    }

                    float dx = 0, dy = 0;
                    if (box.X < 0) dx = (float)(-box.X + 2);
                    else if (box.Right > sw) dx = (float)(sw - box.Right - 2);
                    if (box.Y < 0) dy = (float)(-box.Y + 2);
                    else if (box.Bottom > sh2) dy = (float)(sh2 - box.Bottom - 2);
                    if (dx != 0 || dy != 0)
                    {
                        if (child)
                        {
                            if (dx != 0) partialGroup.IncrementLeft(dx);
                            if (dy != 0) partialGroup.IncrementTop(dy);
                        }
                        else
                        {
                            PowerPoint.ShapeRange all = cs.Shapes.Range(keepTop.ToArray());
                            if (dx != 0) all.IncrementLeft(dx);
                            if (dy != 0) all.IncrementTop(dy);
                        }
                        box = Measure(cs, child, partialGroup, keepTop, keepLeaves, measureSvg);
                    }
                }

                // ---- export that one slide and crop the page to the selection
                string pdfTmp = Path.Combine(dir, "slide.pdf");
                PowerPoint.PrintRange pr = copy.PrintOptions.Ranges.Add(slideIndex, slideIndex);
                copy.ExportAsFixedFormat(pdfTmp,
                    PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                    PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint,
                    False,                                                  // FrameSlides
                    PowerPoint.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                    PowerPoint.PpPrintOutputType.ppPrintOutputSlides,
                    True,                                                   // PrintHiddenSlides (the slide itself may be hidden)
                    pr,
                    PowerPoint.PpPrintRangeType.ppPrintSlideRange,
                    "",
                    false,                                                  // IncludeDocProperties
                    false,                                                  // KeepIRMSettings
                    false,                                                  // DocStructureTags
                    true,                                                   // BitmapMissingFonts
                    false,                                                  // UseISO19005_1
                    Type.Missing);

                SizeD size;
                byte[] cropped = PdfCropper.Crop(File.ReadAllBytes(pdfTmp), box, sw, sh2, out size);
                File.WriteAllBytes(path, cropped);
                return size;
            }
            finally
            {
                if (copy != null)
                {
                    try { copy.Saved = True; copy.Close(); } catch { }
                }
                TryDeleteDir(dir);
            }
        }

        private static void Hide(PowerPoint.Shape sh)
        {
            try { sh.Visible = False; }
            catch
            {
                try { sh.Delete(); } catch { }
            }
        }

        /// <summary>Visual bounding box (slide points) of the kept shapes inside the temporary copy.</summary>
        private static RectD Measure(PowerPoint.Slide cs, bool child, PowerPoint.Shape partialGroup,
                                     List<object> keepTop, List<object> keepLeaves, string tmpSvg)
        {
            PowerPoint.ShapeRange r = child
                ? partialGroup.GroupItems.Range(keepLeaves.ToArray())
                : cs.Shapes.Range(keepTop.ToArray());
            return MeasureVisualBounds(r, tmpSvg);
        }

        /// <summary>
        /// The rectangle PowerPoint itself uses when it saves shapes as a picture.
        /// Read from the SVG it writes; an EMF header is the fallback.
        /// </summary>
        public static RectD MeasureVisualBounds(PowerPoint.ShapeRange r, string tmpSvg)
        {
            r.Export(tmpSvg, SvgFormat);
            RectD b;
            if (SvgTools.TryReadBounds(File.ReadAllText(tmpSvg, Encoding.UTF8), out b)) return b;

            string emf = Path.ChangeExtension(tmpSvg, ".emf");
            r.Export(emf, PowerPoint.PpShapeFormat.ppShapeFormatEMF);
            byte[] hdr = new byte[40];
            using (FileStream fs = File.OpenRead(emf))
            {
                if (fs.Read(hdr, 0, hdr.Length) < hdr.Length) throw new InvalidDataException(Strings.T("err.bounds"));
            }
            // ENHMETAHEADER.rclFrame: picture frame in 0.01 mm, in slide coordinates
            const double k = 72.0 / 2540.0;
            double l = BitConverter.ToInt32(hdr, 24) * k, t = BitConverter.ToInt32(hdr, 28) * k;
            double rr = BitConverter.ToInt32(hdr, 32) * k, bb = BitConverter.ToInt32(hdr, 36) * k;
            if (rr <= l || bb <= t) throw new InvalidDataException(Strings.T("err.bounds"));
            return new RectD(l, t, rr - l, bb - t);
        }

        // ------------------------------------------------------------------ temp files

        internal static string NewTempDir()
        {
            string dir = Path.Combine(Path.GetTempPath(), "PptSelectionExport", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            return dir;
        }

        internal static void TryDeleteDir(string dir)
        {
            try { Directory.Delete(dir, true); } catch { }
        }
    }
}
