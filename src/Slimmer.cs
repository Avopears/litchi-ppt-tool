using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using ImageEncoder = System.Drawing.Imaging.Encoder;

namespace PptSelectionExport
{
    internal sealed class SlimOptions
    {
        /// <summary>Pixels per inch a picture keeps at the size it is shown on the slide; 0 = pictures keep all their pixels.</summary>
        public int Ppi = 220;
        /// <summary>JPEG quality, 1 to 100.</summary>
        public int Quality = 80;
        /// <summary>Whether photos stored losslessly (PNG, BMP, TIFF) and without transparency may become JPEG.</summary>
        public bool PhotosToJpeg = true;
    }

    /// <summary>One picture file inside the presentation, before and after.</summary>
    internal sealed class SlimImage
    {
        public string Name;             // "image3.png"
        public string NewName;          // differs from Name when the format changed
        public long OldBytes, NewBytes;
        public Size OldPixels, NewPixels;
        public bool Changed;
        /// <summary>Why it was left as it is: a key of Strings, or null.</summary>
        public string Note;
    }

    internal sealed class SlimReport
    {
        public long OldFileBytes, NewFileBytes;
        public long OldImageBytes, NewImageBytes;
        public readonly List<SlimImage> Images = new List<SlimImage>();

        public int ChangedCount
        {
            get { int n = 0; foreach (SlimImage i in Images) if (i.Changed) n++; return n; }
        }
    }

    /// <summary>What takes up the space in a presentation file.</summary>
    internal sealed class SlimOverview
    {
        public long FileBytes, ImageBytes, MediaBytes;      // media = video and audio
        public int ImageCount, MediaCount;
        public long OtherBytes { get { return Math.Max(0, FileBytes - ImageBytes - MediaBytes); } }
    }

    /// <summary>
    /// Makes a smaller copy of a presentation file by making its pictures smaller.
    ///
    /// A .pptx is a zip package and nearly all of its size is the pictures under ppt/media. Each
    /// one is resampled to the number of pixels it needs at the size it is shown (the largest
    /// size, when it is used more than once; a cropped picture keeps the pixels of its hidden
    /// part in proportion) and encoded again: JPEG at the chosen quality, PNG where there is
    /// transparency or where JPEG would not pay. A picture is only replaced when that makes it
    /// smaller. Everything else in the package is copied as it is.
    /// </summary>
    internal static class Slimmer
    {
        private const string NsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string NsRel = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string NsTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
        private const string ContentTypes = "[Content_Types].xml";
        private const double Emu = 914400.0;            // per inch
        private const int TooSmallToBother = 8 * 1024;  // bytes

        /// <summary>The picture formats this can read and write.</summary>
        private static readonly HashSet<string> Raster = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "png", "jpg", "jpeg", "jpe", "jfif", "bmp", "dib", "tif", "tiff" };
        /// <summary>Other things under ppt/media that are pictures (left alone).</summary>
        private static readonly HashSet<string> OtherPictures = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "gif", "emf", "wmf", "svg", "wdp", "ico", "webp", "heic", "avif", "jxr" };

        // ------------------------------------------------------------------ overview

        public static SlimOverview Inspect(string file)
        {
            var o = new SlimOverview { FileBytes = new FileInfo(file).Length };
            using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                foreach (ZipArchiveEntry e in zip.Entries)
                {
                    if (!IsMedia(e.FullName)) continue;
                    string ext = Ext(e.FullName);
                    if (Raster.Contains(ext) || OtherPictures.Contains(ext)) { o.ImageCount++; o.ImageBytes += e.CompressedLength; }
                    else { o.MediaCount++; o.MediaBytes += e.CompressedLength; }
                }
            }
            return o;
        }

        // ------------------------------------------------------------------ the whole job

        /// <summary>
        /// Writes a slimmed copy of <paramref name="input"/> to <paramref name="output"/>.
        /// <paramref name="progress"/> is told (pictures done, pictures in all); when
        /// <paramref name="cancelled"/> says so the job stops with an OperationCanceledException
        /// and leaves no output file. Safe to run on a background thread.
        /// </summary>
        public static SlimReport Slim(string input, string output, SlimOptions options, Action<int, int> progress, Func<bool> cancelled)
        {
            var report = new SlimReport { OldFileBytes = new FileInfo(input).Length };
            string work = output + ".work";
            Directory.CreateDirectory(work);
            bool done = false;
            try
            {
                using (var inFs = new FileStream(input, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var zip = new ZipArchive(inFs, ZipArchiveMode.Read))
                {
                    Dictionary<string, Need> needs = MeasureUses(zip);

                    // ---- the pictures, a few at a time
                    var pictures = new List<ZipArchiveEntry>();
                    foreach (ZipArchiveEntry e in zip.Entries)
                    {
                        if (!IsMedia(e.FullName)) continue;
                        string ext = Ext(e.FullName);
                        if (Raster.Contains(ext)) pictures.Add(e);
                        else if (OtherPictures.Contains(ext))
                            report.Images.Add(new SlimImage { Name = FileName(e.FullName), NewName = FileName(e.FullName), OldBytes = e.Length, NewBytes = e.Length, Note = "slim.note.format" });
                    }

                    var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (ZipArchiveEntry e in zip.Entries) names.Add(e.FullName);
                    var replaced = new Dictionary<string, Replacement>(StringComparer.Ordinal);     // by the old part name

                    // a decoded picture can take hundreds of megabytes: one at a time in a 32-bit Office
                    int degree = IntPtr.Size >= 8 ? Math.Max(1, Math.Min(4, Environment.ProcessorCount)) : 1;
                    for (int start = 0; start < pictures.Count; start += degree)
                    {
                        if (cancelled != null && cancelled()) throw new OperationCanceledException();
                        if (progress != null) progress(start, pictures.Count);

                        int count = Math.Min(degree, pictures.Count - start);
                        var data = new byte[count][];
                        var infos = new SlimImage[count];
                        var results = new Encoded[count];
                        for (int i = 0; i < count; i++)                    // the zip is read from one thread only
                        {
                            ZipArchiveEntry e = pictures[start + i];
                            data[i] = ReadAll(e);
                            infos[i] = new SlimImage { Name = FileName(e.FullName), NewName = FileName(e.FullName), OldBytes = data[i].Length, NewBytes = data[i].Length };
                        }
                        int first = start;
                        Parallel.For(0, count, new ParallelOptions { MaxDegreeOfParallelism = degree }, i =>
                        {
                            ZipArchiveEntry e = pictures[first + i];
                            Need need;
                            needs.TryGetValue(e.FullName, out need);
                            try { results[i] = Recompress(data[i], need, options, infos[i]); }
                            catch { infos[i].Note = "slim.note.error"; }      // GDI+ could not read it: it stays as it is
                        });
                        for (int i = 0; i < count; i++)
                        {
                            ZipArchiveEntry e = pictures[start + i];
                            SlimImage info = infos[i];
                            report.Images.Add(info);
                            if (results[i] == null) continue;

                            string newName = e.FullName;
                            if (!string.Equals(Ext(e.FullName), results[i].Ext, StringComparison.OrdinalIgnoreCase) && !SameFormat(Ext(e.FullName), results[i].Ext))
                            {
                                string stem = e.FullName.Substring(0, e.FullName.LastIndexOf('.'));
                                newName = stem + "." + results[i].Ext;
                                for (int n = 2; names.Contains(newName); n++) newName = stem + "_" + n + "." + results[i].Ext;
                                names.Add(newName);
                            }
                            string temp = Path.Combine(work, replaced.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".bin");
                            File.WriteAllBytes(temp, results[i].Data);
                            replaced[e.FullName] = new Replacement { NewName = newName, TempFile = temp };
                            info.Changed = true;
                            info.NewName = FileName(newName);
                            info.NewBytes = results[i].Data.Length;
                            info.NewPixels = results[i].Pixels;
                        }
                    }
                    if (cancelled != null && cancelled()) throw new OperationCanceledException();
                    if (progress != null) progress(pictures.Count, pictures.Count);

                    WritePackage(zip, output, replaced);
                }

                foreach (SlimImage i in report.Images) { report.OldImageBytes += i.OldBytes; report.NewImageBytes += i.NewBytes; }
                report.NewFileBytes = new FileInfo(output).Length;
                done = true;
                return report;
            }
            finally
            {
                try { Directory.Delete(work, true); } catch { }
                if (!done) { try { File.Delete(output); } catch { } }
            }
        }

        private sealed class Replacement { public string NewName, TempFile; }

        // ------------------------------------------------------------------ how large is each picture shown?

        /// <summary>Where and how large one picture file is shown.</summary>
        private sealed class Need
        {
            /// <summary>The size of the whole picture on the page, in inches, once for every place it is used.</summary>
            public readonly List<SizeF> Inches = new List<SizeF>();
            /// <summary>Used in a way this code cannot measure (a tiled texture, an unknown part): keep every pixel.</summary>
            public bool KeepPixels;
            /// <summary>
            /// Shown with "Set Transparent Color", which makes one exact colour see-through. Resampling
            /// or JPEG shifts colours a little, and the picture would get a speckled fringe: leave the file alone.
            /// </summary>
            public bool LeaveAlone;
        }

        private sealed class Rel { public string Id, Target; }

        private static Dictionary<string, Need> MeasureUses(ZipArchive zip)
        {
            SizeF slide = new SizeF(10f, 7.5f), notes = new SizeF(7.5f, 10f);
            ZipArchiveEntry presEntry = zip.GetEntry("ppt/presentation.xml");
            if (presEntry != null)
            {
                XmlDocument pres = LineCaps.LoadXml(presEntry);
                XmlNamespaceManager pnm = LineCaps.Namespaces(pres);
                slide = PageSize(pres.SelectSingleNode("//p:sldSz", pnm) as XmlElement, slide);
                notes = PageSize(pres.SelectSingleNode("//p:notesSz", pnm) as XmlElement, notes);
            }

            var needs = new Dictionary<string, Need>(StringComparer.Ordinal);
            foreach (ZipArchiveEntry relsEntry in zip.Entries)
            {
                if (!relsEntry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
                string source = SourceOf(relsEntry.FullName);
                var images = new List<Rel>();
                foreach (XmlElement rel in Relationships(LineCaps.LoadXml(relsEntry)))
                {
                    if (rel.GetAttribute("TargetMode") == "External") continue;
                    if (!rel.GetAttribute("Type").EndsWith("/image", StringComparison.Ordinal)) continue;
                    images.Add(new Rel { Id = rel.GetAttribute("Id"), Target = Resolve(source, rel.GetAttribute("Target")) });
                }
                if (images.Count == 0) continue;

                XmlDocument doc = null;
                XmlNamespaceManager nm = null;
                ZipArchiveEntry part = zip.GetEntry(source);
                if (part != null && source.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        doc = LineCaps.LoadXml(part);
                        nm = LineCaps.Namespaces(doc);
                        nm.AddNamespace("r", NsR);
                    }
                    catch { doc = null; }
                }
                bool onNotes = source.StartsWith("ppt/notes", StringComparison.Ordinal) || source.StartsWith("ppt/handout", StringComparison.Ordinal);
                SizeF page = onNotes ? notes : slide;

                foreach (Rel rel in images)
                {
                    Need need;
                    if (!needs.TryGetValue(rel.Target, out need)) needs[rel.Target] = need = new Need();
                    XmlNodeList blips = doc == null ? null : doc.SelectNodes("//a:blip[@r:embed='" + rel.Id + "']", nm);
                    if (blips == null || blips.Count == 0) { need.KeepPixels = true; continue; }     // referenced in some other way
                    foreach (XmlNode blip in blips)
                    {
                        if (blip.SelectSingleNode("a:clrChange", nm) != null) need.LeaveAlone = true;
                        SizeF? size = Measure((XmlElement)blip, page, nm);
                        if (size.HasValue) need.Inches.Add(size.Value);
                        else need.KeepPixels = true;
                    }
                }
            }
            return needs;
        }

        private static SizeF PageSize(XmlElement e, SizeF fallback)
        {
            long cx, cy;
            if (e == null || !long.TryParse(e.GetAttribute("cx"), out cx) || !long.TryParse(e.GetAttribute("cy"), out cy) || cx <= 0 || cy <= 0) return fallback;
            return new SizeF((float)(cx / Emu), (float)(cy / Emu));
        }

        /// <summary>
        /// The size, in inches, that the whole picture has where this a:blip shows it; null when it
        /// is not stretched over an area (a tiled texture). Where the area cannot be told (a
        /// bullet, a table cell, a chart), the page is taken: nothing is shown larger than that.
        /// </summary>
        private static SizeF? Measure(XmlElement blip, SizeF page, XmlNamespaceManager nm)
        {
            var fill = blip.ParentNode as XmlElement;
            if (fill == null || fill.LocalName != "blipFill") return page;
            if (fill.SelectSingleNode("a:tile", nm) != null) return null;

            // the object that shows it: a picture, or a shape filled with it
            XmlElement owner = null;
            for (XmlNode n = fill.ParentNode; n is XmlElement; n = n.ParentNode)
            {
                var e = (XmlElement)n;
                if (e.NamespaceURI != LineCaps.NsP) continue;
                if (e.LocalName == "pic") { if (fill.ParentNode == e) owner = e; break; }
                if (e.LocalName == "sp" || e.LocalName == "cxnSp")
                {
                    if (fill.ParentNode.LocalName == "spPr" && fill.ParentNode.ParentNode == e) owner = e;      // not a fill of its text
                    break;
                }
                if (e.LocalName == "bg" || e.LocalName == "graphicFrame" || e.LocalName == "spTree") break;
            }
            if (owner == null) return page;

            double w, h;
            if (!Extent(owner.SelectSingleNode("p:spPr/a:xfrm/a:ext", nm) as XmlElement, out w, out h)) return page;    // a placeholder that takes its size from the layout

            // a group can scale what is inside it
            for (XmlNode n = owner.ParentNode; n is XmlElement; n = n.ParentNode)
            {
                var g = (XmlElement)n;
                if (g.LocalName != "grpSp" || g.NamespaceURI != LineCaps.NsP) continue;
                double gw, gh, cw, ch;
                if (Extent(g.SelectSingleNode("p:grpSpPr/a:xfrm/a:ext", nm) as XmlElement, out gw, out gh)
                    && Extent(g.SelectSingleNode("p:grpSpPr/a:xfrm/a:chExt", nm) as XmlElement, out cw, out ch))
                {
                    w *= gw / cw;
                    h *= gh / ch;
                }
            }

            // the picture can be laid over a part of the area only (fillRect), and show a part of itself only (srcRect)
            double l, t, r, b;
            Insets(fill.SelectSingleNode("a:stretch/a:fillRect", nm) as XmlElement, out l, out t, out r, out b);
            if (1 - l - r > 0.01) w *= 1 - l - r;
            if (1 - t - b > 0.01) h *= 1 - t - b;
            Insets(fill.SelectSingleNode("a:srcRect", nm) as XmlElement, out l, out t, out r, out b);
            w /= Math.Max(0.02, 1 - l - r);
            h /= Math.Max(0.02, 1 - t - b);
            return new SizeF((float)(w / Emu), (float)(h / Emu));
        }

        private static bool Extent(XmlElement e, out double cx, out double cy)
        {
            cx = cy = 0;
            long x, y;
            if (e == null || !long.TryParse(e.GetAttribute("cx"), out x) || !long.TryParse(e.GetAttribute("cy"), out y) || x <= 0 || y <= 0) return false;
            cx = x; cy = y;
            return true;
        }

        /// <summary>l, t, r, b of a relative rectangle, as fractions (the file has thousandths of a percent).</summary>
        private static void Insets(XmlElement e, out double l, out double t, out double r, out double b)
        {
            l = Fraction(e, "l"); t = Fraction(e, "t"); r = Fraction(e, "r"); b = Fraction(e, "b");
        }

        private static double Fraction(XmlElement e, string name)
        {
            long v;
            return e != null && long.TryParse(e.GetAttribute(name), out v) ? v / 100000.0 : 0;
        }

        // ------------------------------------------------------------------ one picture

        private sealed class Encoded { public byte[] Data; public string Ext; public Size Pixels; }

        /// <summary>The picture, smaller; null (with the reason in <paramref name="info"/>) when it stays as it is.</summary>
        private static Encoded Recompress(byte[] data, Need need, SlimOptions o, SlimImage info)
        {
            if (data.Length < TooSmallToBother) { info.Note = "slim.note.small"; return null; }
            if (need != null && need.LeaveAlone) { info.Note = "slim.note.effect"; return null; }

            using (var ms = new MemoryStream(data, false))
            using (Image img = Image.FromStream(ms, false, true))
            {
                info.OldPixels = info.NewPixels = img.Size;
                Guid format = img.RawFormat.Guid;
                bool jpeg = format == ImageFormat.Jpeg.Guid, png = format == ImageFormat.Png.Guid;
                if (!jpeg && !png && format != ImageFormat.Bmp.Guid && format != ImageFormat.Tiff.Guid && format != ImageFormat.MemoryBmp.Guid) { info.Note = "slim.note.format"; return null; }
                foreach (Guid dimension in img.FrameDimensionsList)
                    if (img.GetFrameCount(new FrameDimension(dimension)) > 1) { info.Note = "slim.note.format"; return null; }
                // GDI+ turns CMYK into RGB by rule of thumb: the colours would shift
                if ((img.Flags & (int)(ImageFlags.ColorSpaceCmyk | ImageFlags.ColorSpaceYcck)) != 0) { info.Note = "slim.note.format"; return null; }

                // ---- how many pixels does it need?
                int w = img.Width, h = img.Height;
                double scale = 1;
                if (o.Ppi > 0 && need != null && !need.KeepPixels && need.Inches.Count > 0)
                {
                    // A photo tagged "shown turned by a quarter" (EXIF orientation, as phones write it) has
                    // its width along the height of its frame. Go by the frame's proportions rather than
                    // by the tag alone: they also tell when a program does not follow the tag.
                    bool tagged = Orientation(img) >= 5;
                    double most = 0;
                    foreach (SizeF inch in need.Inches)
                    {
                        double frame = inch.Width / inch.Height, upright = (double)w / h, turned = (double)h / w;
                        bool quarter = tagged && Math.Abs(Math.Log(frame / turned)) < Math.Abs(Math.Log(frame / upright));
                        double pw = quarter ? h : w, ph = quarter ? w : h;
                        most = Math.Max(most, Math.Max(inch.Width * o.Ppi / pw, inch.Height * o.Ppi / ph));
                    }
                    if (most < 0.9) scale = most;               // a few percent fewer pixels are not worth a resampling
                }
                int nw = Math.Max(1, (int)Math.Round(w * scale)), nh = Math.Max(1, (int)Math.Round(h * scale));

                // ---- draw it at that size
                bool alpha = Image.IsAlphaPixelFormat(img.PixelFormat) || ((img.PixelFormat & PixelFormat.Indexed) != 0 && PaletteHasAlpha(img));
                using (var bmp = new Bitmap(nw, nh, alpha ? PixelFormat.Format32bppArgb : PixelFormat.Format24bppRgb))
                {
                    using (Graphics g = Graphics.FromImage(bmp))
                    using (var attributes = new ImageAttributes())
                    {
                        attributes.SetWrapMode(WrapMode.TileFlipXY);       // no dark fringe along the edges
                        g.CompositingMode = CompositingMode.SourceCopy;
                        g.CompositingQuality = CompositingQuality.HighQuality;
                        g.InterpolationMode = scale < 1 ? InterpolationMode.HighQualityBicubic : InterpolationMode.NearestNeighbor;
                        g.PixelOffsetMode = PixelOffsetMode.Half;
                        g.DrawImage(img, new Rectangle(0, 0, nw, nh), 0, 0, w, h, GraphicsUnit.Pixel, attributes);
                    }
                    CopyTags(img, bmp, scale);
                    bool transparent = alpha && HasTransparency(bmp);

                    // ---- and encode it
                    byte[] result;
                    string ext;
                    if (jpeg) { result = Jpeg(bmp, o.Quality); ext = "jpeg"; }
                    else if (transparent) { result = Png(bmp); ext = "png"; }
                    else
                    {
                        // Opaque and stored without loss. A photo is many times smaller as JPEG; a drawing
                        // or a screenshot is not, and JPEG would only blur its edges: let the sizes decide.
                        byte[] asJpeg = o.PhotosToJpeg ? Jpeg(bmp, o.Quality) : null;
                        double asPngAbout = png ? data.Length * ((double)nw * nh / ((double)w * h)) : double.MaxValue;
                        if (asJpeg != null && png && asJpeg.Length < 0.5 * asPngAbout) { result = asJpeg; ext = "jpeg"; }
                        else
                        {
                            byte[] asPng;
                            if (alpha) using (Bitmap opaque = bmp.Clone(new Rectangle(0, 0, nw, nh), PixelFormat.Format24bppRgb)) { CopyTags(img, opaque, scale); asPng = Png(opaque); }
                            else asPng = Png(bmp);
                            if (asJpeg != null && asJpeg.Length < 0.5 * asPng.Length) { result = asJpeg; ext = "jpeg"; }
                            else { result = asPng; ext = "png"; }
                        }
                    }

                    if (result.Length > 0.9 * data.Length) { info.Note = "slim.note.nogain"; return null; }
                    return new Encoded { Data = result, Ext = ext, Pixels = new Size(nw, nh) };
                }
            }
        }

        private static bool SameFormat(string a, string b)
        {
            return IsJpegExt(a) && IsJpegExt(b);
        }

        private static bool IsJpegExt(string ext)
        {
            ext = ext.ToLowerInvariant();
            return ext == "jpg" || ext == "jpeg" || ext == "jpe" || ext == "jfif";
        }

        /// <summary>EXIF orientation, 1 to 8; 5 and up mean the picture is shown turned by a quarter.</summary>
        private static int Orientation(Image img)
        {
            try
            {
                if (Array.IndexOf(img.PropertyIdList, 0x0112) < 0) return 1;
                PropertyItem item = img.GetPropertyItem(0x0112);
                int v = item.Value != null && item.Value.Length >= 2 ? BitConverter.ToUInt16(item.Value, 0) : 1;
                return v >= 1 && v <= 8 ? v : 1;
            }
            catch { return 1; }
        }

        /// <summary>
        /// What PowerPoint reads from the file besides the pixels: the EXIF orientation, the colour
        /// profile and gamma, and the physical size ("Reset Picture" goes by it).
        /// </summary>
        private static void CopyTags(Image from, Image to, double scale)
        {
            foreach (int id in new[] { 0x0112, 0x8773, 0x0301, 0x0302, 0x0303 })
            {
                try { if (Array.IndexOf(from.PropertyIdList, id) >= 0) to.SetPropertyItem(from.GetPropertyItem(id)); }
                catch { }
            }
            try
            {
                float dx = (float)(from.HorizontalResolution * scale), dy = (float)(from.VerticalResolution * scale);
                if (dx >= 1 && dy >= 1 && to is Bitmap) ((Bitmap)to).SetResolution(dx, dy);
            }
            catch { }
        }

        private static bool PaletteHasAlpha(Image img)
        {
            try
            {
                ColorPalette p = img.Palette;
                if ((p.Flags & 1) != 0) return true;        // PaletteFlagsHasAlpha
                foreach (Color c in p.Entries) if (c.A < 255) return true;
            }
            catch { }
            return false;
        }

        private static bool HasTransparency(Bitmap bmp)
        {
            BitmapData bits = bmp.LockBits(new Rectangle(0, 0, bmp.Width, bmp.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            try
            {
                var row = new byte[bits.Width * 4];
                for (int y = 0; y < bits.Height; y++)
                {
                    System.Runtime.InteropServices.Marshal.Copy(new IntPtr(bits.Scan0.ToInt64() + (long)y * bits.Stride), row, 0, row.Length);
                    for (int x = 3; x < row.Length; x += 4) if (row[x] != 255) return true;
                }
                return false;
            }
            finally { bmp.UnlockBits(bits); }
        }

        private static ImageCodecInfo _jpegCodec;

        private static byte[] Jpeg(Image img, int quality)
        {
            if (_jpegCodec == null)
                foreach (ImageCodecInfo c in ImageCodecInfo.GetImageEncoders()) if (c.FormatID == ImageFormat.Jpeg.Guid) _jpegCodec = c;
            using (var ms = new MemoryStream())
            using (var parameters = new EncoderParameters(1))
            {
                parameters.Param[0] = new EncoderParameter(ImageEncoder.Quality, (long)Math.Max(1, Math.Min(100, quality)));
                img.Save(ms, _jpegCodec, parameters);
                return ms.ToArray();
            }
        }

        private static byte[] Png(Image img)
        {
            using (var ms = new MemoryStream())
            {
                img.Save(ms, ImageFormat.Png);
                return ms.ToArray();
            }
        }

        // ------------------------------------------------------------------ the package

        /// <summary>
        /// Copies the package entry by entry, in the same order, with the replaced pictures put in.
        /// A picture whose format changed has a new name: the relationships that point to it and
        /// the list of content types are brought in line.
        /// </summary>
        private static void WritePackage(ZipArchive zip, string output, Dictionary<string, Replacement> replaced)
        {
            var renamed = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (KeyValuePair<string, Replacement> r in replaced)
                if (r.Value.NewName != r.Key) renamed[r.Key] = r.Value.NewName;

            using (var outFs = new FileStream(output, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            using (var outZip = new ZipArchive(outFs, ZipArchiveMode.Create))
            {
                foreach (ZipArchiveEntry e in zip.Entries)
                {
                    Replacement r;
                    byte[] changed = null;
                    string name = e.FullName;
                    if (replaced.TryGetValue(name, out r))
                    {
                        name = r.NewName;
                        changed = File.ReadAllBytes(r.TempFile);
                    }
                    else if (renamed.Count > 0 && name == ContentTypes) changed = PatchContentTypes(e, renamed);
                    else if (renamed.Count > 0 && name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) changed = PatchRelationships(e, renamed);

                    // what is compressed already (pictures, video) only costs time to compress again
                    ZipArchiveEntry copy = outZip.CreateEntry(name, IsMedia(name) ? CompressionLevel.Fastest : CompressionLevel.Optimal);
                    try { copy.LastWriteTime = e.LastWriteTime; } catch { }
                    using (Stream to = copy.Open())
                    {
                        if (changed != null) to.Write(changed, 0, changed.Length);
                        else using (Stream from = e.Open()) from.CopyTo(to);
                    }
                }
            }
        }

        /// <summary>The relationships file with the renamed pictures' new names; null when it names none of them.</summary>
        private static byte[] PatchRelationships(ZipArchiveEntry entry, Dictionary<string, string> renamed)
        {
            XmlDocument doc = LineCaps.LoadXml(entry);
            string source = SourceOf(entry.FullName);
            bool any = false;
            foreach (XmlElement rel in Relationships(doc))
            {
                if (rel.GetAttribute("TargetMode") == "External") continue;
                string target = rel.GetAttribute("Target"), newName;
                if (!renamed.TryGetValue(Resolve(source, target), out newName)) continue;
                rel.SetAttribute("Target", target.Substring(0, target.LastIndexOf('/') + 1) + FileName(newName));
                any = true;
            }
            return any ? new UTF8Encoding(false).GetBytes(doc.OuterXml) : null;
        }

        private static byte[] PatchContentTypes(ZipArchiveEntry entry, Dictionary<string, string> renamed)
        {
            XmlDocument doc = LineCaps.LoadXml(entry);
            XmlElement root = doc.DocumentElement;
            var known = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            XmlNode lastDefault = null;
            foreach (XmlNode n in root.ChildNodes)
            {
                var e = n as XmlElement;
                if (e == null) continue;
                if (e.LocalName == "Default") { known.Add(e.GetAttribute("Extension")); lastDefault = e; }
                else if (e.LocalName == "Override")
                {
                    string newName;         // a part listed by name (PowerPoint lists pictures by extension, but it is allowed)
                    if (renamed.TryGetValue(e.GetAttribute("PartName").TrimStart('/'), out newName))
                    {
                        e.SetAttribute("PartName", "/" + newName);
                        e.SetAttribute("ContentType", MimeOf(Ext(newName)));
                    }
                }
            }
            foreach (string newName in renamed.Values)
            {
                string ext = Ext(newName);
                if (!known.Add(ext)) continue;
                XmlElement d = doc.CreateElement("Default", NsTypes);
                d.SetAttribute("Extension", ext);
                d.SetAttribute("ContentType", MimeOf(ext));
                if (lastDefault != null) root.InsertAfter(d, lastDefault);
                else root.PrependChild(d);
                lastDefault = d;
            }
            return new UTF8Encoding(false).GetBytes(doc.OuterXml);
        }

        private static string MimeOf(string ext)
        {
            return IsJpegExt(ext) ? "image/jpeg" : "image/png";
        }

        private static IEnumerable<XmlElement> Relationships(XmlDocument rels)
        {
            var nm = new XmlNamespaceManager(rels.NameTable);
            nm.AddNamespace("r", NsRel);
            foreach (XmlNode n in rels.SelectNodes("//r:Relationship", nm)) yield return (XmlElement)n;
        }

        // ------------------------------------------------------------------ part names

        private static bool IsMedia(string partName)
        {
            return partName.StartsWith("ppt/media/", StringComparison.OrdinalIgnoreCase);
        }

        private static string Ext(string name)
        {
            int dot = name.LastIndexOf('.');
            return dot < 0 || dot < name.LastIndexOf('/') ? "" : name.Substring(dot + 1);
        }

        private static string FileName(string partName)
        {
            return partName.Substring(partName.LastIndexOf('/') + 1);
        }

        /// <summary>"ppt/slides/_rels/slide1.xml.rels" is about "ppt/slides/slide1.xml"; "_rels/.rels" about the package ("").</summary>
        internal static string SourceOf(string relsName)
        {
            int i = relsName.LastIndexOf("_rels/", StringComparison.Ordinal);
            if (i < 0) return "";
            string file = relsName.Substring(i + 6);
            if (file.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) file = file.Substring(0, file.Length - 5);
            return relsName.Substring(0, i) + file;
        }

        /// <summary>The part a relationship target names, seen from the part that has the relationship.</summary>
        internal static string Resolve(string sourcePart, string target)
        {
            if (target.StartsWith("/", StringComparison.Ordinal)) return target.TrimStart('/');
            var parts = new List<string>();
            int slash = sourcePart.LastIndexOf('/');
            if (slash > 0) parts.AddRange(sourcePart.Substring(0, slash).Split('/'));
            foreach (string segment in target.Split('/'))
            {
                if (segment.Length == 0 || segment == ".") continue;
                if (segment == "..") { if (parts.Count > 0) parts.RemoveAt(parts.Count - 1); }
                else parts.Add(segment);
            }
            return string.Join("/", parts.ToArray());
        }

        private static byte[] ReadAll(ZipArchiveEntry entry)
        {
            using (Stream s = entry.Open())
            using (var ms = new MemoryStream((int)Math.Min(int.MaxValue, Math.Max(0, entry.Length))))
            {
                s.CopyTo(ms);
                return ms.ToArray();
            }
        }
    }
}
