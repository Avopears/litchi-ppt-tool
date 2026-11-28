// Builds a presentation full of oversized pictures in a real PowerPoint instance, slims a copy of it and checks
// the copy: every picture has the pixels its size on the slide calls for, PowerPoint opens the file without
// repairing it, and the slides still look the same.
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
    internal static class SlimTests
    {
        private static readonly Office.MsoTriState T = Office.MsoTriState.msoTrue, F = Office.MsoTriState.msoFalse;
        private const string NsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static int passed, failed;
        private static string dir;

        /// <summary>One picture as an object on a slide shows it: the file behind it.</summary>
        private sealed class Shown { public string Part; public Size Pixels; public Guid Format; public long Bytes; public int Orientation; public bool CornerTransparent; }

        public static void Run(PowerPoint.Application app)
        {
            dir = Path.Combine(Path.GetTempPath(), "PptSelectionExport", "slimtests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            PowerPoint.Presentation pres = app.Presentations.Add(F);
            PowerPoint.Presentation slimmed = null;
            try
            {
                Check("part names", Slimmer.SourceOf("ppt/slides/_rels/slide1.xml.rels") == "ppt/slides/slide1.xml" && Slimmer.SourceOf("_rels/.rels") == ""
                      && Slimmer.Resolve("ppt/slides/slide1.xml", "../media/image1.png") == "ppt/media/image1.png"
                      && Slimmer.Resolve("ppt/slides/slide1.xml", "/ppt/media/image1.png") == "ppt/media/image1.png"
                      && Slimmer.Resolve("", "ppt/presentation.xml") == "ppt/presentation.xml", "");

                // ---- source pictures, all far larger than they are shown
                string jpgA = SaveJpeg(Photo(2400, 1600, 1), "a.jpg"), jpgC = SaveJpeg(Photo(2400, 1600, 3), "c.jpg"), jpgD = SaveJpeg(Photo(2400, 1600, 4), "d.jpg");
                string jpgE = SaveJpeg(Photo(2400, 1600, 5), "e.jpg"), jpgF = SaveJpeg(Photo(2400, 1600, 6), "f.jpg"), jpgBg = SaveJpeg(Photo(4000, 2250, 7), "bg.jpg");
                string pngB = SavePng(Photo(1600, 1200, 2), "b.png"), pngTexture = SavePng(Photo(512, 512, 8), "texture.png"), pngTiny = SavePng(Photo(300, 200, 9), "tiny.png");
                string pngAlpha = SavePng(Cutout(1500, 1500, 10), "alpha.png"), pngGraphic = SavePng(Drawing(1600, 1000), "graphic.png"), pngKeyed = SavePng(Photo(1600, 1200, 13), "keyed.png");
                string jpgTurned = SaveJpeg(Photo(1800, 1200, 11), "turned.jpg", 6), jpgTurned2 = SaveJpeg(Photo(1800, 1200, 12), "turned2.jpg", 6);

                PowerPoint.Slide s1 = pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                Add(s1, jpgA, "photoJ", 40, 40, 216, 144);                  // 3 in wide
                Add(s1, pngB, "photoP", 300, 40, 288, 216);                 // 4 in
                Add(s1, pngAlpha, "alpha", 620, 40, 216, 216);              // 3 in
                Add(s1, pngGraphic, "graphic", 40, 300, 360, 225);          // 5 in
                PowerPoint.Shape keyed = Add(s1, pngKeyed, "keyed", 440, 300, 216, 162);     // "Set Transparent Color": one exact colour is see-through
                keyed.PictureFormat.TransparentBackground = T;
                keyed.PictureFormat.TransparencyColor = 0xFFFFFF;

                PowerPoint.Slide s2 = pres.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutBlank);
                Add(s2, jpgC, "twiceSmall", 40, 40, 144, 96);               // the same file twice: 2 in and 6 in
                Add(s2, jpgC, "twiceBig", 200, 40, 432, 288);
                PowerPoint.Shape cropped = Add(s2, jpgD, "cropped", 40, 340, 432, 288);     // 6 in, then half of it cut away
                float uncropped = cropped.Width;
                cropped.PictureFormat.Crop.ShapeWidth = uncropped / 2;
                bool reallyCropped = cropped.Width < uncropped * 0.6f;

                PowerPoint.Slide s3 = pres.Slides.Add(3, PowerPoint.PpSlideLayout.ppLayoutBlank);
                Add(s3, jpgE, "inGroup", 40, 40, 144, 96);
                s3.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 200, 40, 60, 60).Name = "besideIt";
                PowerPoint.Shape grp = s3.Shapes.Range(new object[] { "inGroup", "besideIt" }).Group();
                grp.LockAspectRatio = T;
                grp.Width = grp.Width * 2;                                   // the group now shows the picture at twice the size
                float inGroupWidth = grp.GroupItems[1].Width;
                PowerPoint.Shape filled = s3.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 500, 40, 360, 216); filled.Name = "filled";     // 5 in x 3 in
                filled.Fill.UserPicture(jpgF);
                PowerPoint.Shape tiled = s3.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 40, 320, 300, 200); tiled.Name = "tiled";
                tiled.Fill.UserTextured(pngTexture);

                PowerPoint.Slide s4 = pres.Slides.Add(4, PowerPoint.PpSlideLayout.ppLayoutBlank);
                s4.FollowMasterBackground = F;
                s4.Background.Fill.UserPicture(jpgBg);
                Add(s4, pngTiny, "tiny", 520, 40, 400, 267);                 // shown larger than its pixels: nothing to take away
                // photos tagged "shown turned by a quarter" (EXIF orientation 6): one in the frame PowerPoint gives it,
                // one in a frame with the turned proportions, as a program that follows the tag would have it
                PowerPoint.Shape turned = s4.Shapes.AddPicture(jpgTurned, F, T, 40, 40); turned.Name = "turned";
                turned.LockAspectRatio = T;
                turned.Height = 150;
                bool shownUpright = turned.Width < turned.Height;
                float turnedWidth = turned.Width;
                PowerPoint.Shape turnedFrame = s4.Shapes.AddPicture(jpgTurned2, F, T, 300, 40); turnedFrame.Name = "turnedFrame";
                turnedFrame.LockAspectRatio = F;
                turnedFrame.Width = 100; turnedFrame.Height = 150;

                string input = Path.Combine(dir, "input.pptx"), output = Path.Combine(dir, "output.pptx");
                pres.SaveCopyAs(input, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, F);
                Dictionary<string, Shown> before = Pictures(input);
                SlimOverview overview = Slimmer.Inspect(input);
                Check("overview of the file", overview.ImageCount == 14 && overview.ImageBytes > overview.FileBytes * 0.9 && overview.MediaCount == 0,
                      overview.ImageCount + " pictures, " + Mb(overview.ImageBytes) + " of " + Mb(overview.FileBytes));

                // ---- slim it: 220 ppi, quality 80
                int calls = 0, lastDone = -1, lastTotal = -1;
                var watch = System.Diagnostics.Stopwatch.StartNew();
                SlimReport report = Slimmer.Slim(input, output, new SlimOptions { Ppi = 220, Quality = 80, PhotosToJpeg = true },
                                                 delegate(int done, int total) { calls++; lastDone = done; lastTotal = total; }, null);
                watch.Stop();
                Dictionary<string, Shown> after = Pictures(output);
                Check("the file is much smaller", report.NewFileBytes < report.OldFileBytes * 0.5 && report.NewFileBytes == new FileInfo(output).Length,
                      Mb(report.OldFileBytes) + " -> " + Mb(report.NewFileBytes) + ", " + report.ChangedCount + " of " + report.Images.Count + " pictures changed, " + watch.ElapsedMilliseconds + " ms");
                Check("progress is reported to the end", calls > 1 && lastDone == lastTotal && lastTotal > 0, calls + " calls, last " + lastDone + "/" + lastTotal);
                Check("no work folder left behind", !Directory.Exists(output + ".work"), "");

                // ---- each picture has the pixels its size on the slide calls for
                Check("picture 3 in wide: 660 px", Wide(after, "photoJ", 660) && after["photoJ"].Format == ImageFormat.Jpeg.Guid, Describe(before, after, "photoJ"));
                Check("photo stored as PNG becomes JPEG, 880 px", Wide(after, "photoP", 880) && after["photoP"].Format == ImageFormat.Jpeg.Guid && after["photoP"].Part.EndsWith(".jpeg"), Describe(before, after, "photoP"));
                Check("transparent PNG stays PNG, 660 px", Wide(after, "alpha", 660) && after["alpha"].Format == ImageFormat.Png.Guid && after["alpha"].CornerTransparent, Describe(before, after, "alpha"));
                Check("a drawing stays PNG", after["graphic"].Format == ImageFormat.Png.Guid, Describe(before, after, "graphic"));
                Check("picture with a transparent colour is left alone", after["keyed"].Part == before["keyed"].Part && after["keyed"].Bytes == before["keyed"].Bytes && after["keyed"].Pixels == new Size(1600, 1200), Describe(before, after, "keyed"));
                Check("used twice: sized for the larger use, 1320 px", Wide(after, "twiceBig", 1320) && after["twiceSmall"].Part == after["twiceBig"].Part && before["twiceSmall"].Part == before["twiceBig"].Part, Describe(before, after, "twiceBig"));
                Check("cropped: the hidden half keeps its share, 1320 px", reallyCropped && Wide(after, "cropped", 1320), "cropped from " + uncropped + " to " + cropped.Width + " pt; " + Describe(before, after, "cropped"));
                Check("inside a scaled group", Wide(after, "inGroup", (int)Math.Round(inGroupWidth / 72.0 * 220)), "shown " + inGroupWidth + " pt wide; " + Describe(before, after, "inGroup"));
                Check("picture fill of a shape, 1100 px", Wide(after, "filled", 1100), Describe(before, after, "filled"));
                Check("tiled texture keeps its pixels", after["tiled"].Pixels == before["tiled"].Pixels, Describe(before, after, "tiled"));
                Check("slide background: the slide's width, 2933 px", Wide(after, "background4", 2933), Describe(before, after, "background4"));
                Check("shown larger than its pixels: none taken away", after["tiny"].Pixels == new Size(300, 200), Describe(before, after, "tiny"));
                int frameWide = shownUpright ? (int)Math.Round(turnedWidth / 72.0 * 220 * 1800 / 1200) : (int)Math.Round(turnedWidth / 72.0 * 220);
                Check("EXIF orientation tag is kept", before["turned"].Orientation == 6 && after["turned"].Orientation == 6 && Wide(after, "turned", frameWide),
                      "PowerPoint turns the frame: " + shownUpright + ", " + turnedWidth + " pt wide; " + Describe(before, after, "turned"));
                // 100 pt x 150 pt for a file of 1800 x 1200: its 1200 side lies along the 100 pt, so 100 / 72 * 220 / 1200 of its pixels
                Check("tagged photo in a turned frame: sized along it", after["turnedFrame"].Orientation == 6 && Wide(after, "turnedFrame", (int)Math.Round(1800 * (100 / 72.0 * 220 / 1200))),
                      Describe(before, after, "turnedFrame"));

                // ---- PowerPoint opens it as it is, and it looks the same
                slimmed = app.Presentations.Open2007(output, T, T, F, F);       // last argument: do not repair
                bool sameShapes = slimmed.Slides.Count == pres.Slides.Count;
                for (int i = 1; sameShapes && i <= pres.Slides.Count; i++) sameShapes = slimmed.Slides[i].Shapes.Count == pres.Slides[i].Shapes.Count;
                Check("PowerPoint opens the slimmed file unrepaired", sameShapes, slimmed.Slides.Count + " slides");
                string look = "";
                bool sameLook = true;
                for (int i = 1; i <= pres.Slides.Count; i++)
                {
                    string a = Path.Combine(dir, "before" + i + ".png"), b = Path.Combine(dir, "after" + i + ".png");
                    pres.Slides[i].Export(a, "PNG", 1280, 720);
                    slimmed.Slides[i].Export(b, "PNG", 1280, 720);
                    double mean, worstBlock;
                    Compare(a, b, out mean, out worstBlock);
                    look += string.Format("slide {0}: {1:0.00} / {2:0.0}   ", i, mean, worstBlock);
                    if (mean > 2.5 || worstBlock > 16) sameLook = false;       // of 255; a wrong, turned or missing picture gives a block of 50 and more
                }
                Check("the slides look the same (mean / worst block)", sameLook, look);
                slimmed.Saved = T; slimmed.Close(); slimmed = null;

                // ---- the settings
                string output2 = Path.Combine(dir, "output2.pptx");
                Slimmer.Slim(input, output2, new SlimOptions { Ppi = 0, Quality = 80, PhotosToJpeg = true }, null, null);
                Dictionary<string, Shown> keep = Pictures(output2);
                Check("ppi 0: every picture keeps its pixels", keep["photoJ"].Pixels == new Size(2400, 1600) && keep["photoP"].Pixels == new Size(1600, 1200) && keep["background4"].Pixels == new Size(4000, 2250), "");
                Slimmer.Slim(input, output2, new SlimOptions { Ppi = 220, Quality = 80, PhotosToJpeg = false }, null, null);
                Dictionary<string, Shown> noJpeg = Pictures(output2);
                Check("photos may stay PNG", noJpeg["photoP"].Format == ImageFormat.Png.Guid && Wide(noJpeg, "photoP", 880) && noJpeg["photoP"].Part.EndsWith(".png"), noJpeg["photoP"].Part);
                long low = Slimmer.Slim(input, output2, new SlimOptions { Ppi = 220, Quality = 40, PhotosToJpeg = true }, null, null).NewFileBytes;
                long high = Slimmer.Slim(input, output2, new SlimOptions { Ppi = 220, Quality = 95, PhotosToJpeg = true }, null, null).NewFileBytes;
                Check("lower quality, smaller file", low < report.NewFileBytes && report.NewFileBytes < high, Mb(low) + " < " + Mb(report.NewFileBytes) + " < " + Mb(high) + " at 40 / 80 / 95");

                // ---- stopping half way leaves nothing behind
                string output3 = Path.Combine(dir, "output3.pptx");
                bool stopped = false;
                int asked = 0;
                try { Slimmer.Slim(input, output3, new SlimOptions(), null, delegate { return ++asked > 1; }); }
                catch (OperationCanceledException) { stopped = true; }
                Check("cancelled: no file, no work folder", stopped && !File.Exists(output3) && !Directory.Exists(output3 + ".work"), "");

                Console.WriteLine("slim tests: {0} passed, {1} failed", passed, failed);
            }
            finally
            {
                if (slimmed != null) { try { slimmed.Saved = T; slimmed.Close(); } catch { } }
                pres.Saved = T;
                pres.Close();
                try { Directory.Delete(dir, true); } catch { }
            }
        }

        private static PowerPoint.Shape Add(PowerPoint.Slide slide, string file, string name, float left, float top, float width, float height)
        {
            PowerPoint.Shape s = slide.Shapes.AddPicture(file, F, T, left, top, width, height);
            s.Name = name;
            return s;
        }

        private static bool Wide(Dictionary<string, Shown> shown, string name, int pixels)
        {
            return shown.ContainsKey(name) && Math.Abs(shown[name].Pixels.Width - pixels) <= 2;
        }

        private static string Describe(Dictionary<string, Shown> before, Dictionary<string, Shown> after, string name)
        {
            if (!before.ContainsKey(name) || !after.ContainsKey(name)) return name + " not found";
            Shown a = before[name], b = after[name];
            return string.Format("{0} {1}x{2} {3} KB -> {4} {5}x{6} {7} KB", Path.GetFileName(a.Part), a.Pixels.Width, a.Pixels.Height, a.Bytes / 1024,
                                 Path.GetFileName(b.Part), b.Pixels.Width, b.Pixels.Height, b.Bytes / 1024);
        }

        private static string Mb(long bytes) { return (bytes / 1048576.0).ToString("0.00") + " MB"; }

        /// <summary>
        /// The picture file behind every picture, picture-filled shape and background of a presentation
        /// file, by the object's name ("background" + slide number for a background).
        /// </summary>
        private static Dictionary<string, Shown> Pictures(string file)
        {
            var shown = new Dictionary<string, Shown>();
            using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                foreach (ZipArchiveEntry entry in zip.Entries)
                {
                    if (!entry.FullName.StartsWith("ppt/slides/slide") || !entry.FullName.EndsWith(".xml")) continue;
                    var targets = new Dictionary<string, string>();
                    XmlDocument rels = LineCaps.LoadXml(zip.GetEntry("ppt/slides/_rels/" + Path.GetFileName(entry.FullName) + ".rels"));
                    foreach (XmlNode n in rels.DocumentElement.ChildNodes)
                    {
                        var rel = n as XmlElement;
                        if (rel != null) targets[rel.GetAttribute("Id")] = Slimmer.Resolve(entry.FullName, rel.GetAttribute("Target"));
                    }
                    XmlDocument doc = LineCaps.LoadXml(entry);
                    XmlNamespaceManager nm = LineCaps.Namespaces(doc);
                    foreach (XmlNode n in doc.SelectNodes("//a:blip", nm))
                    {
                        var blip = (XmlElement)n;
                        string part;
                        if (!targets.TryGetValue(blip.GetAttribute("embed", NsR), out part)) continue;
                        string name = null;
                        for (XmlNode up = blip.ParentNode; up is XmlElement && name == null; up = up.ParentNode)
                        {
                            if (up.LocalName == "bg") name = "background" + Path.GetFileNameWithoutExtension(entry.FullName).Substring(5);
                            else if (up.LocalName == "pic" || up.LocalName == "sp") name = ((XmlElement)up.SelectSingleNode("*/p:cNvPr", nm)).GetAttribute("name");
                        }
                        if (name == null) continue;

                        ZipArchiveEntry media = zip.GetEntry(part);
                        var info = new Shown { Part = part, Bytes = media.Length };
                        using (var ms = new MemoryStream())
                        {
                            using (Stream s = media.Open()) s.CopyTo(ms);
                            ms.Position = 0;
                            using (var img = (Bitmap)Image.FromStream(ms))
                            {
                                info.Pixels = img.Size;
                                info.Format = img.RawFormat.Guid;
                                info.Orientation = Array.IndexOf(img.PropertyIdList, 0x0112) >= 0 ? BitConverter.ToUInt16(img.GetPropertyItem(0x0112).Value, 0) : 1;
                                info.CornerTransparent = img.GetPixel(0, 0).A == 0;
                            }
                        }
                        shown[name] = info;
                    }
                }
            }
            return shown;
        }

        /// <summary>Mean difference per colour channel over the whole picture, and over the worst 32 x 32 block.</summary>
        private static void Compare(string fileA, string fileB, out double mean, out double worstBlock)
        {
            using (var a = new Bitmap(fileA))
            using (var b = new Bitmap(fileB))
            {
                mean = worstBlock = 255;
                if (a.Size != b.Size) return;
                int w = a.Width, h = a.Height;
                const int block = 32;
                var sums = new long[(h / block + 1) * (w / block + 1)];
                var counts = new int[sums.Length];
                BitmapData da = a.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                BitmapData db = b.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                byte[] ra = new byte[w * 3], rb = new byte[w * 3];
                long total = 0;
                for (int y = 0; y < h; y++)
                {
                    System.Runtime.InteropServices.Marshal.Copy(new IntPtr(da.Scan0.ToInt64() + (long)y * da.Stride), ra, 0, ra.Length);
                    System.Runtime.InteropServices.Marshal.Copy(new IntPtr(db.Scan0.ToInt64() + (long)y * db.Stride), rb, 0, rb.Length);
                    for (int x = 0; x < w; x++)
                    {
                        int d = Math.Abs(ra[x * 3] - rb[x * 3]) + Math.Abs(ra[x * 3 + 1] - rb[x * 3 + 1]) + Math.Abs(ra[x * 3 + 2] - rb[x * 3 + 2]);
                        total += d;
                        int cell = (y / block) * (w / block + 1) + x / block;
                        sums[cell] += d;
                        counts[cell] += 3;
                    }
                }
                a.UnlockBits(da); b.UnlockBits(db);
                mean = total / (3.0 * w * h);
                worstBlock = 0;
                for (int i = 0; i < sums.Length; i++) if (counts[i] > 0) worstBlock = Math.Max(worstBlock, (double)sums[i] / counts[i]);
            }
        }

        // ------------------------------------------------------------------ test pictures

        /// <summary>Smooth colour fields with fine grain and a few hard edges: it compresses like a photograph.</summary>
        private static Bitmap Photo(int w, int h, int seed)
        {
            var bmp = new Bitmap(w, h, PixelFormat.Format24bppRgb);
            var random = new Random(seed);
            BitmapData bits = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);
            var row = new byte[bits.Stride];
            for (int y = 0; y < h; y++)
            {
                double fy = (double)y / h;
                for (int x = 0; x < w; x++)
                {
                    double fx = (double)x / w;
                    row[x * 3] = Clamp(127 + 100 * Math.Cos(fx * 4 - fy * 8 + seed * 3) + random.Next(-6, 7));
                    row[x * 3 + 1] = Clamp(127 + 100 * Math.Sin(fx * 5 + fy * 6 + seed * 2) + random.Next(-6, 7));
                    row[x * 3 + 2] = Clamp(127 + 100 * Math.Sin(fx * 9 + seed) * Math.Cos(fy * 7) + random.Next(-6, 7));
                }
                System.Runtime.InteropServices.Marshal.Copy(row, 0, new IntPtr(bits.Scan0.ToInt64() + (long)y * bits.Stride), row.Length);
            }
            bmp.UnlockBits(bits);
            using (Graphics g = Graphics.FromImage(bmp))
            using (var font = new Font("Arial", h / 4f, FontStyle.Bold, GraphicsUnit.Pixel))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.FillEllipse(Brushes.White, w * 0.08f, h * 0.1f, w * 0.2f, h * 0.3f);
                g.DrawString(seed.ToString(), font, Brushes.Black, w * 0.45f, h * 0.3f);     // which picture is it, and which way up
                using (var pen = new Pen(Color.Yellow, Math.Max(2, h / 100))) g.DrawLine(pen, 0, h * 0.9f, w, h * 0.8f);
            }
            return bmp;
        }

        /// <summary>A photo-like disc on a fully transparent ground.</summary>
        private static Bitmap Cutout(int w, int h, int seed)
        {
            var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
            using (Bitmap photo = Photo(w, h, seed))
            using (Graphics g = Graphics.FromImage(bmp))
            using (var brush = new TextureBrush(photo))
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.FillEllipse(brush, w * 0.1f, h * 0.1f, w * 0.8f, h * 0.8f);
            }
            return bmp;
        }

        /// <summary>Flat colours, lines and text: a chart or a screenshot.</summary>
        private static Bitmap Drawing(int w, int h)
        {
            var bmp = new Bitmap(w, h, PixelFormat.Format24bppRgb);
            using (Graphics g = Graphics.FromImage(bmp))
            using (var font = new Font("Arial", h / 14f, FontStyle.Regular, GraphicsUnit.Pixel))
            {
                g.Clear(Color.White);
                g.FillRectangle(Brushes.SteelBlue, w * 0.1f, h * 0.5f, w * 0.15f, h * 0.4f);
                g.FillRectangle(Brushes.IndianRed, w * 0.3f, h * 0.3f, w * 0.15f, h * 0.6f);
                g.FillRectangle(Brushes.SeaGreen, w * 0.5f, h * 0.6f, w * 0.15f, h * 0.3f);
                using (var pen = new Pen(Color.Black, 3)) { g.DrawLine(pen, w * 0.05f, h * 0.9f, w * 0.95f, h * 0.9f); g.DrawLine(pen, w * 0.05f, h * 0.9f, w * 0.05f, h * 0.1f); }
                for (int i = 0; i < 6; i++) g.DrawString("Series " + i + ": the quick brown fox", font, Brushes.Black, w * 0.4f, h * (0.02f + 0.045f * i));
            }
            return bmp;
        }

        private static byte Clamp(double v) { return (byte)Math.Max(0, Math.Min(255, (int)Math.Round(v))); }

        private static string SavePng(Bitmap bmp, string name)
        {
            string file = Path.Combine(dir, name);
            using (bmp) bmp.Save(file, ImageFormat.Png);
            return file;
        }

        /// <param name="orientation">EXIF orientation to tag the file with (6 = shown turned a quarter clockwise); 0 = none.</param>
        private static string SaveJpeg(Bitmap bmp, string name, int orientation = 0)
        {
            string file = Path.Combine(dir, name);
            using (bmp)
            using (var parameters = new EncoderParameters(1))
            {
                if (orientation > 0)
                {
                    var item = (PropertyItem)System.Runtime.Serialization.FormatterServices.GetUninitializedObject(typeof(PropertyItem));
                    item.Id = 0x0112; item.Type = 3; item.Len = 2; item.Value = new byte[] { (byte)orientation, 0 };
                    bmp.SetPropertyItem(item);
                }
                parameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 95L);
                ImageCodecInfo jpeg = null;
                foreach (ImageCodecInfo c in ImageCodecInfo.GetImageEncoders()) if (c.FormatID == ImageFormat.Jpeg.Guid) jpeg = c;
                bmp.Save(file, jpeg, parameters);
            }
            return file;
        }

        private static void Check(string name, bool ok, string detail)
        {
            if (ok) passed++; else failed++;
            Console.WriteLine("{0,-48} {1,-8} {2}", name, ok ? "OK" : "MISMATCH", detail);
        }
    }
}
