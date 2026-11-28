using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptSelectionExport
{
    /// <summary>The cap types of a line's ends, in the order of PowerPoint's own "Cap type" list.</summary>
    internal enum CapKind { Square, Round, Flat }

    /// <summary>Something to tell the user that is not a failure ("this only works on a slide").</summary>
    internal sealed class StyleNotice : Exception
    {
        public StyleNotice(string message) : base(message) { }
    }

    /// <summary>
    /// The cap type of the selected shapes' lines (Format Shape > Line > Cap type).
    ///
    /// PowerPoint's object model has no property for it: the value only exists in the file, as
    /// &lt;a:ln cap="sq|rnd|flat"&gt;. So this saves a temporary copy of the presentation, writes the
    /// cap into the copy's XML, opens the copy without a window and carries the format over with
    /// PickUp / Apply. The shape in the copy is the same shape, so every other property it hands
    /// over is the one the shape already has; its text is taken out of the copy first, because
    /// Apply would otherwise give all of the shape's text the format of its first run.
    /// </summary>
    internal static class LineCaps
    {
        public static readonly CapKind[] Caps = { CapKind.Square, CapKind.Round, CapKind.Flat };

        internal const string NsA = "http://schemas.openxmlformats.org/drawingml/2006/main";
        internal const string NsP = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private const string NsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string NsRel = "http://schemas.openxmlformats.org/package/2006/relationships";

        /// <summary>The shapes that can have a line: ordinary shapes, lines and connectors, pictures.</summary>
        internal const string ShapeXPath = "//p:sp | //p:cxnSp | //p:pic";

        private const Office.MsoTriState True = Office.MsoTriState.msoTrue, False = Office.MsoTriState.msoFalse;

        /// <summary>What comes before a:ln inside a shape's properties (the schema fixes the order).</summary>
        private static readonly string[] BeforeLine =
        {
            "xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"
        };

        internal static string ToXml(CapKind cap)
        {
            switch (cap)
            {
                case CapKind.Square: return "sq";
                case CapKind.Round: return "rnd";
                default: return "flat";
            }
        }

        /// <summary>The slide the shapes are on; null when they are on a master, a layout or a notes page.</summary>
        public static PowerPoint.Slide SlideOf(PowerPoint.ShapeRange range)
        {
            try
            {
                PowerPoint.Slide slide = range.Parent as PowerPoint.Slide;      // null for a master or a layout
                if (slide == null) return null;
                // A notes page answers as a slide too, with its slide's id and index, and its shapes
                // have ids of their own that can repeat the slide's. What tells them apart: a slide
                // belongs to the presentation, a notes page to its slide.
                PowerPoint.Presentation pres = slide.Parent as PowerPoint.Presentation;
                if (pres == null) return null;
                PowerPoint.Slide found = pres.Slides.FindBySlideID(slide.SlideID);
                if (found == null || found.SlideIndex != slide.SlideIndex) return null;
                // to be sure: the slide holds the shape
                PowerPoint.Shape first = range[1];
                var wanted = new HashSet<int>();
                wanted.Add(first.Id);
                PowerPoint.Shape same;
                if (!FindShapes(found, wanted).TryGetValue(first.Id, out same) || same.Name != first.Name) return null;
                return found;
            }
            catch { return null; }
        }

        /// <summary>
        /// Gives the lines of the shapes this cap type and returns how many shapes have it afterwards.
        /// Objects without a line of their own (tables, charts, SmartArt) are skipped.
        /// </summary>
        public static int SetCap(PowerPoint.Application app, PowerPoint.ShapeRange range, CapKind cap)
        {
            PowerPoint.Slide slide = SlideOf(range);
            if (slide == null) throw new StyleNotice(Strings.T("msg.cap.needsslide"));
            PowerPoint.Presentation pres = (PowerPoint.Presentation)slide.Parent;
            int slideId = slide.SlideID;

            var targets = new Dictionary<int, PowerPoint.Shape>();
            foreach (PowerPoint.Shape s in StyleTools.Leaves(range))
            {
                try { targets[s.Id] = s; } catch { }
            }
            if (targets.Count == 0) return 0;

            string dir = Exporter.NewTempDir();
            PowerPoint.Presentation copy = null;
            try
            {
                string file = Path.Combine(dir, "cap.pptx");
                pres.SaveCopyAs(file, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, False);

                HashSet<int> pending;       // the shapes whose cap really has to change
                int have;
                try { have = Patch(file, slideId, targets.Keys, ToXml(cap), out pending); }
                catch (InvalidDataException) { throw new InvalidOperationException(Strings.T("err.cap.copy")); }
                if (pending.Count == 0) return have;

                // read-only, untitled (keeps it out of "recent files"), without a window
                copy = app.Presentations.Open(file, True, True, False);
                Dictionary<int, PowerPoint.Shape> donors = FindShapes(copy.Slides.FindBySlideID(slideId), pending);

                try { app.StartNewUndoEntry(); } catch { }
                foreach (int id in pending)
                {
                    try
                    {
                        PowerPoint.Shape donor, target = targets[id];
                        if (!donors.TryGetValue(id, out donor) || donor.Name != target.Name) continue;
                        donor.PickUp();
                        target.Apply();
                        have++;
                    }
                    catch { }       // a locked shape, for example
                }
                return have;
            }
            finally
            {
                if (copy != null)
                {
                    try { copy.Saved = True; } catch { }
                    try { copy.Close(); } catch { }
                }
                Exporter.TryDeleteDir(dir);
            }
        }

        /// <summary>
        /// Writes the cap into the copy's slide. Returns how many of the shapes have it already;
        /// <paramref name="pending"/> are the ones that were changed in the copy and still have to be
        /// carried over to the open presentation.
        /// </summary>
        private static int Patch(string file, int slideId, ICollection<int> ids, string cap, out HashSet<int> pending)
        {
            pending = new HashSet<int>();
            var unchanged = new HashSet<int>();
            string entryName;
            XmlDocument doc;
            using (var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                ZipArchiveEntry entry = SlideEntry(zip, slideId);
                if (entry == null) return 0;
                entryName = entry.FullName;
                doc = LoadXml(entry);
            }

            XmlNamespaceManager nm = Namespaces(doc);
            foreach (XmlNode node in doc.SelectNodes(ShapeXPath, nm))
            {
                var shape = (XmlElement)node;
                int id;
                if (!TryGetId(shape, nm, out id) || !ids.Contains(id)) continue;
                var spPr = shape.SelectSingleNode("p:spPr", nm) as XmlElement;
                if (spPr == null) continue;

                var ln = spPr.SelectSingleNode("a:ln", nm) as XmlElement;
                if (ln != null && ln.GetAttribute("cap") == cap)
                {
                    if (!pending.Contains(id)) unchanged.Add(id);
                    continue;
                }
                if (ln == null)
                {
                    ln = doc.CreateElement("a", "ln", NsA);
                    XmlNode after = null;
                    foreach (XmlNode c in spPr.ChildNodes)
                        if (c.NodeType == XmlNodeType.Element && c.NamespaceURI == NsA && Array.IndexOf(BeforeLine, c.LocalName) >= 0) after = c;
                    if (after != null) spPr.InsertAfter(ln, after);
                    else spPr.PrependChild(ln);
                }
                ln.SetAttribute("cap", cap);

                // without its text the shape hands over its fill, line and effects only
                XmlNode text = shape.SelectSingleNode("p:txBody", nm);
                if (text != null) shape.RemoveChild(text);

                unchanged.Remove(id);
                pending.Add(id);
            }
            if (pending.Count == 0) return unchanged.Count;

            using (var fs = new FileStream(file, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Update))
            using (Stream s = zip.GetEntry(entryName).Open())
            {
                byte[] bytes = new UTF8Encoding(false).GetBytes(doc.OuterXml);
                s.SetLength(0);
                s.Write(bytes, 0, bytes.Length);
            }
            return unchanged.Count;
        }

        /// <summary>The part that holds the slide with this id: presentation.xml names it through a relationship.</summary>
        internal static ZipArchiveEntry SlideEntry(ZipArchive zip, int slideId)
        {
            ZipArchiveEntry presEntry = zip.GetEntry("ppt/presentation.xml");
            ZipArchiveEntry relsEntry = zip.GetEntry("ppt/_rels/presentation.xml.rels");
            if (presEntry == null || relsEntry == null) return null;

            XmlDocument pres = LoadXml(presEntry);
            var sldId = pres.SelectSingleNode("//p:sldIdLst/p:sldId[@id='" + slideId + "']", Namespaces(pres)) as XmlElement;
            if (sldId == null) return null;
            string relId = sldId.GetAttribute("id", NsR);

            XmlDocument rels = LoadXml(relsEntry);
            var nm = new XmlNamespaceManager(rels.NameTable);
            nm.AddNamespace("r", NsRel);
            foreach (XmlNode node in rels.SelectNodes("//r:Relationship", nm))
            {
                var rel = (XmlElement)node;
                if (rel.GetAttribute("Id") != relId) continue;
                string target = rel.GetAttribute("Target");
                return zip.GetEntry(target.StartsWith("/", StringComparison.Ordinal) ? target.TrimStart('/') : "ppt/" + target);
            }
            return null;
        }

        internal static XmlDocument LoadXml(ZipArchiveEntry entry)
        {
            var doc = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
            using (Stream s = entry.Open())
            using (XmlReader reader = XmlReader.Create(s, settings))
                doc.Load(reader);
            return doc;
        }

        internal static XmlNamespaceManager Namespaces(XmlDocument doc)
        {
            var nm = new XmlNamespaceManager(doc.NameTable);
            nm.AddNamespace("a", NsA);
            nm.AddNamespace("p", NsP);
            return nm;
        }

        /// <summary>A shape's id in the file; the object model's Shape.Id is the same number.</summary>
        internal static bool TryGetId(XmlElement shape, XmlNamespaceManager nm, out int id)
        {
            id = 0;
            var props = shape.SelectSingleNode("*/p:cNvPr", nm) as XmlElement;
            return props != null && int.TryParse(props.GetAttribute("id"), out id);
        }

        /// <summary>The shapes of a slide with these ids, looking inside groups too.</summary>
        private static Dictionary<int, PowerPoint.Shape> FindShapes(PowerPoint.Slide slide, HashSet<int> ids)
        {
            var map = new Dictionary<int, PowerPoint.Shape>();
            if (slide == null) return map;
            PowerPoint.Shapes shapes = slide.Shapes;
            int n = shapes.Count;
            for (int i = 1; i <= n && map.Count < ids.Count; i++)
            {
                PowerPoint.Shape s = shapes[i];
                int id = s.Id;
                if (ids.Contains(id)) { map[id] = s; continue; }

                bool group = false;
                try { group = s.Type == Office.MsoShapeType.msoGroup; } catch { }
                if (!group) continue;
                PowerPoint.GroupShapes items = s.GroupItems;      // already flat: nested groups are expanded
                int count = items.Count;
                for (int j = 1; j <= count; j++)
                {
                    PowerPoint.Shape child = items[j];
                    int childId = child.Id;
                    if (ids.Contains(childId)) map[childId] = child;
                }
            }
            return map;
        }
    }
}
