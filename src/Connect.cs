using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

[assembly: AssemblyTitle("PptSelectionExport")]
[assembly: AssemblyDescription("PowerPoint add-in: export the selected objects as PDF / SVG / PNG / JPG, and align, rotate, merge or restyle them")]
[assembly: AssemblyProduct("PptSelectionExport")]
[assembly: AssemblyVersion("1.7.0.0")]
[assembly: AssemblyFileVersion("1.7.0.0")]
[assembly: ComVisible(false)]

namespace PptSelectionExport
{
    // ---- the classic COM add-in contract (same IID as Extensibility.IDTExtensibility2)
    public enum ext_ConnectMode { ext_cm_AfterStartup = 0, ext_cm_Startup = 1, ext_cm_External = 2, ext_cm_CommandLine = 3 }
    public enum ext_DisconnectMode { ext_dm_HostShutdown = 0, ext_dm_UserClosed = 1 }

    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object Application, [In] ext_ConnectMode ConnectMode, [In, MarshalAs(UnmanagedType.IDispatch)] object AddInInst, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection([In] ext_DisconnectMode RemoveMode, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    /// <summary>The add-in: owns the "Litchi" ribbon tab and reacts to its buttons.</summary>
    [ComVisible(true)]
    [Guid("A0181EA2-7061-43C3-9086-B72894478235")]
    [ProgId("PptSelectionExport.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class Connect : IDTExtensibility2, Office.IRibbonExtensibility
    {
        private enum Kind { Pdf, Svg, Image }

        private static readonly int[] DpiChoices = { 96, 150, 200, 300, 600 };

        private PowerPoint.Application _app;
        private Office.IRibbonUI _ribbon;
        private Settings _settings = Settings.Load();

        // ------------------------------------------------------------ add-in lifetime

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _app = (PowerPoint.Application)Application;
            try
            {
                // follow the display language of Office, not of Windows
                int lcid = _app.LanguageSettings.LanguageID[Office.MsoAppLanguageID.msoLanguageIDUI];
                Strings.Chinese = Strings.IsChineseLcid(lcid);
            }
            catch { }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            _ribbon = null;
            _app = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // ------------------------------------------------------------ ribbon

        public string GetCustomUI(string RibbonID)
        {
            string xml;
            using (Stream s = typeof(Connect).Assembly.GetManifestResourceStream("PptSelectionExport.Ribbon.xml"))
            using (var r = new StreamReader(s, Encoding.UTF8))
                xml = r.ReadToEnd();
            return Regex.Replace(xml, @"\{\{([a-z.]+)\}\}", m => SecurityElement.Escape(Strings.T(m.Groups[1].Value)));
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbon) { _ribbon = ribbon; }

        public object GetImage(Office.IRibbonControl control)
        {
            try
            {
                Bitmap bmp;
                AlignCommand cmd;
                TransformCommand turn;
                MergeCommand merge;
                OrderCommand order;
                if (AlignButtons.TryGetValue(control.Id, out cmd)) bmp = Icons.Align(cmd, 16);
                else if (OrderButtons.TryGetValue(control.Id, out order)) bmp = Icons.Order(order, 16);
                else if (TransformButtons.TryGetValue(control.Id, out turn)) bmp = Icons.Transform(turn, 16);
                else if (MergeButtons.TryGetValue(control.Id, out merge)) bmp = Icons.Merge(merge, 16);
                else switch (control.Id)
                {
                    case "btnPdf": bmp = Icons.Document("PDF", Color.FromArgb(212, 48, 40)); break;
                    case "btnSvg": bmp = Icons.Document("SVG", Color.FromArgb(234, 138, 20)); break;
                    case "btnImage": bmp = Icons.Document("PNG", Color.FromArgb(30, 136, 90)); break;
                    case "btnSlim": bmp = Icons.Slim(); break;
                    case "btnSlideSize": bmp = Icons.SlideSize(); break;
                    case "tglTransparent": bmp = Icons.Checker(32); break;
                    case "tglColor": bmp = Icons.Swatch(_settings.Color, 32); break;
                    case "btnMoreColors": bmp = Icons.Palette(16); break;
                    case "tglAlignToObject": bmp = Icons.AlignToObject(16); break;
                    // shown when a group is collapsed
                    case "grpAlign": bmp = Icons.Align(AlignCommand.Left, 32); break;
                    case "grpRotate": bmp = Icons.Transform(TransformCommand.RotateRight, 32); break;
                    case "grpMerge": bmp = Icons.Merge(MergeCommand.Union, 32); break;
                    case "grpOrder": bmp = Icons.Order(OrderCommand.BringToFront, 32); break;
                    // shape style: the buttons show what they applied last
                    case "galFill": bmp = Icons.StyleColor(false, _settings.FillNone ? (Color?)null : _settings.FillColor, 16); break;
                    case "galOutline": bmp = Icons.StyleColor(true, _settings.OutlineNone ? (Color?)null : _settings.OutlineColor, 16); break;
                    case "galWeight": bmp = Icons.Weight(16); break;
                    case "galDash": bmp = Icons.Dashes(16); break;
                    case "galCap": bmp = Icons.Cap(16); break;
                    case "btnFillNone":
                    case "btnOutlineNone": bmp = Icons.NoColor(16); break;
                    case "btnFillMore":
                    case "btnOutlineMore": bmp = Icons.Palette(16); break;
                    case "grpStyle": bmp = Icons.StyleColor(false, _settings.FillColor, 32); break;
                    default:
                        Color c;
                        if (!Settings.TryParseHex(control.Tag, out c)) return null;
                        bmp = Icons.Swatch(c, 16);
                        break;
                }
                using (bmp) return Icons.ToPictureDisp(bmp);
            }
            catch { return null; }
        }

        public bool GetPressed(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "tglTransparent": return _settings.Transparent;
                case "tglColor": return !_settings.Transparent;
                case "chkOpenFolder": return _settings.OpenFolder;
                case "tglAlignToObject": return _settings.AlignToObject;
            }
            return false;
        }

        public string GetColorSupertip(Office.IRibbonControl control)
        {
            return Strings.T("color.super", Settings.ToHex(_settings.Color));
        }

        /// <summary>The two background buttons behave like radio buttons.</summary>
        public void OnBackgroundMode(Office.IRibbonControl control, bool pressed)
        {
            _settings.Transparent = control.Id == "tglTransparent";
            _settings.Save();
            Refresh();
        }

        public void OnPresetColor(Office.IRibbonControl control)
        {
            Color c;
            if (!Settings.TryParseHex(control.Tag, out c)) return;
            SetColor(c);
        }

        public void OnPickColor(Office.IRibbonControl control)
        {
            IWin32Window owner = Owner();
            using (var dlg = new ColorPickerForm(_settings.Color, _settings.RecentColors, owner))
            {
                if (dlg.ShowDialog(owner) != DialogResult.OK) return;
                _settings.AddRecent(dlg.SelectedColor);
                SetColor(dlg.SelectedColor);
            }
        }

        public int GetDpiIndex(Office.IRibbonControl control)
        {
            int i = Array.IndexOf(DpiChoices, _settings.Dpi);
            return i >= 0 ? i : Array.IndexOf(DpiChoices, 300);
        }

        public void OnDpiChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex < 0 || selectedIndex >= DpiChoices.Length) return;
            _settings.Dpi = DpiChoices[selectedIndex];
            _settings.Save();
        }

        public void OnOption(Office.IRibbonControl control, bool pressed)
        {
            if (control.Id == "chkOpenFolder") _settings.OpenFolder = pressed;
            _settings.Save();
        }

        private void SetColor(Color c)
        {
            _settings.Color = Color.FromArgb(255, c);
            _settings.Transparent = false;     // choosing a colour means "use it"
            _settings.Save();
            Refresh();
        }

        private void Refresh()
        {
            if (_ribbon != null)
            {
                try { _ribbon.Invalidate(); } catch { }
            }
        }

        // ------------------------------------------------------------ export

        public void OnExportPdf(Office.IRibbonControl control) { Export(Kind.Pdf); }
        public void OnExportSvg(Office.IRibbonControl control) { Export(Kind.Svg); }
        public void OnExportImage(Office.IRibbonControl control) { Export(Kind.Image); }

        private void Export(Kind kind)
        {
            IWin32Window owner = Owner();
            try
            {
                PowerPoint.DocumentWindow win;
                PowerPoint.Selection sel;
                bool isChild;
                PowerPoint.ShapeRange range = SelectedShapes(owner, "title", "msg.nowindow", "msg.noselection", out win, out sel, out isChild);
                if (range == null) return;

                PowerPoint.Presentation pres = win.Presentation;
                PowerPoint.Slide slide = null;
                try { slide = sel.SlideRange[1]; } catch { }
                if (kind == Kind.Pdf && slide == null)
                {
                    Info(owner, Strings.T("msg.pdfneedsslide"));
                    return;
                }

                string path = AskPath(owner, kind, pres, slide);
                if (path == null) return;

                var bg = new BackgroundOptions { Transparent = _settings.Transparent, Color = _settings.Color };
                Cursor old = Cursor.Current;
                Cursor.Current = Cursors.WaitCursor;
                try
                {
                    switch (kind)
                    {
                        case Kind.Pdf:
                            Exporter.ExportPdf(_app, pres, slide, range, isChild, path, bg);
                            break;
                        case Kind.Svg:
                            Exporter.ExportSvg(range, path, bg);
                            break;
                        default:
                            string ext = Path.GetExtension(path).ToLowerInvariant();
                            bool jpeg = ext == ".jpg" || ext == ".jpeg";
                            Exporter.ExportImage(range, pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight, path, bg, _settings.Dpi, jpeg);
                            break;
                    }
                }
                finally { Cursor.Current = old; }

                _settings.LastDir = Path.GetDirectoryName(path);
                _settings.Save();
                if (_settings.OpenFolder)
                {
                    try { Process.Start("explorer.exe", "/select,\"" + path + "\""); } catch { }
                }
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                if (ex is IOException || ex is UnauthorizedAccessException || ex is ExternalException) msg += Strings.T("msg.locked");
                MessageBox.Show(owner, Strings.T("msg.failed", msg), Strings.T("title"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private string AskPath(IWin32Window owner, Kind kind, PowerPoint.Presentation pres, PowerPoint.Slide slide)
        {
            string name = Strings.T("name.selection");
            try
            {
                name = Path.GetFileNameWithoutExtension(pres.Name);
                if (slide != null) name += "_P" + slide.SlideIndex;
            }
            catch { }
            foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');

            string dir = _settings.LastDir;
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
            {
                dir = null;
                try
                {
                    string p = pres.Path;
                    if (!string.IsNullOrEmpty(p) && p.IndexOf("://", StringComparison.Ordinal) < 0 && Directory.Exists(p)) dir = p;
                }
                catch { }
                if (dir == null) dir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            }

            using (var dlg = new SaveFileDialog())
            {
                switch (kind)
                {
                    case Kind.Pdf: dlg.Title = Strings.T("dlg.pdf"); dlg.Filter = Strings.T("filter.pdf"); dlg.DefaultExt = "pdf"; break;
                    case Kind.Svg: dlg.Title = Strings.T("dlg.svg"); dlg.Filter = Strings.T("filter.svg"); dlg.DefaultExt = "svg"; break;
                    default:
                        dlg.Title = Strings.T("dlg.img");
                        dlg.Filter = Strings.T("filter.img");
                        dlg.FilterIndex = _settings.ImageJpeg ? 2 : 1;
                        dlg.DefaultExt = _settings.ImageJpeg ? "jpg" : "png";
                        break;
                }
                dlg.AddExtension = true;
                dlg.OverwritePrompt = true;
                dlg.InitialDirectory = dir;
                dlg.FileName = name;
                if (dlg.ShowDialog(owner) != DialogResult.OK) return null;

                string path = dlg.FileName;
                if (kind == Kind.Image)
                {
                    // the chosen file type wins; make the extension agree with it
                    bool jpeg = dlg.FilterIndex == 2;
                    string ext = Path.GetExtension(path).ToLowerInvariant();
                    if (ext == ".png") jpeg = false;
                    else if (ext == ".jpg" || ext == ".jpeg") jpeg = true;
                    else path += jpeg ? ".jpg" : ".png";
                    _settings.ImageJpeg = jpeg;
                }
                return path;
            }
        }

        // ------------------------------------------------------------ slide size

        /// <summary>Asks for a slide size in centimetres and scales what is on the slides to fit it, as one undo step.</summary>
        public void OnSlideSize(Office.IRibbonControl control)
        {
            IWin32Window owner = Owner();
            PowerPoint.Presentation pres = null;
            try { pres = _app.ActivePresentation; } catch { }
            if (pres == null)
            {
                Info(owner, Strings.T("msg.slim.nowindow"), "size.title");
                return;
            }
            try
            {
                float width, height;
                using (var dlg = new SlideSizeForm(pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight, _settings, owner))
                {
                    if (dlg.ShowDialog(owner) != DialogResult.OK) return;
                    width = dlg.WidthPoints;
                    height = dlg.HeightPoints;
                }
                Cursor old = Cursor.Current;
                Cursor.Current = Cursors.WaitCursor;
                try
                {
                    StartUndoEntry();
                    SlideResizer.Resize(pres, width, height);
                }
                finally { Cursor.Current = old; }
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, Strings.T("msg.size.failed", ex.Message), Strings.T("size.title"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ------------------------------------------------------------ slim down

        /// <summary>
        /// Opens the "slim down" window on a temporary copy of the active presentation (as it is
        /// now, unsaved changes included). The window compresses that copy and saves the result
        /// under a new name; the open presentation is not touched.
        /// </summary>
        public void OnSlim(Office.IRibbonControl control)
        {
            IWin32Window owner = Owner();
            PowerPoint.Presentation pres = null;
            try { pres = _app.ActivePresentation; } catch { }
            if (pres == null)
            {
                Info(owner, Strings.T("msg.slim.nowindow"), "slim.title");
                return;
            }

            string dir = Exporter.NewTempDir();
            try
            {
                // a presentation with macros stays one: as .pptx the copy would lose them
                bool macros = false;
                try { macros = pres.HasVBProject; } catch { }
                string input = Path.Combine(dir, macros ? "input.pptm" : "input.pptx");
                Cursor old = Cursor.Current;
                Cursor.Current = Cursors.WaitCursor;
                try
                {
                    pres.SaveCopyAs(input, macros ? PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentationMacroEnabled
                                                  : PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Office.MsoTriState.msoFalse);
                }
                finally { Cursor.Current = old; }

                string name = Strings.T("name.selection"), original = "", folder = null;
                try { name = Path.GetFileNameWithoutExtension(pres.Name); } catch { }
                foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
                try
                {
                    string p = pres.Path;       // empty for a presentation that was never saved, a URL for one in the cloud
                    if (!string.IsNullOrEmpty(p) && p.IndexOf("://", StringComparison.Ordinal) < 0 && Directory.Exists(p))
                    {
                        folder = p;
                        original = Path.GetFullPath(pres.FullName);
                    }
                }
                catch { }
                if (folder == null)
                    folder = !string.IsNullOrEmpty(_settings.LastDir) && Directory.Exists(_settings.LastDir) ? _settings.LastDir
                           : Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

                using (var dlg = new SlimForm(input, dir, name, folder, original, _settings, owner)) dlg.ShowDialog(owner);
            }
            catch (InvalidDataException)
            {
                // the copy is not a zip package: PowerPoint encrypts a presentation that has a password to open
                Info(owner, Strings.T("err.cap.copy"), "slim.title");
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, Strings.T("msg.slim.failed", ex.Message), Strings.T("slim.title"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { Exporter.TryDeleteDir(dir); }
        }

        // ------------------------------------------------------------ align

        private static readonly Dictionary<string, AlignCommand> AlignButtons = new Dictionary<string, AlignCommand>
        {
            { "btnAlignLeft", AlignCommand.Left },
            { "btnAlignCenter", AlignCommand.Center },
            { "btnAlignRight", AlignCommand.Right },
            { "btnAlignTop", AlignCommand.Top },
            { "btnAlignMiddle", AlignCommand.Middle },
            { "btnAlignBottom", AlignCommand.Bottom },
            { "btnDistributeH", AlignCommand.DistributeHorizontally },
            { "btnDistributeV", AlignCommand.DistributeVertically },
        };

        public void OnAlign(Office.IRibbonControl control)
        {
            AlignCommand cmd;
            if (AlignButtons.TryGetValue(control.Id, out cmd)) Align(cmd);
        }

        public void OnAlignToObject(Office.IRibbonControl control, bool pressed)
        {
            _settings.AlignToObject = pressed;
            _settings.Save();
        }

        private void Align(AlignCommand cmd)
        {
            IWin32Window owner = Owner();
            try
            {
                PowerPoint.DocumentWindow win;
                PowerPoint.Selection sel;
                bool isChild;
                PowerPoint.ShapeRange range = SelectedShapes(owner, "title.align", "msg.align.nowindow", "msg.align.noselection", out win, out sel, out isChild);
                if (range == null) return;

                PowerPoint.PageSetup page = win.Presentation.PageSetup;
                StartUndoEntry();
                int unmoved;
                AlignResult result = Aligner.Apply(range, cmd, _settings.AlignToObject, page.SlideWidth, page.SlideHeight, out unmoved);
                if (result == AlignResult.NeedThree) Info(owner, Strings.T("msg.align.needthree"), "title.align");
                else if (unmoved > 0) Info(owner, Strings.T("msg.align.unmoved", unmoved), "title.align");
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, Strings.T("msg.align.failed", ex.Message), Strings.T("title.align"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ------------------------------------------------------------ rotate / flip, merge shapes

        private static readonly Dictionary<string, TransformCommand> TransformButtons = new Dictionary<string, TransformCommand>
        {
            { "btnRotateLeft", TransformCommand.RotateLeft },
            { "btnRotateRight", TransformCommand.RotateRight },
            { "btnFlipH", TransformCommand.FlipHorizontal },
            { "btnFlipV", TransformCommand.FlipVertical },
        };

        private static readonly Dictionary<string, MergeCommand> MergeButtons = new Dictionary<string, MergeCommand>
        {
            { "btnUnion", MergeCommand.Union },
            { "btnCombine", MergeCommand.Combine },
            { "btnSubtract", MergeCommand.Subtract },
            { "btnFragment", MergeCommand.Fragment },
        };

        public void OnTransform(Office.IRibbonControl control)
        {
            TransformCommand cmd;
            if (!TransformButtons.TryGetValue(control.Id, out cmd)) return;
            IWin32Window owner = Owner();
            try
            {
                PowerPoint.DocumentWindow win;
                PowerPoint.Selection sel;
                bool isChild;
                PowerPoint.ShapeRange range = SelectedShapes(owner, "title.rotate", "msg.shape.nowindow", "msg.shape.noselection", out win, out sel, out isChild);
                if (range == null) return;
                StartUndoEntry();
                ShapeTools.Transform(range, cmd);
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, Strings.T("msg.shape.failed", ex.Message), Strings.T("title.rotate"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static readonly Dictionary<string, OrderCommand> OrderButtons = new Dictionary<string, OrderCommand>
        {
            { "btnBringToFront", OrderCommand.BringToFront },
            { "btnBringForward", OrderCommand.BringForward },
            { "btnSendToBack", OrderCommand.SendToBack },
            { "btnSendBackward", OrderCommand.SendBackward },
        };

        /// <summary>Stacking order: to front, to back, one step forward, one step backward.</summary>
        public void OnOrder(Office.IRibbonControl control)
        {
            OrderCommand cmd;
            if (!OrderButtons.TryGetValue(control.Id, out cmd)) return;
            IWin32Window owner = Owner();
            bool isChild = false;
            try
            {
                PowerPoint.DocumentWindow win;
                PowerPoint.Selection sel;
                PowerPoint.ShapeRange range = SelectedShapes(owner, "title.order", "msg.shape.nowindow", "msg.shape.noselection", out win, out sel, out isChild);
                if (range == null) return;
                StartUndoEntry();
                // The object model cannot reorder shapes inside a group; PowerPoint's own command can.
                if (isChild) _app.CommandBars.ExecuteMso(ShapeTools.OrderMso(cmd));
                else ShapeTools.Order(range, cmd);
            }
            catch (Exception ex)
            {
                string text = isChild ? Strings.T("msg.order.child") : Strings.T("msg.shape.failed", ex.Message);
                MessageBox.Show(owner, text, Strings.T("title.order"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void OnMerge(Office.IRibbonControl control)
        {
            MergeCommand cmd;
            if (!MergeButtons.TryGetValue(control.Id, out cmd)) return;
            IWin32Window owner = Owner();
            MergeResult result;
            try
            {
                PowerPoint.DocumentWindow win;
                PowerPoint.Selection sel;
                bool isChild;
                PowerPoint.ShapeRange range = SelectedShapes(owner, "title.merge", "msg.shape.nowindow", "msg.shape.noselection", out win, out sel, out isChild);
                if (range == null) return;
                PowerPoint.Shapes shapes = null;
                try { shapes = sel.SlideRange[1].Shapes; } catch { }      // not on a slide (a master, say): the result is not selected
                StartUndoEntry();
                result = ShapeTools.Merge(range, isChild, cmd, shapes);
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, Strings.T("msg.merge.failed", ex.Message), Strings.T("title.merge"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (result == MergeResult.NeedTwo) Info(owner, Strings.T("msg.merge.needtwo"), "title.merge");
            else if (result == MergeResult.Grouped) Info(owner, Strings.T("msg.merge.grouped"), "title.merge");
        }

        // ------------------------------------------------------------ shape style: fill, outline, weight, dashes

        private const int ColorCellSize = 18, LineItemWidth = 96, LineItemHeight = 18;   // as in Ribbon.xml

        /// <summary>The colour grid of the fill and outline lists: the picker's palette, then the colours used recently.</summary>
        private List<Color> StyleColors()
        {
            var list = new List<Color>();
            foreach (string hex in ColorPickerForm.Palette)
            {
                Color c;
                if (Settings.TryParseHex(hex, out c)) list.Add(c);
            }
            foreach (Color recent in _settings.RecentColors)
            {
                bool known = false;
                foreach (Color c in list) if (c.ToArgb() == Color.FromArgb(255, recent).ToArgb()) { known = true; break; }
                if (!known) list.Add(Color.FromArgb(255, recent));
            }
            return list;
        }

        public int GetColorCount(Office.IRibbonControl control) { return StyleColors().Count; }

        public string GetColorId(Office.IRibbonControl control, int index) { return control.Id + "Color" + index; }

        public object GetColorImage(Office.IRibbonControl control, int index)
        {
            try
            {
                List<Color> colors = StyleColors();
                if (index < 0 || index >= colors.Count) return null;
                using (Bitmap bmp = Icons.ColorCell(colors[index], ColorCellSize)) return Icons.ToPictureDisp(bmp);
            }
            catch { return null; }
        }

        public string GetColorTip(Office.IRibbonControl control, int index)
        {
            List<Color> colors = StyleColors();
            return index >= 0 && index < colors.Count ? "#" + Settings.ToHex(colors[index]) : "";
        }

        /// <summary>Marks the colour the button applied last, when it is in the grid.</summary>
        public int GetColorIndex(Office.IRibbonControl control)
        {
            bool outline = control.Id == "galOutline";
            if (outline ? _settings.OutlineNone : _settings.FillNone) return -1;
            int argb = Color.FromArgb(255, outline ? _settings.OutlineColor : _settings.FillColor).ToArgb();
            List<Color> colors = StyleColors();
            for (int i = 0; i < colors.Count; i++) if (colors[i].ToArgb() == argb) return i;
            return -1;
        }

        public void OnColorPicked(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            List<Color> colors = StyleColors();
            if (selectedIndex < 0 || selectedIndex >= colors.Count) return;
            ApplyColor(control.Id == "galOutline", colors[selectedIndex]);
        }

        /// <summary>The buttons under the lists: no fill / no outline, more colours, more weights.</summary>
        public void OnStyleButton(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "btnFillNone": ApplyColor(false, null); break;
                case "btnOutlineNone": ApplyColor(true, null); break;
                case "btnFillMore": PickStyleColor(false); break;
                case "btnOutlineMore": PickStyleColor(true); break;
                case "btnWeightMore":
                    IWin32Window owner = Owner();
                    using (var dlg = new WeightForm(_settings.OutlineWeight, owner))
                    {
                        if (dlg.ShowDialog(owner) != DialogResult.OK) return;
                        ApplyWeight(dlg.Weight);
                    }
                    break;
            }
        }

        private void PickStyleColor(bool outline)
        {
            IWin32Window owner = Owner();
            Color current = outline ? _settings.OutlineColor : _settings.FillColor;
            using (var dlg = new ColorPickerForm(current, _settings.RecentColors, owner, Strings.T(outline ? "picker.outline" : "picker.fill")))
            {
                if (dlg.ShowDialog(owner) != DialogResult.OK) return;
                _settings.AddRecent(dlg.SelectedColor);
                ApplyColor(outline, dlg.SelectedColor);
            }
        }

        /// <param name="color">Null = no fill / no outline.</param>
        private void ApplyColor(bool outline, Color? color)
        {
            // the button remembers a colour only once it has really been applied
            if (ApplyStyle(range => outline ? StyleTools.SetOutlineColor(range, color) : StyleTools.SetFill(range, color)))
            {
                if (outline)
                {
                    _settings.OutlineNone = !color.HasValue;
                    if (color.HasValue) _settings.OutlineColor = Color.FromArgb(255, color.Value);
                }
                else
                {
                    _settings.FillNone = !color.HasValue;
                    if (color.HasValue) _settings.FillColor = Color.FromArgb(255, color.Value);
                }
            }
            _settings.Save();
            RefreshStyle("galFill", "galOutline");      // both: the recent colours are shared
        }

        public int GetWeightCount(Office.IRibbonControl control) { return StyleTools.Weights.Length; }
        public string GetWeightId(Office.IRibbonControl control, int index) { return "weight" + index; }
        public string GetWeightLabel(Office.IRibbonControl control, int index) { return StyleTools.FormatWeight(StyleTools.Weights[index]); }

        public object GetWeightImage(Office.IRibbonControl control, int index)
        {
            try
            {
                using (Bitmap bmp = Icons.LinePreview(LineItemWidth, LineItemHeight, StyleTools.Weights[index], DashKind.Solid)) return Icons.ToPictureDisp(bmp);
            }
            catch { return null; }
        }

        public int GetWeightIndex(Office.IRibbonControl control)
        {
            for (int i = 0; i < StyleTools.Weights.Length; i++)
                if (Math.Abs(StyleTools.Weights[i] - _settings.OutlineWeight) < 0.005f) return i;
            return -1;
        }

        public void OnWeightPicked(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex >= 0 && selectedIndex < StyleTools.Weights.Length) ApplyWeight(StyleTools.Weights[selectedIndex]);
        }

        private void ApplyWeight(float points)
        {
            if (ApplyStyle(range => StyleTools.SetWeight(range, points)))
            {
                _settings.OutlineWeight = points;
                _settings.Save();
            }
            RefreshStyle("galWeight");
        }

        private static readonly string[] DashKeys =
        {
            "dash.solid", "dash.rounddot", "dash.squaredot", "dash.dash", "dash.dashdot", "dash.longdash", "dash.longdashdot", "dash.longdashdotdot"
        };

        public int GetDashCount(Office.IRibbonControl control) { return StyleTools.Dashes.Length; }
        public string GetDashId(Office.IRibbonControl control, int index) { return "dash" + index; }
        public string GetDashLabel(Office.IRibbonControl control, int index) { return Strings.T(DashKeys[index]); }
        public int GetDashIndex(Office.IRibbonControl control) { return _settings.OutlineDash; }

        public object GetDashImage(Office.IRibbonControl control, int index)
        {
            try
            {
                using (Bitmap bmp = Icons.LinePreview(LineItemWidth, LineItemHeight, 1.5f, StyleTools.Dashes[index])) return Icons.ToPictureDisp(bmp);
            }
            catch { return null; }
        }

        public void OnDashPicked(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex < 0 || selectedIndex >= StyleTools.Dashes.Length) return;
            DashKind dash = StyleTools.Dashes[selectedIndex];
            if (ApplyStyle(range => StyleTools.SetDash(range, dash)))
            {
                _settings.OutlineDash = selectedIndex;
                _settings.Save();
            }
            RefreshStyle("galDash");
        }

        private static readonly string[] CapKeys = { "cap.square", "cap.round", "cap.flat" };

        public int GetCapCount(Office.IRibbonControl control) { return LineCaps.Caps.Length; }
        public string GetCapId(Office.IRibbonControl control, int index) { return "cap" + index; }
        public string GetCapLabel(Office.IRibbonControl control, int index) { return Strings.T(CapKeys[index]); }
        public int GetCapIndex(Office.IRibbonControl control) { return _settings.OutlineCap; }

        public object GetCapImage(Office.IRibbonControl control, int index)
        {
            try
            {
                using (Bitmap bmp = Icons.CapPreview(LineItemWidth, LineItemHeight, LineCaps.Caps[index])) return Icons.ToPictureDisp(bmp);
            }
            catch { return null; }
        }

        public void OnCapPicked(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex < 0 || selectedIndex >= LineCaps.Caps.Length) return;
            CapKind cap = LineCaps.Caps[selectedIndex];
            bool applied = ApplyStyle(range =>
            {
                // it works through a temporary copy of the presentation, which takes a moment for a large file
                Cursor old = Cursor.Current;
                Cursor.Current = Cursors.WaitCursor;
                try { return LineCaps.SetCap(_app, range, cap); }
                finally { Cursor.Current = old; }
            });
            if (applied)
            {
                _settings.OutlineCap = selectedIndex;
                _settings.Save();
            }
            RefreshStyle("galCap");
        }

        /// <summary>
        /// Applies one style setting to the selected shapes, as a single undo step.
        /// False (after telling the user why) when nothing was changed.
        /// </summary>
        private bool ApplyStyle(Func<PowerPoint.ShapeRange, int> change)
        {
            IWin32Window owner = Owner();
            try
            {
                PowerPoint.DocumentWindow win;
                PowerPoint.Selection sel;
                bool isChild;
                PowerPoint.ShapeRange range = SelectedShapes(owner, "title.style", "msg.shape.nowindow", "msg.shape.noselection", out win, out sel, out isChild);
                if (range == null) return false;
                StartUndoEntry();
                if (change(range) > 0) return true;
                Info(owner, Strings.T("msg.style.none"), "title.style");
            }
            catch (StyleNotice notice)
            {
                Info(owner, notice.Message, "title.style");
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, Strings.T("msg.shape.failed", ex.Message), Strings.T("title.style"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return false;
        }

        private void RefreshStyle(params string[] controlIds)
        {
            if (_ribbon == null) return;
            foreach (string id in controlIds)
            {
                try { _ribbon.InvalidateControl(id); } catch { }
            }
        }

        // ------------------------------------------------------------ helpers

        /// <summary>So that one Undo takes back the whole command, and nothing the user did before it.</summary>
        private void StartUndoEntry()
        {
            try { _app.StartNewUndoEntry(); } catch { }
        }

        /// <summary>
        /// The shapes selected in the active window (or inside a group, when the selection is there),
        /// in the order they were selected. Null, after telling the user what to do, when there are none.
        /// </summary>
        private PowerPoint.ShapeRange SelectedShapes(IWin32Window owner, string titleKey, string noWindowKey, string noSelectionKey,
                                                     out PowerPoint.DocumentWindow win, out PowerPoint.Selection sel, out bool isChild)
        {
            sel = null;
            isChild = false;
            try { win = _app.ActiveWindow; }
            catch { win = null; }
            if (win == null)
            {
                Info(owner, Strings.T(noWindowKey), titleKey);
                return null;
            }

            sel = win.Selection;
            PowerPoint.PpSelectionType type = sel.Type;
            if (type != PowerPoint.PpSelectionType.ppSelectionShapes && type != PowerPoint.PpSelectionType.ppSelectionText)
            {
                Info(owner, Strings.T(noSelectionKey), titleKey);
                return null;
            }

            try { isChild = sel.HasChildShapeRange; } catch { }
            PowerPoint.ShapeRange range = isChild ? sel.ChildShapeRange : sel.ShapeRange;
            if (range.Count == 0)
            {
                Info(owner, Strings.T(noSelectionKey), titleKey);
                return null;
            }
            return range;
        }

        private static void Info(IWin32Window owner, string text, string titleKey = "title")
        {
            MessageBox.Show(owner, text, Strings.T(titleKey), MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private IWin32Window Owner()
        {
            try { return new WindowHandle(new IntPtr(_app.HWND)); }
            catch { return null; }
        }

        private sealed class WindowHandle : IWin32Window
        {
            private readonly IntPtr _h;
            public WindowHandle(IntPtr h) { _h = h; }
            public IntPtr Handle { get { return _h; } }
        }
    }
}
