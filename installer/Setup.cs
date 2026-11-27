// Litchi-Setup.exe: a single-file installer / uninstaller for the add-in.
//
// It needs no administrator rights: files go to a folder of the user's choice
// (default %LOCALAPPDATA%\PptSelectionExport) and the registration to HKEY_CURRENT_USER.
// The add-in DLL travels inside this exe as an embedded resource. A copy of the exe is kept
// next to it as Uninstall.exe so the add-in can be removed from Windows "Installed apps".
//
//   Litchi-Setup.exe                    interactive
//   Litchi-Setup.exe /S                 install silently to the default folder
//   Litchi-Setup.exe /S /DIR="D:\Apps\Litchi"  install silently to that folder
//   Litchi-Setup.exe /uninstall         uninstall (add /S for silent)
//
// Exit codes: 0 = done, 1 = cancelled, 2 = PowerPoint is running (silent mode), 3 = error.
using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("PptSelectionExport Setup")]
[assembly: AssemblyDescription("Installer for the PowerPoint add-in \"Export Selection\"")]
[assembly: AssemblyProduct("PptSelectionExport")]
[assembly: AssemblyVersion("1.4.0.0")]
[assembly: AssemblyFileVersion("1.4.0.0")]

namespace PptSelectionExportSetup
{
    internal static class Program
    {
        // Chinese when either the Windows display language or the regional format is Chinese
        // (many Chinese-speaking users run an English Windows with Chinese regional settings).
        internal static readonly bool Zh =
            CultureInfo.CurrentUICulture.Name.StartsWith("zh", StringComparison.OrdinalIgnoreCase) ||
            CultureInfo.CurrentCulture.Name.StartsWith("zh", StringComparison.OrdinalIgnoreCase);
        internal static string T(string zh, string en) { return Zh ? zh : en; }

        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [STAThread]
        private static int Main(string[] args)
        {
            bool silent = false, uninstall = false;
            string dir = null, preview = null;
            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i];
                string s = a.TrimStart('/', '-');
                string lower = s.ToLowerInvariant();
                if (lower == "s" || lower == "silent" || lower == "quiet") silent = true;
                else if (lower == "uninstall" || lower == "u") uninstall = true;
                else if (lower.StartsWith("dir=")) dir = s.Substring(4).Trim('"');
                else if (lower == "preview" && i + 1 < args.Length) preview = args[++i];   // developer aid: render the pages to PNG files
            }

            try
            {
                if (silent)
                {
                    if (Installer.PowerPointRunning()) return 2;
                    if (uninstall) Installer.Uninstall();
                    else Installer.Install(Installer.NormalizeDir(dir ?? Installer.SuggestedDir()));
                    return 0;
                }

                try { SetProcessDPIAware(); } catch { }
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                if (preview != null) { SetupForm.RenderPreviews(preview); return 0; }

                using (var form = new SetupForm(uninstall))
                {
                    Application.Run(form);
                    return form.Completed ? 0 : 1;
                }
            }
            catch (Exception ex)
            {
                if (!silent)
                    MessageBox.Show(T("操作失败：\n", "Setup failed:\n") + ex.Message, T("选区导出 安装程序", "Export Selection Setup"),
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 3;
            }
        }
    }

    // ====================================================================== what setup does

    internal static class Installer
    {
        private const string ProgId = "PptSelectionExport.Connect";
        private const string Clsid = "{A0181EA2-7061-43C3-9086-B72894478235}";
        private const string FolderName = "PptSelectionExport";
        private const string DllName = "PptSelectionExport.dll";
        private const string UninstallerName = "Uninstall.exe";
        private const string AddinKey = @"Software\Microsoft\Office\PowerPoint\Addins\" + ProgId;
        private const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\PptSelectionExport";
        private const string SettingsKey = @"Software\PptSelectionExport";
        private const string ManagedCategory = "{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}";
        private const string HomePage = "https://github.com/Avopears/litchi-ppt-tool";
        private static readonly string[] ClassRoots = { @"Software\Classes", @"Software\Classes\Wow6432Node" };

        public static string DefaultDir
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), FolderName); }
        }

        /// <summary>Version of the add-in carried inside this installer, e.g. "1.2.0".</summary>
        public static string Version
        {
            get { return typeof(Installer).Assembly.GetName().Version.ToString(3); }
        }

        public static bool IsInstalled()
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(AddinKey)) return k != null;
        }

        public static bool PowerPointRunning()
        {
            return Process.GetProcessesByName("POWERPNT").Length > 0;
        }

        /// <summary>Where the add-in currently lives, or null when that is unknown.</summary>
        public static string InstalledDir()
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(UninstallKey))
                {
                    string d = k == null ? null : k.GetValue("InstallLocation") as string;
                    if (!string.IsNullOrEmpty(d)) return d;
                }
                // installed by install.ps1: no "Installed apps" entry, but the COM registration knows the file
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\Classes\CLSID\" + Clsid + @"\InprocServer32"))
                {
                    string cb = k == null ? null : k.GetValue("CodeBase") as string;
                    if (!string.IsNullOrEmpty(cb)) return Path.GetDirectoryName(new Uri(cb).LocalPath);
                }
            }
            catch { }
            return null;
        }

        /// <summary>The folder to offer: where it is installed now, otherwise the default.</summary>
        public static string SuggestedDir()
        {
            return InstalledDir() ?? DefaultDir;
        }

        /// <summary>
        /// Turns what the user typed or picked into the real target folder. A folder that is new,
        /// empty, or already holds the add-in is used as it is. A folder with other things in it
        /// (say "D:\Tools") gets a subfolder, "D:\Tools\PptSelectionExport", so the add-in's files
        /// never mix with someone else's.
        /// </summary>
        public static string NormalizeDir(string input)
        {
            string p = Environment.ExpandEnvironmentVariables((input ?? "").Trim().Trim('"'));
            if (p.Length == 0) throw new ArgumentException(Program.T("请填写安装位置。", "Enter a folder to install to."));
            if (!Path.IsPathRooted(p) || p.StartsWith(@"\\?\", StringComparison.Ordinal))
                throw new ArgumentException(Program.T("安装位置需要是完整路径，例如 D:\\Tools。", "The install folder must be a full path, for example D:\\Tools."));
            p = Path.GetFullPath(p).TrimEnd('\\', '/');
            if (p.EndsWith(":", StringComparison.Ordinal)) p += "\\";      // a drive root such as "D:"
            if (!string.Equals(Path.GetFileName(p), FolderName, StringComparison.OrdinalIgnoreCase) && !IsOursOrEmpty(p)) p = Path.Combine(p, FolderName);
            return p;
        }

        private static bool IsOursOrEmpty(string dir)
        {
            try
            {
                if (Path.GetPathRoot(dir).TrimEnd('\\').Equals(dir.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase)) return false;   // never a drive root
                if (!Directory.Exists(dir)) return !File.Exists(dir);
                foreach (string entry in Directory.GetFileSystemEntries(dir))
                {
                    string name = Path.GetFileName(entry);
                    if (!name.Equals(DllName, StringComparison.OrdinalIgnoreCase) && !name.Equals(UninstallerName, StringComparison.OrdinalIgnoreCase)) return false;
                }
                return true;
            }
            catch { return false; }
        }

        private static bool SamePath(string a, string b)
        {
            return string.Equals(Path.GetFullPath(a).TrimEnd('\\'), Path.GetFullPath(b).TrimEnd('\\'), StringComparison.OrdinalIgnoreCase);
        }

        // ------------------------------------------------------------------ install

        /// <summary>The installation as separate steps, so the window can show progress between them.</summary>
        public static Action[] InstallSteps(string dir)
        {
            string dll = Path.Combine(dir, DllName);
            string uninstaller = Path.Combine(dir, UninstallerName);
            string previous = InstalledDir();
            return new Action[]
            {
                delegate   // an older copy is removed completely first, so nothing of it can linger:
                {          // its registration, its "Installed apps" entry and its files (saved settings stay)
                    if (previous == null && !IsInstalled()) return;
                    RemoveRegistration();
                    if (previous != null && Directory.Exists(previous)) RemovePreviousFiles(previous, dir);
                },
                delegate   // files
                {
                    try { Directory.CreateDirectory(dir); }
                    catch (UnauthorizedAccessException)
                    {
                        throw new UnauthorizedAccessException(Program.T(
                            "没有权限写入这个位置：\n" + dir + "\n请换一个文件夹（例如 D 盘或“文档”里的文件夹）。",
                            "No permission to write to:\n" + dir + "\nChoose another folder (on another drive or inside Documents, for example)."));
                    }
                    using (Stream src = typeof(Installer).Assembly.GetManifestResourceStream(DllName))
                    {
                        if (src == null) throw new InvalidOperationException("The installer is damaged: the add-in is missing from it.");
                        try
                        {
                            using (var dst = new FileStream(dll, FileMode.Create, FileAccess.Write)) src.CopyTo(dst);
                        }
                        catch (UnauthorizedAccessException)
                        {
                            throw new UnauthorizedAccessException(Program.T(
                                "没有权限写入这个位置：\n" + dir + "\n请换一个文件夹（例如 D 盘或“文档”里的文件夹）。",
                                "No permission to write to:\n" + dir + "\nChoose another folder (on another drive or inside Documents, for example)."));
                        }
                    }
                    string self = Assembly.GetExecutingAssembly().Location;
                    if (!SamePath(self, uninstaller))
                    {
                        File.Copy(self, uninstaller, true);
                        try { DeleteFile(uninstaller + ":Zone.Identifier"); } catch { }   // drop the "downloaded from the Internet" mark
                    }
                },
                delegate   // COM registration of the managed class (what "regasm /codebase" writes), per user,
                {          // in the 64-bit and the 32-bit view so either edition of Office finds it
                    AssemblyName name = AssemblyName.GetAssemblyName(dll);
                    string codeBase = "file:///" + dll.Replace('\\', '/');
                    foreach (string root in ClassRoots)
                    {
                        string k = root + @"\CLSID\" + Clsid;
                        Registry.CurrentUser.DeleteSubKeyTree(k, false);
                        SetDefault(k, ProgId);
                        SetDefault(k + @"\ProgId", ProgId);
                        using (Registry.CurrentUser.CreateSubKey(k + @"\Implemented Categories\" + ManagedCategory)) { }
                        foreach (string sub in new[] { k + @"\InprocServer32", k + @"\InprocServer32\" + name.Version })
                        {
                            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(sub))
                            {
                                key.SetValue("Class", ProgId);
                                key.SetValue("Assembly", name.FullName);
                                key.SetValue("RuntimeVersion", "v4.0.30319");
                                key.SetValue("CodeBase", codeBase);
                            }
                        }
                        using (RegistryKey key = Registry.CurrentUser.CreateSubKey(k + @"\InprocServer32"))
                        {
                            key.SetValue("", "mscoree.dll");
                            key.SetValue("ThreadingModel", "Both");
                        }
                    }
                    SetDefault(@"Software\Classes\" + ProgId, ProgId);
                    SetDefault(@"Software\Classes\" + ProgId + @"\CLSID", Clsid);
                },
                delegate   // tell PowerPoint to load it at startup
                {
                    using (RegistryKey key = Registry.CurrentUser.CreateSubKey(AddinKey))
                    {
                        key.SetValue("FriendlyName", "选区导出 Export Selection (PDF / SVG / PNG / JPG)");
                        key.SetValue("Description", "把选中的对象按最小外接矩形导出为 PDF、SVG、PNG 或 JPG，可设置背景颜色或透明。 Export the selected objects, cropped to their bounding box, as PDF, SVG, PNG or JPG.");
                        key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                    }
                },
                delegate   // entry in Windows "Installed apps"
                {
                    AssemblyName name = AssemblyName.GetAssemblyName(dll);
                    using (RegistryKey key = Registry.CurrentUser.CreateSubKey(UninstallKey))
                    {
                        key.SetValue("DisplayName", "选区导出 Export Selection for PowerPoint");
                        key.SetValue("DisplayVersion", name.Version.ToString(3));
                        key.SetValue("Publisher", "Avopears");
                        key.SetValue("URLInfoAbout", HomePage);
                        key.SetValue("InstallLocation", dir);
                        key.SetValue("DisplayIcon", uninstaller);
                        key.SetValue("UninstallString", "\"" + uninstaller + "\" /uninstall");
                        key.SetValue("QuietUninstallString", "\"" + uninstaller + "\" /uninstall /S");
                        key.SetValue("NoModify", 1, RegistryValueKind.DWord);
                        key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                        long kb = (new FileInfo(dll).Length + new FileInfo(uninstaller).Length) / 1024;
                        key.SetValue("EstimatedSize", (int)kb, RegistryValueKind.DWord);
                    }
                }
            };
        }

        /// <summary>What to show while each of <see cref="InstallSteps"/> runs.</summary>
        public static string[] InstallLabels()
        {
            bool upgrade = IsInstalled() || InstalledDir() != null;
            return new[]
            {
                upgrade ? Program.T("正在卸载旧版本…", "Removing the previous version…") : Program.T("正在准备…", "Preparing…"),
                Program.T("正在复制文件…", "Copying files…"),
                Program.T("正在注册插件…", "Registering the add-in…"),
                Program.T("正在注册插件…", "Registering the add-in…"),
                Program.T("正在完成安装…", "Finishing…")
            };
        }

        public static void Install(string dir)
        {
            foreach (Action step in InstallSteps(dir)) step();
        }

        /// <summary>Version of the copy that is installed now, e.g. "1.3.0", or null when unknown.</summary>
        public static string InstalledVersion()
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(UninstallKey))
                {
                    string v = k == null ? null : k.GetValue("DisplayVersion") as string;
                    if (!string.IsNullOrEmpty(v)) return v;
                }
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\Classes\CLSID\" + Clsid + @"\InprocServer32"))
                {
                    string asm = k == null ? null : k.GetValue("Assembly") as string;
                    if (asm != null)
                    {
                        System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(asm, @"Version=(\d+\.\d+\.\d+)");
                        if (m.Success) return m.Groups[1].Value;
                    }
                }
            }
            catch { }
            return null;
        }

        /// <summary>Everything in the registry that makes PowerPoint and Windows know the add-in (not its saved settings).</summary>
        private static void RemoveRegistration()
        {
            Registry.CurrentUser.DeleteSubKeyTree(AddinKey, false);
            Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\" + ProgId, false);
            foreach (string root in ClassRoots) Registry.CurrentUser.DeleteSubKeyTree(root + @"\CLSID\" + Clsid, false);
            Registry.CurrentUser.DeleteSubKeyTree(UninstallKey, false);
        }

        /// <summary>
        /// Deletes the files of the copy being replaced. Its folder goes too when the new copy
        /// is installed somewhere else and nothing else is left in the old one.
        /// </summary>
        private static void RemovePreviousFiles(string previous, string target)
        {
            try { File.Delete(Path.Combine(previous, DllName)); } catch { }
            string oldUninstaller = Path.Combine(previous, UninstallerName);
            if (!SamePath(Assembly.GetExecutingAssembly().Location, oldUninstaller))
            {
                try { File.Delete(oldUninstaller); } catch { }
            }
            if (!SamePath(previous, target))
            {
                try { Directory.Delete(previous, false); } catch { }
            }
        }

        private static void SetDefault(string subKey, string value)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(subKey)) key.SetValue("", value);
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool DeleteFile(string name);

        // ------------------------------------------------------------------ uninstall

        public static Action[] UninstallSteps()
        {
            string dir = InstalledDir() ?? DefaultDir;
            return new Action[]
            {
                delegate { RemoveRegistration(); },
                delegate { Registry.CurrentUser.DeleteSubKeyTree(SettingsKey, false); },
                delegate { if (Directory.Exists(dir)) RemoveFiles(dir); }
            };
        }

        public static string[] UninstallLabels()
        {
            string s = Program.T("正在卸载…", "Removing…");
            return new[] { s, s, s };
        }

        public static void Uninstall()
        {
            foreach (Action step in UninstallSteps()) step();
        }

        /// <summary>
        /// Removes our two files and then the folder, but only if nothing else is left in it:
        /// the folder was chosen by the user, so anything we did not put there stays.
        /// </summary>
        private static void RemoveFiles(string dir)
        {
            try { File.Delete(Path.Combine(dir, DllName)); } catch { }
            string uninstaller = Path.Combine(dir, UninstallerName);
            string self = Assembly.GetExecutingAssembly().Location;
            if (!SamePath(self, uninstaller))
            {
                try { File.Delete(uninstaller); } catch { }
                try { Directory.Delete(dir, false); } catch { }
                return;
            }
            // We are that Uninstall.exe and cannot delete ourselves: leave it to a short-lived
            // command prompt that waits for this process to end.
            try
            {
                var psi = new ProcessStartInfo("cmd.exe",
                    "/c ping 127.0.0.1 -n 3 >nul & del /f /q \"" + uninstaller + "\" & rmdir \"" + dir + "\"");
                psi.CreateNoWindow = true;
                psi.UseShellExecute = false;
                psi.WorkingDirectory = Path.GetTempPath();
                Process.Start(psi);
            }
            catch { }
        }
    }

    // ====================================================================== the window

    internal static class Ui
    {
        public static float Scale = 1f;
        public static int S(double v) { return (int)Math.Round(v * Scale); }

        public static readonly Color Accent = Color.FromArgb(196, 62, 28);          // PowerPoint's orange-red
        public static readonly Color AccentHover = Color.FromArgb(172, 52, 22);
        public static readonly Color AccentPressed = Color.FromArgb(146, 44, 18);
        public static readonly Color Text = Color.FromArgb(27, 27, 27);
        public static readonly Color TextSoft = Color.FromArgb(96, 96, 96);
        public static readonly Color Border = Color.FromArgb(208, 208, 208);
        public static readonly Color Line = Color.FromArgb(229, 229, 229);
        public static readonly Color Bar = Color.FromArgb(246, 246, 246);
        public static readonly Color Danger = Color.FromArgb(196, 43, 28);
        public static readonly Color Success = Color.FromArgb(16, 124, 16);

        public static Font Font(float points, FontStyle style)
        {
            return new Font(SystemFonts.MessageBoxFont.FontFamily, points, style, GraphicsUnit.Point);
        }

        public static GraphicsPath Round(RectangleF r, float radius)
        {
            float d = radius * 2;
            var p = new GraphicsPath();
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        public static void Quality(Graphics g)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
        }
    }

    /// <summary>A flat, rounded button: filled with the accent colour (primary) or outlined (secondary).</summary>
    internal sealed class FlatButton : Control
    {
        public bool Primary;
        public bool Danger;
        private bool _hover, _down;

        public FlatButton()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            Cursor = Cursors.Hand;
            TabStop = true;
        }

        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; _down = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { if (e.Button == MouseButtons.Left) { _down = true; Focus(); Invalidate(); } base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { _down = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }
        protected override void OnTextChanged(EventArgs e) { Invalidate(); base.OnTextChanged(e); }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }
        protected override bool IsInputKey(Keys keyData) { return keyData == Keys.Enter || keyData == Keys.Space || base.IsInputKey(keyData); }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space) { OnClick(EventArgs.Empty); e.Handled = true; }
            base.OnKeyDown(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Ui.Quality(g);
            g.Clear(Parent == null ? Color.White : BackColor);
            var r = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);
            Color fill, text, border;
            if (!Enabled) { fill = Color.FromArgb(236, 236, 236); text = Color.FromArgb(160, 160, 160); border = fill; }
            else if (Primary)
            {
                Color baseColor = Danger ? Ui.Danger : Ui.Accent;
                fill = _down ? ControlPaint.Dark(baseColor, 0.12f) : _hover ? ControlPaint.Dark(baseColor, 0.05f) : baseColor;
                if (!Danger) fill = _down ? Ui.AccentPressed : _hover ? Ui.AccentHover : Ui.Accent;
                text = Color.White; border = fill;
            }
            else
            {
                fill = _down ? Color.FromArgb(232, 232, 232) : _hover ? Color.FromArgb(243, 243, 243) : Color.White;
                text = Ui.Text; border = Ui.Border;
            }
            using (GraphicsPath p = Ui.Round(r, Ui.S(4)))
            {
                using (var b = new SolidBrush(fill)) g.FillPath(b, p);
                using (var pen = new Pen(border, 1f)) g.DrawPath(pen, p);
            }
            if (Focused && ShowFocusCues)
            {
                var fr = new RectangleF(2.5f, 2.5f, Width - 5f, Height - 5f);
                using (GraphicsPath p = Ui.Round(fr, Ui.S(3)))
                using (var pen = new Pen(Primary ? Color.FromArgb(200, 255, 255, 255) : Ui.TextSoft, 1f) { DashStyle = DashStyle.Dot })
                    g.DrawPath(pen, p);
            }
            TextRenderer.DrawText(g, Text, Font, ClientRectangle, text,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPadding);
        }
    }

    /// <summary>The "x" in the top right corner.</summary>
    internal sealed class CloseButton : Control
    {
        private bool _hover;
        public CloseButton()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        }
        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(_hover ? Ui.Danger : Color.White);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            float cx = Width / 2f, cy = Height / 2f, d = Ui.S(5);
            using (var pen = new Pen(_hover ? Color.White : Ui.TextSoft, Math.Max(1f, Ui.Scale)))
            {
                g.DrawLine(pen, cx - d, cy - d, cx + d, cy + d);
                g.DrawLine(pen, cx - d, cy + d, cx + d, cy - d);
            }
        }
    }

    /// <summary>A text box with a rounded border that turns accent-coloured when focused.</summary>
    internal sealed class PathBox : Control
    {
        public readonly TextBox Box = new TextBox();
        public PathBox()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            BackColor = Color.White;
            Box.BorderStyle = BorderStyle.None;
            Box.BackColor = Color.White;
            Box.ForeColor = Ui.Text;
            Box.GotFocus += delegate { Invalidate(); };
            Box.LostFocus += delegate { Invalidate(); };
            Controls.Add(Box);
        }
        protected override void OnFontChanged(EventArgs e) { Box.Font = Font; Place(); base.OnFontChanged(e); }
        protected override void OnResize(EventArgs e) { Place(); base.OnResize(e); }
        private void Place()
        {
            // a borderless TextBox reports a height that clips descenders, so size it from the font
            int pad = Ui.S(10);
            int h = Math.Min(Height - 4, TextRenderer.MeasureText("Ayg国", Box.Font).Height + Ui.S(2));
            Box.AutoSize = false;
            Box.SetBounds(pad, (Height - h) / 2, Math.Max(10, Width - 2 * pad), h);
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Ui.Quality(g);
            g.Clear(Color.White);
            using (GraphicsPath p = Ui.Round(new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f), Ui.S(4)))
            using (var pen = new Pen(Box.Focused ? Ui.Accent : Ui.Border, 1f))
                g.DrawPath(pen, p);
        }
    }

    /// <summary>A thin rounded progress bar.</summary>
    internal sealed class ThinProgress : Control
    {
        private float _value;
        public float Value { get { return _value; } set { _value = Math.Max(0, Math.Min(1, value)); Invalidate(); } }
        public ThinProgress()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Ui.Quality(g);
            g.Clear(Color.White);
            float radius = Height / 2f - 0.01f;
            using (GraphicsPath p = Ui.Round(new RectangleF(0, 0, Width, Height), radius))
            using (var b = new SolidBrush(Color.FromArgb(232, 232, 232))) g.FillPath(b, p);
            float w = Math.Max(Height, Width * _value);
            if (_value > 0)
                using (GraphicsPath p = Ui.Round(new RectangleF(0, 0, w, Height), radius))
                using (var b = new SolidBrush(Ui.Accent)) g.FillPath(b, p);
        }
    }

    /// <summary>
    /// One borderless window with three states: start (choose folder, Install / Uninstall),
    /// working (progress), done.
    /// </summary>
    internal sealed class SetupForm : Form
    {
        private enum Page { Start, Working, Done }
        private const int W = 620, H = 430, BarH = 72;

        private bool _uninstall;
        private Page _page = Page.Start;
        private bool _failed;
        public bool Completed;

        private readonly CloseButton _close = new CloseButton();
        private readonly Label _title = new Label(), _subtitle = new Label(), _desc = new Label();
        private readonly Label _pathCaption = new Label(), _note = new Label(), _status = new Label();
        private readonly Label _doneTitle = new Label(), _doneText = new Label(), _error = new Label();
        private readonly PathBox _path = new PathBox();
        private readonly FlatButton _browse = new FlatButton(), _primary = new FlatButton(), _secondary = new FlatButton();
        private readonly ThinProgress _progress = new ThinProgress();
        private readonly Timer _timer = new Timer();
        private Action[] _steps;
        private int _step;

        private static string T(string zh, string en) { return Program.T(zh, en); }

        public SetupForm(bool uninstall)
        {
            _uninstall = uninstall;
            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) Ui.Scale = Math.Max(1f, g.DpiX / 96f);

            Text = T("选区导出 安装程序", "Export Selection Setup");
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.White;
            ClientSize = new Size(Ui.S(W), Ui.S(H));
            DoubleBuffered = true;
            KeyPreview = true;
            try { Icon = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location); } catch { }

            _close.SetBounds(Ui.S(W - 46) - 1, 1, Ui.S(46), Ui.S(34));   // 1 px in, so the window's hairline border stays visible
            _close.Click += delegate { if (_page != Page.Working) Close(); };
            Controls.Add(_close);

            AddLabel(_title, 112, 36, 440, 40, Ui.Font(18f, FontStyle.Bold), Ui.Text);
            AddLabel(_subtitle, 114, 78, 440, 22, Ui.Font(9.5f, FontStyle.Regular), Ui.TextSoft);
            AddLabel(_desc, 40, 124, 540, 48, Ui.Font(10f, FontStyle.Regular), Color.FromArgb(60, 60, 60));
            AddLabel(_pathCaption, 40, 228, 540, 20, Ui.Font(9f, FontStyle.Regular), Ui.TextSoft);
            AddLabel(_note, 40, 294, 540, 38, Ui.Font(9f, FontStyle.Regular), Ui.TextSoft);
            AddLabel(_status, 40, 236, 540, 22, Ui.Font(10f, FontStyle.Regular), Ui.Text);
            AddLabel(_doneTitle, 92, 132, 480, 34, Ui.Font(15f, FontStyle.Bold), Ui.Text);
            AddLabel(_doneText, 40, 190, 540, 80, Ui.Font(10f, FontStyle.Regular), Color.FromArgb(60, 60, 60));
            AddLabel(_error, 40, H - BarH + 10, 250, BarH - 20, Ui.Font(9f, FontStyle.Regular), Ui.Danger);
            _error.BackColor = Ui.Bar;
            _error.TextAlign = ContentAlignment.MiddleLeft;

            _path.Font = Ui.Font(10f, FontStyle.Regular);
            _path.SetBounds(Ui.S(40), Ui.S(250), Ui.S(440), Ui.S(34));
            _path.Box.Text = Installer.SuggestedDir();
            _path.Box.TextChanged += delegate { ShowError(null); };
            Controls.Add(_path);

            SetupButton(_browse, false, 490, 250, 90, 34, Color.White);
            _browse.Text = T("浏览…", "Browse…");
            _browse.Click += delegate { Browse(); };

            SetupButton(_primary, true, W - 40 - 124, H - BarH + 18, 124, 36, Ui.Bar);
            SetupButton(_secondary, false, W - 40 - 124 - 12 - 108, H - BarH + 18, 108, 36, Ui.Bar);
            _primary.Click += delegate { OnPrimary(); };
            _secondary.Click += delegate { OnSecondary(); };

            _progress.SetBounds(Ui.S(40), Ui.S(268), Ui.S(540), Ui.S(6));
            Controls.Add(_progress);

            _timer.Interval = 170;
            _timer.Tick += delegate { RunNextStep(); };

            foreach (Control c in new Control[] { _title, _subtitle, _desc }) c.MouseDown += DragWindow;
            MouseDown += DragWindow;
            ShowPage(Page.Start);
        }

        private void AddLabel(Label l, int x, int y, int w, int h, Font font, Color color)
        {
            l.AutoSize = false;
            l.SetBounds(Ui.S(x), Ui.S(y), Ui.S(w), Ui.S(h));
            l.Font = font;
            l.ForeColor = color;
            l.BackColor = Color.White;
            l.UseMnemonic = false;
            Controls.Add(l);
        }

        private void SetupButton(FlatButton b, bool primary, int x, int y, int w, int h, Color back)
        {
            b.Primary = primary;
            b.BackColor = back;
            b.Font = Ui.Font(10f, FontStyle.Regular);
            b.SetBounds(Ui.S(x), Ui.S(y), Ui.S(w), Ui.S(h));
            Controls.Add(b);
        }

        // ------------------------------------------------------------------ pages

        private void ShowPage(Page page)
        {
            _page = page;
            bool installed = _previewOldVersion != null || (!_previewFresh && Installer.IsInstalled());
            _title.Text = T("选区导出", "Export Selection");
            _subtitle.Text = T("PowerPoint 插件", "Add-in for PowerPoint") + "  ·  v" + Installer.Version;

            bool start = page == Page.Start, working = page == Page.Working, done = page == Page.Done;
            _desc.Visible = start || working;
            _pathCaption.Visible = start;
            _path.Visible = start && !_uninstall;
            _browse.Visible = start && !_uninstall;
            _note.Visible = start;
            _status.Visible = working;
            _progress.Visible = working;
            _doneTitle.Visible = done;
            _doneText.Visible = done;
            _close.Visible = !working;
            _primary.Enabled = !working;
            _secondary.Enabled = !working;
            _primary.Danger = false;
            ShowError(null);

            if (_uninstall)
            {
                _desc.Text = T("将从这台电脑移除“选区导出”插件和它记住的设置。你的演示文稿和已经导出的文件不受影响。",
                               "This removes the Export Selection add-in and its saved settings from this computer. Your presentations and exported files are not affected.");
                _pathCaption.Text = T("安装位置：", "Installed in: ") + (Installer.InstalledDir() ?? Installer.DefaultDir);
                _note.Text = "";
                _status.Text = T("正在卸载…", "Removing…");
                _primary.Text = done ? T("完成", "Finish") : T("卸载", "Uninstall");
                _primary.Danger = !done;
                _secondary.Text = T("取消", "Cancel");
                _secondary.Visible = start;
                _doneTitle.Text = T("已卸载", "Removed");
                _doneText.Text = T("“选区导出”已经从这台电脑移除。", "Export Selection has been removed from this computer.");
            }
            else
            {
                _desc.Text = T("把幻灯片上选中的对象导出为图片或矢量图，范围是刚好包住它们的最小矩形，背景可选透明或纯色。",
                               "Export the objects you select on a slide as pictures or vector graphics, cropped to the smallest rectangle around them, with a transparent or solid background.");
                _pathCaption.Text = T("安装位置", "Install location");
                string old = _previewOldVersion ?? (installed ? Installer.InstalledVersion() : null);
                bool sameVersion = old == Installer.Version;
                if (!installed)
                    _note.Text = T("只为当前用户安装，不需要管理员权限。", "Installs for the current user only. No administrator rights needed.");
                else if (old == null)
                    _note.Text = T("这台电脑上已经装有旧版本，安装时会先把它卸载干净。", "An earlier copy is installed. It is removed completely before this one is installed.");
                else if (sameVersion)
                    _note.Text = T("已安装 " + old + "。重新安装时会先把它卸载干净。", "Version " + old + " is installed. It is removed completely before reinstalling.");
                else
                    _note.Text = T("已安装 " + old + "。升级时会先把它卸载干净，再安装 " + Installer.Version + "；你的设置会保留。",
                                   "Version " + old + " is installed. It is removed completely before " + Installer.Version + " is installed; your settings are kept.");
                _status.Text = T("正在安装…", "Installing…");
                _primary.Text = done ? T("完成", "Finish") : !installed ? T("安装", "Install") : sameVersion || old == null ? T("重新安装", "Reinstall") : T("升级", "Upgrade");
                _secondary.Text = done ? T("打开 PowerPoint", "Open PowerPoint") : T("卸载", "Uninstall");
                _secondary.Visible = done || (start && installed);
                _doneTitle.Text = T("安装完成", "Installed");
                _doneText.Text = T("打开 PowerPoint，在“视图”选项卡后面可以找到“Litchi”选项卡。\n先在幻灯片上选中对象，再点导出按钮。",
                                   "Start PowerPoint: the Litchi tab appears after the View tab.\nSelect objects on a slide, then click one of the export buttons.");
            }
            LayoutButtons();
            Invalidate();
            if (start && !_uninstall) ActiveControl = _primary;
        }

        /// <summary>Buttons are as wide as their text needs, right-aligned in the bottom bar.</summary>
        private void LayoutButtons()
        {
            int y = Ui.S(H - BarH + 18), h = Ui.S(36), right = Ui.S(W - 40);
            int pw = Math.Max(Ui.S(124), TextRenderer.MeasureText(_primary.Text, _primary.Font).Width + Ui.S(36));
            int sw = Math.Max(Ui.S(108), TextRenderer.MeasureText(_secondary.Text, _secondary.Font).Width + Ui.S(36));
            _primary.SetBounds(right - pw, y, pw, h);
            _secondary.SetBounds(right - pw - Ui.S(12) - sw, y, sw, h);
            int errorRight = (_secondary.Visible ? _secondary.Left : _primary.Left) - Ui.S(16);
            _error.SetBounds(Ui.S(40), Ui.S(H - BarH + 10), Math.Max(Ui.S(120), errorRight - Ui.S(40)), Ui.S(BarH - 20));
        }

        private void ShowError(string text)
        {
            _error.Text = text ?? "";
            _error.Visible = !string.IsNullOrEmpty(text);
        }

        // ------------------------------------------------------------------ actions

        private void Browse()
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = T("选择安装位置（文件夹里已有其他内容时，会在其中新建 PptSelectionExport 文件夹）", "Choose where to install (if the folder already contains other things, a PptSelectionExport folder is created inside it)");
                dlg.ShowNewFolderButton = true;
                try
                {
                    string cur = _path.Box.Text.Trim();
                    while (!string.IsNullOrEmpty(cur) && !Directory.Exists(cur)) cur = Path.GetDirectoryName(cur);
                    if (!string.IsNullOrEmpty(cur)) dlg.SelectedPath = cur;
                }
                catch { }
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                try { _path.Box.Text = Installer.NormalizeDir(dlg.SelectedPath); }
                catch { _path.Box.Text = dlg.SelectedPath; }
            }
        }

        private void OnPrimary()
        {
            if (_page == Page.Done) { Close(); return; }
            if (_page != Page.Start) return;

            string dir = null;
            if (!_uninstall)
            {
                try { dir = Installer.NormalizeDir(_path.Box.Text); }
                catch (Exception ex) { ShowError(ex.Message); return; }
                _path.Box.Text = dir;
            }
            if (Installer.PowerPointRunning())
            {
                ShowError(T("PowerPoint 正在运行。请先保存并关闭它，然后再点一次。", "PowerPoint is running. Save your work, close it, then click again."));
                return;
            }

            _labels = _uninstall ? Installer.UninstallLabels() : Installer.InstallLabels();
            _steps = _uninstall ? Installer.UninstallSteps() : Installer.InstallSteps(dir);
            _step = 0;
            _failed = false;
            ShowPage(Page.Working);
            _status.Text = _labels[0];
            _progress.Value = 0.06f;
            _timer.Start();
        }
        private string[] _labels;

        private void RunNextStep()
        {
            if (_step < _steps.Length)
            {
                _status.Text = _labels[Math.Min(_step, _labels.Length - 1)];
                _status.Update();
                try { _steps[_step](); }
                catch (Exception ex)
                {
                    _timer.Stop();
                    _failed = true;
                    ShowPage(Page.Start);
                    ShowError(ex.Message);
                    return;
                }
                _step++;
                _progress.Value = (float)_step / _steps.Length;
                return;
            }
            _timer.Stop();
            if (_failed) return;
            Completed = true;
            ShowPage(Page.Done);
            ActiveControl = _primary;
        }

        private void OnSecondary()
        {
            if (_page == Page.Done)
            {
                try { Process.Start("powerpnt.exe"); } catch { }
                Close();
                return;
            }
            if (_page != Page.Start) return;
            if (_uninstall)
            {
                // "Cancel": back to the install page when we came from there, otherwise leave
                if (_cameFromInstall) { _uninstall = false; _cameFromInstall = false; ShowPage(Page.Start); }
                else Close();
            }
            else { _uninstall = true; _cameFromInstall = true; ShowPage(Page.Start); }
        }
        private bool _cameFromInstall;

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape && _page != Page.Working) { Close(); e.Handled = true; }
            base.OnKeyDown(e);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_page == Page.Working) e.Cancel = true;
            base.OnFormClosing(e);
        }

        // ------------------------------------------------------------------ painting

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Ui.Quality(g);
            float s = Ui.Scale;

            // bottom bar
            var bar = new Rectangle(0, Ui.S(H - BarH), Ui.S(W), Ui.S(BarH));
            using (var b = new SolidBrush(Ui.Bar)) g.FillRectangle(b, bar);
            using (var pen = new Pen(Ui.Line, 1f)) g.DrawLine(pen, 0, bar.Top, bar.Right, bar.Top);

            DrawLogo(g, 40, 34, 56);

            if (_page == Page.Start && !_uninstall)
            {
                // what it exports, as small tags
                string[] tags = { "PDF", "SVG", "PNG", "JPG", T("透明 / 纯色背景", "Transparent / solid background") };
                float x = Ui.S(40), y = Ui.S(182);
                using (Font f = Ui.Font(8.5f, FontStyle.Regular))
                {
                    foreach (string tag in tags)
                    {
                        Size sz = TextRenderer.MeasureText(g, tag, f, Size.Empty, TextFormatFlags.NoPadding);
                        var r = new RectangleF(x, y, sz.Width + Ui.S(18), Ui.S(24));
                        using (GraphicsPath p = Ui.Round(r, Ui.S(12)))
                        using (var b = new SolidBrush(Color.FromArgb(253, 240, 235))) g.FillPath(b, p);
                        TextRenderer.DrawText(g, tag, f, Rectangle.Round(r), Ui.AccentPressed,
                            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
                        x += r.Width + Ui.S(8);
                    }
                }
            }

            if (_page == Page.Done)
            {
                // a green tick
                var c = new RectangleF(Ui.S(40), Ui.S(130), Ui.S(38), Ui.S(38));
                using (var b = new SolidBrush(Ui.Success)) g.FillEllipse(b, c);
                using (var pen = new Pen(Color.White, 2.6f * s) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                    g.DrawLines(pen, new[]
                    {
                        new PointF(c.X + c.Width * 0.28f, c.Y + c.Height * 0.52f),
                        new PointF(c.X + c.Width * 0.44f, c.Y + c.Height * 0.68f),
                        new PointF(c.X + c.Width * 0.73f, c.Y + c.Height * 0.36f)
                    });
            }

            // hairline around the window (it has no system border)
            using (var pen = new Pen(Color.FromArgb(200, 200, 200), 1f)) g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
        }

        /// <summary>The add-in's emblem: a sheet of paper with an accent label.</summary>
        private static void DrawLogo(Graphics g, int lx, int ly, int lsize)
        {
            GraphicsState state = g.Save();
            g.TranslateTransform(Ui.S(lx), Ui.S(ly));
            float k = Ui.Scale * lsize / 32f;
            g.ScaleTransform(k, k);
            using (var page = new GraphicsPath())
            {
                page.AddPolygon(new[] { new PointF(6.5f, 1.5f), new PointF(19.5f, 1.5f), new PointF(26.5f, 8.5f), new PointF(26.5f, 30.5f), new PointF(6.5f, 30.5f) });
                g.FillPath(Brushes.White, page);
                using (var pen = new Pen(Color.FromArgb(150, 150, 150), 0.8f)) g.DrawPath(pen, page);
            }
            using (var fold = new GraphicsPath())
            {
                fold.AddPolygon(new[] { new PointF(19.5f, 1.5f), new PointF(19.5f, 8.5f), new PointF(26.5f, 8.5f) });
                using (var b = new SolidBrush(Color.FromArgb(222, 222, 222))) g.FillPath(b, fold);
                using (var pen = new Pen(Color.FromArgb(150, 150, 150), 0.8f)) g.DrawPath(pen, fold);
            }
            using (var pen = new Pen(Color.FromArgb(165, 165, 165), 0.8f) { DashPattern = new[] { 2f, 1.5f } })
                g.DrawRectangle(pen, 10.5f, 6.5f, 7f, 5f);
            using (GraphicsPath band = Ui.Round(new RectangleF(1.5f, 14.5f, 25f, 12f), 2.2f))
            using (var b = new SolidBrush(Ui.Accent)) g.FillPath(b, band);
            using (var font = new Font("Segoe UI", 7.6f, FontStyle.Bold, GraphicsUnit.Pixel))
            using (var fmt = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
            {
                g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                g.DrawString("PPT", font, Brushes.White, new RectangleF(1.5f, 14.5f, 25f, 12.6f), fmt);
            }
            g.Restore(state);
        }

        // ------------------------------------------------------------------ window chrome

        [StructLayout(LayoutKind.Sequential)]
        private struct MARGINS { public int Left, Right, Top, Bottom; }
        [DllImport("dwmapi.dll")] private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
        [DllImport("dwmapi.dll")] private static extern int DwmExtendFrameIntoClientArea(IntPtr hwnd, ref MARGINS margins);
        [DllImport("user32.dll")] private static extern bool ReleaseCapture();
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr w, IntPtr l);

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.Style |= 0x00020000;        // WS_MINIMIZEBOX: taskbar click minimises
                cp.ClassStyle |= 0x00020000;   // CS_DROPSHADOW: a shadow even where DWM gives none
                return cp;
            }
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            try
            {
                int round = 2;                 // DWMWCP_ROUND, Windows 11 rounded corners
                DwmSetWindowAttribute(Handle, 33, ref round, sizeof(int));
            }
            catch { }
        }

        private void DragWindow(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            ReleaseCapture();
            SendMessage(Handle, 0x00A1 /* WM_NCLBUTTONDOWN */, new IntPtr(2) /* HTCAPTION */, IntPtr.Zero);
        }

        // ------------------------------------------------------------------ developer aid

        private static bool _previewFresh;
        private static string _previewOldVersion;

        /// <summary>
        /// Renders every page of the window to PNG files (Litchi-Setup.exe /preview folder),
        /// as it looks on a computer where the add-in is not installed yet. Used for the README pictures.
        /// </summary>
        public static void RenderPreviews(string folder)
        {
            Directory.CreateDirectory(folder);
            _previewFresh = true;
            const string samplePath = @"D:\Tools\PptSelectionExport";
            foreach (bool uninstall in new[] { false, true })
            {
                using (var f = new SetupForm(uninstall))
                {
                    f.StartPosition = FormStartPosition.Manual;
                    f.Location = new Point(-5000, -5000);
                    f.ShowInTaskbar = false;
                    f._path.Box.Text = samplePath;
                    f.Show();
                    foreach (Page p in new[] { Page.Start, Page.Working, Page.Done })
                    {
                        f.ShowPage(p);
                        if (p == Page.Working) f._progress.Value = 0.6f;
                        if (uninstall) f._pathCaption.Text = T("安装位置：", "Installed in: ") + samplePath;
                        f.ActiveControl = null;
                        Application.DoEvents();
                        using (var bmp = new Bitmap(f.Width, f.Height))
                        {
                            f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
                            bmp.Save(Path.Combine(folder, (uninstall ? "uninstall-" : "install-") + p.ToString().ToLowerInvariant() + ".png"));
                        }
                    }
                    f._page = Page.Start;
                    f.Close();
                }
            }

            // the start page as someone with an older version sees it
            _previewOldVersion = "1.3.0";
            using (var f = new SetupForm(false))
            {
                f.StartPosition = FormStartPosition.Manual;
                f.Location = new Point(-5000, -5000);
                f.ShowInTaskbar = false;
                f._path.Box.Text = samplePath;
                f.Show();
                f.ShowPage(Page.Start);
                f.ActiveControl = null;
                Application.DoEvents();
                using (var bmp = new Bitmap(f.Width, f.Height))
                {
                    f.DrawToBitmap(bmp, new Rectangle(0, 0, f.Width, f.Height));
                    bmp.Save(Path.Combine(folder, "install-upgrade.png"));
                }
                f.Close();
            }
            _previewOldVersion = null;
        }
    }
}
