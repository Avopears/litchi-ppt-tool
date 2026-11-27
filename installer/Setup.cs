// PptSelectionExport-Setup.exe: a single-file installer / uninstaller for the add-in.
//
// It needs no administrator rights: everything goes to the current user
// (%LOCALAPPDATA%\PptSelectionExport and HKEY_CURRENT_USER). The add-in DLL travels inside
// this exe as an embedded resource. A copy of the exe is kept as Uninstall.exe so the add-in
// can be removed from Windows "Installed apps".
//
//   PptSelectionExport-Setup.exe               interactive
//   PptSelectionExport-Setup.exe /S            install silently
//   PptSelectionExport-Setup.exe /uninstall    uninstall (add /S for silent)
//
// Exit codes: 0 = done, 1 = cancelled, 2 = PowerPoint is running (silent mode), 3 = error.
using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("PptSelectionExport Setup")]
[assembly: AssemblyDescription("Installer for the PowerPoint add-in \"Export Selection\"")]
[assembly: AssemblyProduct("PptSelectionExport")]
[assembly: AssemblyVersion("1.1.0.0")]
[assembly: AssemblyFileVersion("1.1.0.0")]

namespace PptSelectionExportSetup
{
    internal static class Program
    {
        private const string ProgId = "PptSelectionExport.Connect";
        private const string Clsid = "{A0181EA2-7061-43C3-9086-B72894478235}";
        private const string DllName = "PptSelectionExport.dll";
        private const string AddinKey = @"Software\Microsoft\Office\PowerPoint\Addins\" + ProgId;
        private const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\PptSelectionExport";
        private const string SettingsKey = @"Software\PptSelectionExport";
        private const string ManagedCategory = "{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}";
        private const string HomePage = "https://github.com/Avopears/ppt-selection-export";
        private static readonly string[] ClassRoots = { @"Software\Classes", @"Software\Classes\Wow6432Node" };

        // Chinese when either the Windows display language or the regional format is Chinese
        // (many Chinese-speaking users run an English Windows with Chinese regional settings).
        internal static readonly bool Zh =
            CultureInfo.CurrentUICulture.Name.StartsWith("zh", StringComparison.OrdinalIgnoreCase) ||
            CultureInfo.CurrentCulture.Name.StartsWith("zh", StringComparison.OrdinalIgnoreCase);
        internal static string T(string zh, string en) { return Zh ? zh : en; }
        private static readonly string Title = T("选区导出 (PowerPoint 插件) 安装程序", "Export Selection for PowerPoint - Setup");

        private static string InstallDir
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PptSelectionExport"); }
        }

        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [STAThread]
        private static int Main(string[] args)
        {
            bool silent = false, uninstall = false;
            foreach (string a in args)
            {
                string s = a.TrimStart('/', '-').ToLowerInvariant();
                if (s == "s" || s == "silent" || s == "quiet") silent = true;
                else if (s == "uninstall" || s == "u") uninstall = true;
            }

            try
            {
                if (!silent)
                {
                    try { SetProcessDPIAware(); } catch { }
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                }

                if (silent) return uninstall ? Uninstall(true) : Install(true);

                if (uninstall)
                {
                    if (MessageBox.Show(T("要卸载 PowerPoint 插件“选区导出”吗？", "Remove the PowerPoint add-in \"Export Selection\"?"),
                            Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return 1;
                    return Uninstall(false);
                }

                switch (ChoiceDialog.Ask(IsInstalled()))
                {
                    case Choice.Install: return Install(false);
                    case Choice.Uninstall: return Uninstall(false);
                    default: return 1;
                }
            }
            catch (Exception ex)
            {
                if (!silent)
                    MessageBox.Show(T("操作失败：\n", "Setup failed:\n") + ex.Message, Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 3;
            }
        }

        private static bool IsInstalled()
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(AddinKey)) return k != null;
        }

        /// <summary>The add-in DLL cannot be replaced or removed while PowerPoint has it loaded.</summary>
        private static bool WaitForPowerPointClosed(bool silent)
        {
            while (Process.GetProcessesByName("POWERPNT").Length > 0)
            {
                if (silent) return false;
                DialogResult r = MessageBox.Show(
                    T("PowerPoint 正在运行。\n请先保存文档并关闭所有 PowerPoint 窗口，然后点“重试”。",
                      "PowerPoint is running.\nSave your work, close every PowerPoint window, then choose Retry."),
                    Title, MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning);
                if (r != DialogResult.Retry) return false;
            }
            return true;
        }

        // ------------------------------------------------------------------ install

        private static int Install(bool silent)
        {
            if (!WaitForPowerPointClosed(silent)) return silent ? 2 : 1;

            string dir = InstallDir;
            Directory.CreateDirectory(dir);
            string dll = Path.Combine(dir, DllName);
            using (Stream src = typeof(Program).Assembly.GetManifestResourceStream(DllName))
            {
                if (src == null) throw new InvalidOperationException("The installer is damaged: the add-in is missing from it.");
                using (var dst = new FileStream(dll, FileMode.Create, FileAccess.Write)) src.CopyTo(dst);
            }

            // keep a copy of this program as the uninstaller
            string self = Assembly.GetExecutingAssembly().Location;
            string uninstaller = Path.Combine(dir, "Uninstall.exe");
            if (!string.Equals(Path.GetFullPath(self), Path.GetFullPath(uninstaller), StringComparison.OrdinalIgnoreCase))
            {
                File.Copy(self, uninstaller, true);
                RemoveInternetMark(uninstaller);
            }

            AssemblyName name = AssemblyName.GetAssemblyName(dll);
            string codeBase = new Uri(dll).AbsoluteUri;

            // COM registration of the managed class (what "regasm /codebase" would write), per user,
            // for both the 64-bit and the 32-bit view so either edition of Office finds it.
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

            // tell PowerPoint to load it at startup
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(AddinKey))
            {
                key.SetValue("FriendlyName", "选区导出 Export Selection (PDF / SVG / PNG / JPG)");
                key.SetValue("Description", "把选中的对象按最小外接矩形导出为 PDF、SVG、PNG 或 JPG，可设置背景颜色或透明。 Export the selected objects, cropped to their bounding box, as PDF, SVG, PNG or JPG.");
                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
            }

            // entry in Windows "Installed apps"
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

            if (!silent)
                MessageBox.Show(
                    T("安装完成。\n\n打开 PowerPoint，在“视图”选项卡后面可以找到“选区导出”。",
                      "Installed.\n\nStart PowerPoint: the \"Export Selection\" tab appears after the View tab."),
                    Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return 0;
        }

        private static void SetDefault(string subKey, string value)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(subKey)) key.SetValue("", value);
        }

        /// <summary>A downloaded exe carries a "from the Internet" mark; drop it from our own copy.</summary>
        private static void RemoveInternetMark(string file)
        {
            try { DeleteFile(file + ":Zone.Identifier"); } catch { }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool DeleteFile(string name);

        // ------------------------------------------------------------------ uninstall

        private static int Uninstall(bool silent)
        {
            if (!WaitForPowerPointClosed(silent)) return silent ? 2 : 1;

            Registry.CurrentUser.DeleteSubKeyTree(AddinKey, false);
            Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\" + ProgId, false);
            foreach (string root in ClassRoots) Registry.CurrentUser.DeleteSubKeyTree(root + @"\CLSID\" + Clsid, false);
            Registry.CurrentUser.DeleteSubKeyTree(SettingsKey, false);
            Registry.CurrentUser.DeleteSubKeyTree(UninstallKey, false);

            string dir = InstallDir;
            try { File.Delete(Path.Combine(dir, DllName)); } catch { }

            try { Directory.Delete(dir, true); } catch { }

            // This program may be running from inside that folder (as Uninstall.exe) and cannot delete
            // itself, so hand the folder to a short-lived command prompt that waits for us to exit.
            if (Directory.Exists(dir))
            {
                try
                {
                    var psi = new ProcessStartInfo("cmd.exe", "/c ping 127.0.0.1 -n 3 >nul & rmdir /s /q \"" + dir + "\"");
                    psi.CreateNoWindow = true;
                    psi.UseShellExecute = false;
                    psi.WorkingDirectory = Path.GetTempPath();
                    Process.Start(psi);
                }
                catch { }
            }

            if (!silent)
                MessageBox.Show(T("已卸载。", "The add-in has been removed."), Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return 0;
        }
    }

    internal enum Choice { Cancel, Install, Uninstall }

    /// <summary>The one window of the installer: what it is, and Install / Uninstall / Cancel.</summary>
    internal sealed class ChoiceDialog : Form
    {
        private Choice _choice = Choice.Cancel;

        private static string T(string zh, string en) { return Program.T(zh, en); }

        public static Choice Ask(bool installed)
        {
            using (var d = new ChoiceDialog(installed))
            {
                d.ShowDialog();
                return d._choice;
            }
        }

        private ChoiceDialog(bool installed)
        {
            Text = T("选区导出 (PowerPoint 插件) 安装程序", "Export Selection for PowerPoint - Setup");
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(16);
            try { Icon = Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location); } catch { }

            var layout = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, Dock = DockStyle.Fill };

            var heading = new Label
            {
                AutoSize = true,
                Font = new Font(Font.FontFamily, Font.Size * 1.35f, FontStyle.Bold),
                Text = T("选区导出", "Export Selection"),
                Margin = new Padding(3, 0, 3, 8)
            };
            var body = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(TextRenderer.MeasureText(new string('M', 38), Font).Width, 0),
                Margin = new Padding(3, 0, 3, 14),
                Text = T(
                    "PowerPoint 插件：把选中的对象导出为 PDF、SVG、PNG 或 JPG，范围是刚好包住它们的最小矩形，背景可选透明或纯色。\n\n" +
                    "只安装到当前用户，不需要管理员权限。" +
                    (installed ? "\n\n这台电脑上已经安装过，可以重新安装（更新）或卸载。" : ""),
                    "A PowerPoint add-in that exports the objects you selected as PDF, SVG, PNG or JPG, cropped to the smallest rectangle around them, with a transparent or solid background.\n\n" +
                    "It installs for the current user only; no administrator rights are needed." +
                    (installed ? "\n\nIt is already installed on this computer. You can install it again (update) or remove it." : ""))
            };

            var buttons = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill, Margin = new Padding(0) };
            Button cancel = MakeButton(T("取消", "Cancel"), Choice.Cancel);
            Button install = MakeButton(installed ? T("重新安装", "Reinstall") : T("安装", "Install"), Choice.Install);
            buttons.Controls.Add(cancel);
            if (installed) buttons.Controls.Add(MakeButton(T("卸载", "Uninstall"), Choice.Uninstall));
            buttons.Controls.Add(install);

            layout.Controls.Add(heading);
            layout.Controls.Add(body);
            layout.Controls.Add(buttons);
            Controls.Add(layout);
            AcceptButton = install;
            CancelButton = cancel;
        }

        private Button MakeButton(string text, Choice choice)
        {
            var b = new Button { Text = text, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(12, 3, 12, 3), Margin = new Padding(8, 0, 0, 0) };
            b.Click += delegate { _choice = choice; Close(); };
            return b;
        }
    }
}
