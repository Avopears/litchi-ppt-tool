using System;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PptSelectionExport
{
    /// <summary>
    /// Asks for an outline weight that is not in the list ("More weights…").
    /// Same look as the colour picker: frameless, white, button bar at the bottom.
    /// </summary>
    internal sealed class WeightForm : Form
    {
        private const int W = 360, H = 214, BarH = 64, Pad = 24;

        private readonly IWin32Window _owner;
        private readonly CloseButton _close = new CloseButton();
        private readonly FieldBox _field = new FieldBox();
        private readonly FlatButton _ok = new FlatButton(), _cancel = new FlatButton();
        private Font _titleFont, _labelFont, _fieldFont, _buttonFont;
        private bool _invalid;

        /// <summary>The weight entered, in points.</summary>
        public float Weight { get; private set; }

        public WeightForm(float current, IWin32Window owner)
        {
            _owner = owner;
            Weight = current;

            Text = Strings.T("weight.title");
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            BackColor = Color.White;
            KeyPreview = true;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);

            _close.AccessibleName = Strings.T("picker.close");
            _close.Click += delegate { DialogResult = DialogResult.Cancel; };
            _ok.Primary = true;
            _ok.Text = Strings.T("picker.ok");
            _ok.Click += delegate { Accept(); };
            _cancel.Text = Strings.T("picker.cancel");
            _cancel.Click += delegate { DialogResult = DialogResult.Cancel; };
            _field.Box.MaxLength = 8;
            _field.Box.AccessibleName = Strings.T("weight.caption");
            _field.Box.Text = current.ToString("0.##", CultureInfo.InvariantCulture);
            _field.Box.TextChanged += delegate { if (_invalid) { _invalid = false; Invalidate(); } };

            Controls.AddRange(new Control[] { _close, _field, _cancel, _ok });
        }

        private void Accept()
        {
            float w;
            if (!StyleTools.TryParseWeight(_field.Box.Text, out w))
            {
                _invalid = true;
                Invalidate();
                _field.Box.Focus();
                _field.Box.SelectAll();
                return;
            }
            Weight = w;
            DialogResult = DialogResult.OK;
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= 0x00020000;   // CS_DROPSHADOW
                return cp;
            }
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            try
            {
                int round = 2;                 // DWMWCP_ROUND: Windows 11 rounded corners
                DwmSetWindowAttribute(Handle, 33, ref round, sizeof(int));
            }
            catch { }
            try
            {
                uint dpi = GetDpiForWindow(Handle);
                if (dpi >= 96) Look.Scale = dpi / 96f;
            }
            catch { }
            LayoutAll();
        }

        private void LayoutAll()
        {
            _titleFont = Look.Font(17, FontStyle.Bold);
            _labelFont = Look.Font(12.5f, FontStyle.Regular);
            _fieldFont = Look.Font(13.5f, FontStyle.Regular);
            _buttonFont = Look.Font(13.5f, FontStyle.Regular);
            Size = new Size(Look.S(W), Look.S(H));

            _close.SetBounds(Look.S(W - 46) - 1, 1, Look.S(46), Look.S(34));
            _field.Font = _fieldFont;
            _field.SetBounds(Look.S(Pad), Look.S(84), Look.S(W - 2 * Pad), Look.S(34));
            int by = H - BarH + (BarH - 36) / 2;
            _ok.SetBounds(Look.S(W - Pad - 96), Look.S(by), Look.S(96), Look.S(36));
            _cancel.SetBounds(Look.S(W - Pad - 96 - 10 - 96), Look.S(by), Look.S(96), Look.S(36));
            _ok.BackColor = _cancel.BackColor = Look.Bar;
            _ok.Font = _cancel.Font = _buttonFont;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            Rectangle area = Screen.PrimaryScreen.WorkingArea;
            Point center = new Point(area.Left + area.Width / 2, area.Top + area.Height / 2);
            RECT r;
            if (_owner != null && _owner.Handle != IntPtr.Zero && GetWindowRect(_owner.Handle, out r))
            {
                center = new Point((r.L + r.R) / 2, (r.T + r.B) / 2);
                area = Screen.FromHandle(_owner.Handle).WorkingArea;
            }
            Location = new Point(Math.Max(area.Left, Math.Min(center.X - Width / 2, area.Right - Width)),
                                 Math.Max(area.Top, Math.Min(center.Y - Height / 2, area.Bottom - Height)));
            ActiveControl = _field.Box;
            _field.Box.SelectAll();
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) { Accept(); e.Handled = e.SuppressKeyPress = true; }
            else if (e.KeyCode == Keys.Escape) { DialogResult = DialogResult.Cancel; e.Handled = e.SuppressKeyPress = true; }
            base.OnKeyDown(e);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button == MouseButtons.Left && e.Y < Look.S(52))
            {
                ReleaseCapture();
                SendMessage(Handle, 0x00A1 /* WM_NCLBUTTONDOWN */, new IntPtr(2) /* HTCAPTION */, IntPtr.Zero);
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Look.Quality(g);
            g.Clear(Color.White);
            TextRenderer.DrawText(g, Text, _titleFont, new Point(Look.S(Pad), Look.S(20)), Look.Text, TextFormatFlags.NoPadding);
            TextRenderer.DrawText(g, Strings.T("weight.caption"), _labelFont, new Point(Look.S(Pad), Look.S(62)), Look.TextSoft, TextFormatFlags.NoPadding);
            TextRenderer.DrawText(g, Strings.T("weight.hint"), _labelFont, new Point(Look.S(Pad), Look.S(124)),
                _invalid ? Look.Danger : Color.FromArgb(150, 150, 150), TextFormatFlags.NoPadding);

            var bar = new Rectangle(0, ClientSize.Height - Look.S(BarH), ClientSize.Width, Look.S(BarH));
            using (var b = new SolidBrush(Look.Bar)) g.FillRectangle(b, bar);
            using (var pen = new Pen(Look.Line, 1f)) g.DrawLine(pen, 0, bar.Top, bar.Right, bar.Top);
            using (var pen = new Pen(Color.FromArgb(200, 200, 200), 1f)) g.DrawRectangle(pen, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                foreach (Font f in new[] { _titleFont, _labelFont, _fieldFont, _buttonFont }) if (f != null) f.Dispose();
            base.Dispose(disposing);
        }

        [StructLayout(LayoutKind.Sequential)] private struct RECT { public int L, T, R, B; }
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr h, out RECT r);
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr h);
        [DllImport("user32.dll")] private static extern bool ReleaseCapture();
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr w, IntPtr l);
        [DllImport("dwmapi.dll")] private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
    }
}
