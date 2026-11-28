using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PptSelectionExport
{
    /// <summary>
    /// Asks for a slide size in centimetres and tells, while the user types, by how much the
    /// content will be scaled to fit. Same look as the add-in's other windows.
    /// </summary>
    internal sealed class SlideSizeForm : Form
    {
        private const int W = 480, H = 392, BarH = 64, Pad = 24;
        private const int FieldTop = 110, PresetTop = 184, LockTop = 232, HintTop = 272;

        // width x height in centimetres; the first two are PowerPoint's own (13.333 in x 7.5 in, 10 in x 7.5 in)
        private static readonly double[][] Presets =
        {
            new[] { 13.0 + 1.0 / 3, 7.5 }, new[] { 10.0, 7.5 }, new[] { 29.7 / 2.54, 21.0 / 2.54 }      // in inches
        };
        private static readonly string[] PresetKeys = { "size.preset.wide", "size.preset.standard", "size.preset.a4" };

        private readonly IWin32Window _owner;
        private readonly Settings _settings;
        private readonly float _oldWidth, _oldHeight;
        private readonly CloseButton _close = new CloseButton();
        private readonly FieldBox _width = new FieldBox(), _height = new FieldBox();
        private readonly FlatButton _swap = new FlatButton(), _ok = new FlatButton(), _cancel = new FlatButton();
        private readonly FlatButton[] _presets = new FlatButton[Presets.Length];
        private readonly CheckLine _lock = new CheckLine();
        private Font _titleFont, _labelFont, _fieldFont, _buttonFont, _smallFont;

        /// <summary>What the fields show for values put there by the program, so that 33.87 stays the exact 33.8667 it stands for.</summary>
        private readonly Dictionary<string, float> _exact = new Dictionary<string, float>();
        private bool _filling;
        private double _ratio;          // width / height that "lock the proportions" keeps

        /// <summary>The size chosen, in points.</summary>
        public float WidthPoints { get; private set; }
        public float HeightPoints { get; private set; }

        public SlideSizeForm(float width, float height, Settings settings, IWin32Window owner)
        {
            _owner = owner;
            _settings = settings;
            _oldWidth = WidthPoints = width;
            _oldHeight = HeightPoints = height;
            _ratio = width / (double)height;

            Text = Strings.T("size.title");
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            BackColor = Color.White;
            KeyPreview = true;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);

            _close.AccessibleName = Strings.T("picker.close");
            _close.Click += delegate { DialogResult = DialogResult.Cancel; };
            _width.Box.MaxLength = _height.Box.MaxLength = 10;
            _width.Box.AccessibleName = Strings.T("size.width");
            _height.Box.AccessibleName = Strings.T("size.height");
            _width.Box.TextChanged += delegate { Typed(true); };
            _height.Box.TextChanged += delegate { Typed(false); };
            _swap.Text = Strings.T("size.swap");
            _swap.BackColor = Color.White;
            _swap.Click += delegate { Fill(HeightNow(), WidthNow()); };
            for (int i = 0; i < _presets.Length; i++)
            {
                int index = i;
                _presets[i] = new FlatButton { Text = Strings.T(PresetKeys[i]), BackColor = Color.White };
                _presets[i].Click += delegate { Fill((float)(Presets[index][0] * 72), (float)(Presets[index][1] * 72)); };
            }
            _lock.Text = Strings.T("size.lock");
            _lock.Checked = settings.SizeLockRatio;
            _lock.CheckedChanged += delegate { if (Valid()) _ratio = WidthNow() / (double)HeightNow(); };
            _ok.Primary = true;
            _ok.Text = Strings.T("picker.ok");
            _ok.Click += delegate { Accept(); };
            _cancel.Text = Strings.T("picker.cancel");
            _cancel.Click += delegate { DialogResult = DialogResult.Cancel; };

            Controls.AddRange(new Control[] { _close, _width, _height, _swap });
            Controls.AddRange(_presets);
            Controls.AddRange(new Control[] { _lock, _cancel, _ok });
            Fill(width, height);
        }

        // ------------------------------------------------------------ values

        /// <summary>"16", "16.5", "16,5", "16 cm" or "16 厘米": points, or NaN when it is not a size PowerPoint takes.</summary>
        private float Parse(string text)
        {
            float exact;
            string t = (text ?? "").Trim();
            if (_exact.TryGetValue(t, out exact)) return exact;
            t = t.ToLowerInvariant().Replace("厘米", "").Replace("cm", "").Replace(',', '.').Trim();
            double cm;
            if (!double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out cm) || double.IsNaN(cm)) return float.NaN;
            if (cm < SlideResizer.MinCm - 0.005 || cm > SlideResizer.MaxCm + 0.005) return float.NaN;
            return SlideResizer.ToPoints(Math.Max(SlideResizer.MinCm, Math.Min(SlideResizer.MaxCm, cm)));
        }

        private float WidthNow() { return Parse(_width.Box.Text); }
        private float HeightNow() { return Parse(_height.Box.Text); }
        private bool Valid() { return !float.IsNaN(WidthNow()) && !float.IsNaN(HeightNow()); }

        private static string Cm(float points)
        {
            return SlideResizer.ToCm(points).ToString("0.##", CultureInfo.InvariantCulture);
        }

        private void Put(FieldBox field, float points)
        {
            string text = Cm(points);
            _exact[text] = points;
            field.Box.Text = text;
        }

        /// <summary>Puts a size into both fields.</summary>
        private void Fill(float width, float height)
        {
            if (float.IsNaN(width) || float.IsNaN(height)) return;
            _filling = true;
            Put(_width, width);
            Put(_height, height);
            _filling = false;
            _ratio = width / (double)height;
            Changed();
        }

        private void Typed(bool inWidth)
        {
            if (_filling) return;
            if (_lock.Checked)
            {
                // the other field follows, as long as what was typed is a size
                float typed = inWidth ? WidthNow() : HeightNow();
                if (!float.IsNaN(typed))
                {
                    float other = (float)(inWidth ? typed / _ratio : typed * _ratio);
                    float min = SlideResizer.ToPoints(SlideResizer.MinCm), max = SlideResizer.ToPoints(SlideResizer.MaxCm);
                    _filling = true;
                    if (other >= min - 0.01f && other <= max + 0.01f) Put(inWidth ? _height : _width, other);
                    else (inWidth ? _height : _width).Box.Text = Cm(other);      // out of range: shown, and refused below
                    _filling = false;
                }
            }
            else if (Valid()) _ratio = WidthNow() / (double)HeightNow();
            Changed();
        }

        private void Changed()
        {
            _ok.Enabled = Valid();
            Invalidate();
        }

        private void Accept()
        {
            if (!Valid()) return;
            WidthPoints = WidthNow();
            HeightPoints = HeightNow();
            _settings.SizeLockRatio = _lock.Checked;
            _settings.Save();
            DialogResult = DialogResult.OK;
        }

        // ------------------------------------------------------------ window

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
            _smallFont = Look.Font(12, FontStyle.Regular);
            Size = new Size(Look.S(W), Look.S(H));

            _close.SetBounds(Look.S(W - 46) - 1, 1, Look.S(46), Look.S(34));
            const int field = 150, cross = 32, swap = 72;
            _width.Font = _height.Font = _fieldFont;
            _width.SetBounds(Look.S(Pad), Look.S(FieldTop), Look.S(field), Look.S(34));
            _height.SetBounds(Look.S(Pad + field + cross), Look.S(FieldTop), Look.S(field), Look.S(34));
            _swap.Font = _buttonFont;
            _swap.SetBounds(Look.S(W - Pad - swap), Look.S(FieldTop), Look.S(swap), Look.S(34));

            const int gap = 8;
            double each = (W - 2 * Pad - gap * (_presets.Length - 1)) / (double)_presets.Length;
            for (int i = 0; i < _presets.Length; i++)
            {
                _presets[i].Font = _buttonFont;
                _presets[i].SetBounds(Look.S(Pad + i * (each + gap)), Look.S(PresetTop), Look.S(each), Look.S(32));
            }
            _lock.Font = _fieldFont;
            _lock.SetBounds(Look.S(Pad), Look.S(LockTop), Look.S(W - 2 * Pad), Look.S(26));

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
            ActiveControl = _width.Box;
            _width.Box.SelectAll();
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

        private const TextFormatFlags Plain = TextFormatFlags.NoPadding | TextFormatFlags.SingleLine;

        private void Draw(Graphics g, string text, Font font, int x, int y, Color color)
        {
            TextRenderer.DrawText(g, text, font, new Point(Look.S(x), Look.S(y)), color, Plain);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Look.Quality(g);
            g.Clear(Color.White);
            Draw(g, Text, _titleFont, Pad, 20, Look.Text);
            Draw(g, Strings.T("size.current", Cm(_oldWidth), Cm(_oldHeight)), _labelFont, Pad, 58, Look.TextSoft);

            Draw(g, Strings.T("size.width"), _labelFont, Pad, FieldTop - 22, Look.TextSoft);
            Draw(g, Strings.T("size.height"), _labelFont, Pad + 150 + 32, FieldTop - 22, Look.TextSoft);
            TextRenderer.DrawText(g, "×", _fieldFont, new Rectangle(Look.S(Pad + 150), Look.S(FieldTop), Look.S(32), Look.S(34)), Look.TextSoft,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | Plain);
            Draw(g, Strings.T("size.presets"), _labelFont, Pad, PresetTop - 22, Look.TextSoft);

            // ---- what will happen to the content
            string hint;
            Color color = Look.Text;
            if (!Valid()) { hint = Strings.T("size.hint.invalid"); color = Look.Danger; }
            else
            {
                float w = WidthNow(), h = HeightNow();
                double scale = SlideResizer.FitScale(_oldWidth, _oldHeight, w, h);
                if (Math.Abs(w - _oldWidth) < 0.005f && Math.Abs(h - _oldHeight) < 0.005f) { hint = Strings.T("size.hint.nochange"); color = Look.TextSoft; }
                else if (Math.Abs(scale - 1) < 0.0005) hint = Strings.T("size.hint.same");
                else hint = Strings.T("size.hint.scale", (scale * 100).ToString("0.#", CultureInfo.InvariantCulture));
            }
            Draw(g, hint, _fieldFont, Pad, HintTop, color);
            Draw(g, Strings.T("size.note"), _smallFont, Pad, HintTop + 26, Color.FromArgb(150, 150, 150));

            var bar = new Rectangle(0, ClientSize.Height - Look.S(BarH), ClientSize.Width, Look.S(BarH));
            using (var b = new SolidBrush(Look.Bar)) g.FillRectangle(b, bar);
            using (var pen = new Pen(Look.Line, 1f)) g.DrawLine(pen, 0, bar.Top, bar.Right, bar.Top);
            using (var pen = new Pen(Color.FromArgb(200, 200, 200), 1f)) g.DrawRectangle(pen, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                foreach (Font f in new[] { _titleFont, _labelFont, _fieldFont, _buttonFont, _smallFont }) if (f != null) f.Dispose();
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
