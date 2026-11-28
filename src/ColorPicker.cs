using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PptSelectionExport
{
    /// <summary>Look of the add-in's own windows: the same as the installer's.</summary>
    internal static class Look
    {
        public static float Scale = 1f;
        public static int S(double v) { return (int)Math.Round(v * Scale); }

        public static readonly Color Accent = Color.FromArgb(196, 62, 28);          // PowerPoint's orange-red
        public static readonly Color AccentHover = Color.FromArgb(172, 52, 22);
        public static readonly Color AccentPressed = Color.FromArgb(146, 44, 18);
        public static readonly Color Text = Color.FromArgb(27, 27, 27);
        public static readonly Color TextSoft = Color.FromArgb(112, 112, 112);
        public static readonly Color Border = Color.FromArgb(208, 208, 208);
        public static readonly Color Line = Color.FromArgb(229, 229, 229);
        public static readonly Color Bar = Color.FromArgb(246, 246, 246);
        public static readonly Color Danger = Color.FromArgb(196, 43, 28);

        /// <summary>The system UI font (Microsoft YaHei UI on Chinese Windows), sized in logical pixels.</summary>
        public static Font Font(float px, FontStyle style)
        {
            return new Font(SystemFonts.MessageBoxFont.FontFamily, px * Scale, style, GraphicsUnit.Pixel);
        }

        public static GraphicsPath Round(RectangleF r, float radius)
        {
            float d = Math.Min(radius * 2, Math.Min(r.Width, r.Height));
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

    /// <summary>A flat, rounded button: filled with the accent colour (primary) or outlined.</summary>
    internal sealed class FlatButton : Control
    {
        public bool Primary;
        private bool _hover, _down;

        public FlatButton()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            Cursor = Cursors.Hand;
            TabStop = true;
        }

        public void PerformClick() { OnClick(EventArgs.Empty); }

        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; _down = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { if (e.Button == MouseButtons.Left) { _down = true; Focus(); Invalidate(); } base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { _down = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnTextChanged(EventArgs e) { Invalidate(); base.OnTextChanged(e); }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }
        protected override bool IsInputKey(Keys keyData) { return keyData == Keys.Space || base.IsInputKey(keyData); }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Space) { OnClick(EventArgs.Empty); e.Handled = true; }
            base.OnKeyDown(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Look.Quality(g);
            g.Clear(BackColor);
            var r = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);
            Color fill, text, border;
            if (Primary)
            {
                fill = _down ? Look.AccentPressed : _hover ? Look.AccentHover : Look.Accent;
                text = Color.White; border = fill;
            }
            else
            {
                fill = _down ? Color.FromArgb(232, 232, 232) : _hover ? Color.FromArgb(243, 243, 243) : Color.White;
                text = Look.Text; border = Look.Border;
            }
            using (GraphicsPath p = Look.Round(r, Look.S(4)))
            {
                using (var b = new SolidBrush(fill)) g.FillPath(b, p);
                using (var pen = new Pen(border, 1f)) g.DrawPath(pen, p);
            }
            if (Focused && ShowFocusCues)
            {
                using (GraphicsPath p = Look.Round(new RectangleF(2.5f, 2.5f, Width - 5f, Height - 5f), Look.S(3)))
                using (var pen = new Pen(Primary ? Color.FromArgb(200, 255, 255, 255) : Look.TextSoft, 1f) { DashStyle = DashStyle.Dot })
                    g.DrawPath(pen, p);
            }
            TextRenderer.DrawText(g, Text, Font, ClientRectangle, text,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPadding);
        }

        protected override AccessibleObject CreateAccessibilityInstance() { return new Accessible(this); }

        /// <summary>So that screen readers (and UI tests) see a push button they can press.</summary>
        private sealed class Accessible : ControlAccessibleObject
        {
            public Accessible(FlatButton owner) : base(owner) { }
            public override AccessibleRole Role { get { return AccessibleRole.PushButton; } }
            public override string Name { get { return Owner.Text; } set { } }
            public override string DefaultAction { get { return "Press"; } }
            public override void DoDefaultAction() { ((FlatButton)Owner).PerformClick(); }
        }
    }

    /// <summary>The "x" in the top right corner.</summary>
    internal sealed class CloseButton : Control
    {
        private bool _hover;
        public CloseButton()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            AccessibleRole = AccessibleRole.PushButton;
        }
        protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hover = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(_hover ? Look.Danger : Color.White);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            float cx = Width / 2f, cy = Height / 2f, d = Look.S(5);
            using (var pen = new Pen(_hover ? Color.White : Look.TextSoft, Math.Max(1f, Look.Scale)))
            {
                g.DrawLine(pen, cx - d, cy - d, cx + d, cy + d);
                g.DrawLine(pen, cx - d, cy + d, cx + d, cy - d);
            }
        }
    }

    /// <summary>A text box with a rounded border that turns accent-coloured when focused, with an optional prefix such as "#".</summary>
    internal sealed class FieldBox : Control
    {
        public readonly TextBox Box = new TextBox();
        private string _prefix = "";

        public FieldBox()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            BackColor = Color.White;
            Cursor = Cursors.IBeam;
            Box.BorderStyle = BorderStyle.None;
            Box.BackColor = Color.White;
            Box.ForeColor = Look.Text;
            Box.GotFocus += delegate { Invalidate(); };
            Box.LostFocus += delegate { Invalidate(); };
            Controls.Add(Box);
        }

        public string Prefix { get { return _prefix; } set { _prefix = value ?? ""; Place(); Invalidate(); } }

        protected override void OnFontChanged(EventArgs e) { Box.Font = Font; Place(); base.OnFontChanged(e); }
        protected override void OnResize(EventArgs e) { Place(); base.OnResize(e); }
        protected override void OnMouseDown(MouseEventArgs e) { Box.Focus(); Box.SelectAll(); base.OnMouseDown(e); }

        private int PrefixWidth
        {
            get { return _prefix.Length == 0 ? 0 : TextRenderer.MeasureText(_prefix, Font, Size.Empty, TextFormatFlags.NoPadding).Width + Look.S(3); }
        }

        private void Place()
        {
            // a borderless TextBox reports a height that clips descenders, so size it from the font
            int pad = Look.S(9);
            int h = Math.Min(Height - 4, TextRenderer.MeasureText("Ayg国", Box.Font).Height + Look.S(1));
            Box.AutoSize = false;
            int x = pad + PrefixWidth;
            Box.SetBounds(x, (Height - h) / 2, Math.Max(10, Width - x - pad), h);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Look.Quality(g);
            g.Clear(Parent != null ? Parent.BackColor : Color.White);
            using (GraphicsPath p = Look.Round(new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f), Look.S(4)))
            {
                using (var b = new SolidBrush(Color.White)) g.FillPath(b, p);
                using (var pen = new Pen(Box.Focused ? Look.Accent : Look.Border, Box.Focused ? Math.Max(1f, Look.Scale * 1.5f) : 1f)) g.DrawPath(pen, p);
            }
            if (_prefix.Length > 0)
                TextRenderer.DrawText(g, _prefix, Font, new Rectangle(Look.S(9), 0, PrefixWidth, Height), Look.TextSoft,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.NoPadding);
        }
    }

    /// <summary>
    /// A colour picker in the style of current design tools: a saturation / brightness square,
    /// a hue slider, hex and RGB fields, a palette and the colours used recently.
    /// </summary>
    internal sealed class ColorPickerForm : Form
    {
        private const int W = 560, H = 556, BarH = 64, Pad = 24;
        private const int SwatchSize = 28, SwatchColumns = 12;

        /// <summary>Two rows: plain colours, then soft tints that suit figure backgrounds.</summary>
        internal static readonly string[] Palette =
        {
            "FFFFFF", "BFBFBF", "7F7F7F", "000000", "EF4444", "F97316", "EAB308", "22C55E", "14B8A6", "3B82F6", "8B5CF6", "EC4899",
            "F8F9FA", "F2F2F2", "FFF8E1", "E8F1FB", "FEE2E2", "FFEDD5", "FEF9C3", "DCFCE7", "CCFBF1", "DBEAFE", "EDE9FE", "FCE7F3",
        };

        private readonly Color _original;
        private readonly List<Color> _recent;
        private readonly IWin32Window _owner;
        private Color _color;
        private float _h, _s, _v;            // hue 0..360, saturation and brightness 0..1
        private bool _updating;

        private readonly CloseButton _close = new CloseButton();
        private readonly FieldBox _hex = new FieldBox(), _r = new FieldBox(), _g = new FieldBox(), _b = new FieldBox();
        private readonly FlatButton _ok = new FlatButton(), _cancel = new FlatButton();

        private Rectangle _svRect, _hueRect, _previewRect;
        private readonly List<KeyValuePair<Rectangle, Color>> _swatches = new List<KeyValuePair<Rectangle, Color>>();
        private int _paletteLabelY, _recentLabelY, _recentRowY, _hexLabelY, _rgbLabelY;
        private Font _titleFont, _labelFont, _fieldFont, _buttonFont;
        private Bitmap _svBitmap;
        private float _svBitmapHue = -1;
        private enum Drag { None, Sv, Hue }
        private Drag _drag;
        private int _hoverSwatch = -1;

        public Color SelectedColor { get { return _color; } }

        public ColorPickerForm(Color current, IList<Color> recent, IWin32Window owner) : this(current, recent, owner, null) { }

        /// <param name="title">Heading of the window; the background-colour heading when null.</param>
        public ColorPickerForm(Color current, IList<Color> recent, IWin32Window owner, string title)
        {
            _original = Color.FromArgb(255, current);
            _recent = new List<Color>(recent ?? new Color[0]);
            _owner = owner;
            _color = _original;
            ToHsv(_color, out _h, out _s, out _v);

            Text = title ?? Strings.T("picker.title");
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
            _ok.Click += delegate { DialogResult = DialogResult.OK; };
            _cancel.Text = Strings.T("picker.cancel");
            _cancel.Click += delegate { DialogResult = DialogResult.Cancel; };
            _hex.Prefix = "#";
            _hex.Box.MaxLength = 7;
            _hex.Box.AccessibleName = "Hex";
            _r.Box.AccessibleName = "R"; _g.Box.AccessibleName = "G"; _b.Box.AccessibleName = "B";
            foreach (FieldBox f in new[] { _r, _g, _b }) { f.Box.MaxLength = 3; f.Box.TextAlign = HorizontalAlignment.Center; }

            _hex.Box.KeyPress += (s, e) => { if (!char.IsControl(e.KeyChar) && !Uri.IsHexDigit(e.KeyChar) && e.KeyChar != '#') e.Handled = true; };
            _hex.Box.TextChanged += delegate { OnHexTyped(); };
            foreach (FieldBox f in new[] { _r, _g, _b })
            {
                FieldBox field = f;
                field.Box.KeyPress += (s, e) => { if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar)) e.Handled = true; };
                field.Box.TextChanged += delegate { OnRgbTyped(field); };
                field.Box.KeyDown += (s, e) => StepChannel(field, e);
            }
            foreach (FieldBox f in new[] { _hex, _r, _g, _b }) f.Box.Leave += delegate { ShowValues(null); };

            Controls.AddRange(new Control[] { _close, _hex, _r, _g, _b, _cancel, _ok });
        }

        // ------------------------------------------------------------------ window

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
            Look.Scale = WindowScale(Handle);
            LayoutAll();
        }

        /// <summary>The scaling this window is drawn at (its monitor's, when PowerPoint is per-monitor aware).</summary>
        private static float WindowScale(IntPtr hwnd)
        {
            try
            {
                uint dpi = GetDpiForWindow(hwnd);
                if (dpi >= 96) return dpi / 96f;
            }
            catch { }   // before Windows 10 1607
            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) return Math.Max(1f, g.DpiX / 96f);
        }

        private void LayoutAll()
        {
            _titleFont = Look.Font(17, FontStyle.Bold);
            _labelFont = Look.Font(12.5f, FontStyle.Regular);
            _fieldFont = Look.Font(13.5f, FontStyle.Regular);
            _buttonFont = Look.Font(13.5f, FontStyle.Regular);
            // the window has no frame, so its size is its client size (setting ClientSize once the
            // handle exists adds room for a caption it does not have)
            Size = new Size(Look.S(W), Look.S(H));

            _close.SetBounds(Look.S(W - 46) - 1, 1, Look.S(46), Look.S(34));
            _svRect = R(Pad, 64, 300, 196);
            _hueRect = R(Pad, 278, 300, 12);

            const int colX = 348, colW = W - Pad - colX;
            _previewRect = R(colX, 64, colW, 52);
            _hexLabelY = 146;
            Place(_hex, colX, 166, colW, 34);
            _rgbLabelY = 212;
            int fw = (colW - 2 * 10) / 3;
            Place(_r, colX, 232, fw, 34);
            Place(_g, colX + fw + 10, 232, fw, 34);
            Place(_b, colX + 2 * (fw + 10), 232, fw, 34);

            _paletteLabelY = 314;
            _swatches.Clear();
            int gap = (W - 2 * Pad - SwatchColumns * SwatchSize) / (SwatchColumns - 1);
            for (int i = 0; i < Palette.Length; i++)
            {
                Color c;
                Settings.TryParseHex(Palette[i], out c);
                int col = i % SwatchColumns, row = i / SwatchColumns;
                _swatches.Add(new KeyValuePair<Rectangle, Color>(R(Pad + col * (SwatchSize + gap), 338 + row * (SwatchSize + 10), SwatchSize, SwatchSize), c));
            }
            _recentLabelY = 422;
            _recentRowY = 446;
            for (int i = 0; i < _recent.Count && i < SwatchColumns; i++)
                _swatches.Add(new KeyValuePair<Rectangle, Color>(R(Pad + i * (SwatchSize + gap), _recentRowY, SwatchSize, SwatchSize), _recent[i]));

            int by = H - BarH + (BarH - 36) / 2;
            Place(_ok, W - Pad - 104, by, 104, 36);
            Place(_cancel, W - Pad - 104 - 10 - 104, by, 104, 36);
            _ok.BackColor = _cancel.BackColor = Look.Bar;
            _ok.Font = _cancel.Font = _buttonFont;
            foreach (FieldBox f in new[] { _hex, _r, _g, _b }) f.Font = _fieldFont;
            ShowValues(null);
        }

        private static Rectangle R(int x, int y, int w, int h) { return new Rectangle(Look.S(x), Look.S(y), Look.S(w), Look.S(h)); }
        private static void Place(Control c, int x, int y, int w, int h) { c.SetBounds(Look.S(x), Look.S(y), Look.S(w), Look.S(h)); }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // centre on PowerPoint's window, inside the screen it is on
            Rectangle area = Screen.PrimaryScreen.WorkingArea;
            Point center = new Point(area.Left + area.Width / 2, area.Top + area.Height / 2);
            RECT r;
            if (_owner != null && _owner.Handle != IntPtr.Zero && GetWindowRect(_owner.Handle, out r))
            {
                center = new Point((r.L + r.R) / 2, (r.T + r.B) / 2);
                area = Screen.FromHandle(_owner.Handle).WorkingArea;
            }
            int x = Math.Max(area.Left, Math.Min(center.X - Width / 2, area.Right - Width));
            int y = Math.Max(area.Top, Math.Min(center.Y - Height / 2, area.Bottom - Height));
            Location = new Point(x, y);
            ActiveControl = _ok;
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) { DialogResult = DialogResult.OK; e.Handled = e.SuppressKeyPress = true; }
            else if (e.KeyCode == Keys.Escape) { DialogResult = DialogResult.Cancel; e.Handled = e.SuppressKeyPress = true; }
            base.OnKeyDown(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_svBitmap != null) _svBitmap.Dispose();
                foreach (Font f in new[] { _titleFont, _labelFont, _fieldFont, _buttonFont }) if (f != null) f.Dispose();
            }
            base.Dispose(disposing);
        }

        // ------------------------------------------------------------------ colour state

        private void SetColor(Color c, FieldBox typedIn)
        {
            _color = Color.FromArgb(255, c);
            float h, s, v;
            ToHsv(_color, out h, out s, out v);
            if (v > 0 && s > 0) _h = h;          // grey and black have no hue: keep the slider where it is
            if (v > 0) _s = s;
            _v = v;
            ShowValues(typedIn);
            Invalidate();
        }

        private void SetHsv(float h, float s, float v)
        {
            _h = h; _s = s; _v = v;
            _color = FromHsv(h, s, v);
            ShowValues(null);
            Invalidate();
        }

        /// <summary>Writes the colour into the fields, except into the one being typed in.</summary>
        private void ShowValues(FieldBox typedIn)
        {
            _updating = true;
            try
            {
                if (typedIn != _hex) _hex.Box.Text = Settings.ToHex(_color);
                if (typedIn != _r) _r.Box.Text = _color.R.ToString(CultureInfo.InvariantCulture);
                if (typedIn != _g) _g.Box.Text = _color.G.ToString(CultureInfo.InvariantCulture);
                if (typedIn != _b) _b.Box.Text = _color.B.ToString(CultureInfo.InvariantCulture);
            }
            finally { _updating = false; }
        }

        private void OnHexTyped()
        {
            if (_updating) return;
            string t = _hex.Box.Text.Trim().TrimStart('#');
            if (t.Length == 3) t = new string(new[] { t[0], t[0], t[1], t[1], t[2], t[2] });   // #abc = #aabbcc
            Color c;
            if (Settings.TryParseHex(t, out c)) SetColor(c, _hex);
        }

        private void OnRgbTyped(FieldBox field)
        {
            if (_updating) return;
            int r, g, b;
            if (!Channel(_r, out r) || !Channel(_g, out g) || !Channel(_b, out b)) return;
            SetColor(Color.FromArgb(r, g, b), field);
        }

        private static bool Channel(FieldBox f, out int v)
        {
            return int.TryParse(f.Box.Text, NumberStyles.None, CultureInfo.InvariantCulture, out v) && v >= 0 && v <= 255;
        }

        /// <summary>Up / Down change a channel by 1 (with Shift by 10).</summary>
        private void StepChannel(FieldBox field, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Up && e.KeyCode != Keys.Down) return;
            int v;
            if (!Channel(field, out v)) v = 0;
            int step = (e.Shift ? 10 : 1) * (e.KeyCode == Keys.Up ? 1 : -1);
            field.Box.Text = Math.Max(0, Math.Min(255, v + step)).ToString(CultureInfo.InvariantCulture);
            field.Box.SelectionStart = field.Box.Text.Length;
            e.Handled = e.SuppressKeyPress = true;
        }

        // ------------------------------------------------------------------ mouse

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button != MouseButtons.Left) return;
            int knob = Look.S(8);
            Rectangle sv = Rectangle.Inflate(_svRect, knob, knob), hue = Rectangle.Inflate(_hueRect, knob, knob);
            int swatch = SwatchAt(e.Location);
            if (swatch >= 0)
            {
                SetColor(_swatches[swatch].Value, null);
                if (e.Clicks == 2) DialogResult = DialogResult.OK;
            }
            else if (sv.Contains(e.Location)) { _drag = Drag.Sv; DragTo(e.Location); }
            else if (hue.Contains(e.Location)) { _drag = Drag.Hue; DragTo(e.Location); }
            else if (_previewRect.Contains(e.Location) && e.X >= _previewRect.X + _previewRect.Width / 2) SetColor(_original, null);   // back to the current colour
            else if (e.Y < Look.S(52)) MoveWindow();
            Focus();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (_drag != Drag.None) { DragTo(e.Location); return; }
            int swatch = SwatchAt(e.Location);
            if (swatch != _hoverSwatch) { _hoverSwatch = swatch; Invalidate(); }
            int knob = Look.S(8);
            bool hand = swatch >= 0 || Rectangle.Inflate(_svRect, knob, knob).Contains(e.Location) || Rectangle.Inflate(_hueRect, knob, knob).Contains(e.Location)
                        || (_previewRect.Contains(e.Location) && e.X >= _previewRect.X + _previewRect.Width / 2);
            Cursor = hand ? Cursors.Hand : Cursors.Default;
        }

        protected override void OnMouseUp(MouseEventArgs e) { _drag = Drag.None; base.OnMouseUp(e); }
        protected override void OnMouseLeave(EventArgs e) { if (_hoverSwatch >= 0) { _hoverSwatch = -1; Invalidate(); } base.OnMouseLeave(e); }

        private void DragTo(Point p)
        {
            if (_drag == Drag.Sv)
            {
                float s = Clamp01((p.X - _svRect.X) / (float)(_svRect.Width - 1));
                float v = 1 - Clamp01((p.Y - _svRect.Y) / (float)(_svRect.Height - 1));
                SetHsv(_h, s, v);
            }
            else if (_drag == Drag.Hue)
            {
                SetHsv(Clamp01((p.X - _hueRect.X) / (float)(_hueRect.Width - 1)) * 359.99f, _s, _v);
            }
        }

        private int SwatchAt(Point p)
        {
            for (int i = 0; i < _swatches.Count; i++)
                if (Rectangle.Inflate(_swatches[i].Key, Look.S(2), Look.S(2)).Contains(p)) return i;
            return -1;
        }

        private void MoveWindow()
        {
            ReleaseCapture();
            SendMessage(Handle, 0x00A1 /* WM_NCLBUTTONDOWN */, new IntPtr(2) /* HTCAPTION */, IntPtr.Zero);
        }

        // ------------------------------------------------------------------ painting

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Look.Quality(g);
            g.Clear(Color.White);

            TextRenderer.DrawText(g, Text, _titleFont, new Point(Look.S(Pad), Look.S(20)), Look.Text, TextFormatFlags.NoPadding);

            // saturation / brightness square
            if (_svBitmap == null || _svBitmapHue != _h || _svBitmap.Size != _svRect.Size)
            {
                if (_svBitmap != null) _svBitmap.Dispose();
                _svBitmap = MakeSvBitmap(_svRect.Size, _h);
                _svBitmapHue = _h;
            }
            using (GraphicsPath p = Look.Round(_svRect, Look.S(6)))
            using (var tb = new TextureBrush(_svBitmap, WrapMode.Clamp))
            {
                tb.TranslateTransform(_svRect.X, _svRect.Y);
                g.FillPath(tb, p);
                using (var pen = new Pen(Color.FromArgb(28, 0, 0, 0), 1f)) g.DrawPath(pen, p);
            }
            DrawKnob(g, new PointF(_svRect.X + _s * (_svRect.Width - 1), _svRect.Y + (1 - _v) * (_svRect.Height - 1)), Look.S(8), _color);

            // hue slider
            using (GraphicsPath p = Look.Round(_hueRect, _hueRect.Height / 2f))
            using (var lb = new LinearGradientBrush(new Rectangle(_hueRect.X - 1, _hueRect.Y, _hueRect.Width + 2, _hueRect.Height), Color.Red, Color.Red, LinearGradientMode.Horizontal))
            {
                var blend = new ColorBlend(7);
                blend.Colors = new[] { Color.FromArgb(255, 0, 0), Color.FromArgb(255, 255, 0), Color.FromArgb(0, 255, 0), Color.FromArgb(0, 255, 255),
                                       Color.FromArgb(0, 0, 255), Color.FromArgb(255, 0, 255), Color.FromArgb(255, 0, 0) };
                blend.Positions = new[] { 0f, 1f / 6, 2f / 6, 3f / 6, 4f / 6, 5f / 6, 1f };
                lb.InterpolationColors = blend;
                g.FillPath(lb, p);
            }
            DrawKnob(g, new PointF(_hueRect.X + _h / 360f * (_hueRect.Width - 1), _hueRect.Y + _hueRect.Height / 2f), Look.S(9), FromHsv(_h, 1, 1));

            // new | current
            using (GraphicsPath p = Look.Round(_previewRect, Look.S(6)))
            {
                Region old = g.Clip;
                g.SetClip(p, CombineMode.Intersect);
                int half = _previewRect.Width / 2;
                using (var b = new SolidBrush(_color)) g.FillRectangle(b, _previewRect.X, _previewRect.Y, half, _previewRect.Height);
                using (var b = new SolidBrush(_original)) g.FillRectangle(b, _previewRect.X + half, _previewRect.Y, _previewRect.Width - half, _previewRect.Height);
                g.Clip = old;
                using (var pen = new Pen(Color.FromArgb(40, 0, 0, 0), 1f)) g.DrawPath(pen, p);
            }
            int capY = _previewRect.Bottom + Look.S(5);
            TextRenderer.DrawText(g, Strings.T("picker.new"), _labelFont, new Point(_previewRect.X, capY), Look.TextSoft, TextFormatFlags.NoPadding);
            TextRenderer.DrawText(g, Strings.T("picker.current"), _labelFont, new Point(_previewRect.X + _previewRect.Width / 2, capY), Look.TextSoft, TextFormatFlags.NoPadding);

            Caption(g, Strings.T("picker.hex"), _hex.Left, _hexLabelY);
            Caption(g, "R", _r.Left, _rgbLabelY);
            Caption(g, "G", _g.Left, _rgbLabelY);
            Caption(g, "B", _b.Left, _rgbLabelY);

            // palette and recent colours
            Caption(g, Strings.T("picker.palette"), Look.S(Pad), _paletteLabelY);
            Caption(g, Strings.T("picker.recent"), Look.S(Pad), _recentLabelY);
            if (_recent.Count == 0)
                TextRenderer.DrawText(g, Strings.T("picker.recent.empty"), _labelFont, new Point(Look.S(Pad), Look.S(_recentRowY + 6)),
                    Color.FromArgb(160, 160, 160), TextFormatFlags.NoPadding);
            bool marked = false;
            for (int i = 0; i < _swatches.Count; i++)
            {
                Rectangle r = _swatches[i].Key;
                Color c = _swatches[i].Value;
                bool selected = !marked && c.ToArgb() == _color.ToArgb();
                if (selected) marked = true;          // mark one swatch only, the first that matches
                DrawSwatch(g, r, c, selected, i == _hoverSwatch);
            }

            // button bar and the window's hairline border
            var bar = new Rectangle(0, ClientSize.Height - Look.S(BarH), ClientSize.Width, Look.S(BarH));
            using (var b = new SolidBrush(Look.Bar)) g.FillRectangle(b, bar);
            using (var pen = new Pen(Look.Line, 1f)) g.DrawLine(pen, 0, bar.Top, bar.Right, bar.Top);
            using (var pen = new Pen(Color.FromArgb(200, 200, 200), 1f)) g.DrawRectangle(pen, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
        }

        private void Caption(Graphics g, string text, int x, int logicalY)
        {
            TextRenderer.DrawText(g, text, _labelFont, new Point(x, Look.S(logicalY)), Look.TextSoft, TextFormatFlags.NoPadding);
        }

        private static void DrawSwatch(Graphics g, Rectangle r, Color c, bool selected, bool hover)
        {
            using (GraphicsPath p = Look.Round(r, Look.S(5)))
            {
                using (var b = new SolidBrush(c)) g.FillPath(b, p);
                double light = (0.299 * c.R + 0.587 * c.G + 0.114 * c.B) / 255;
                if (light > 0.85)
                    using (var pen = new Pen(Color.FromArgb(214, 214, 214), 1f)) g.DrawPath(pen, p);
            }
            if (selected || hover)
            {
                float o = Look.S(3);
                var ring = new RectangleF(r.X - o, r.Y - o, r.Width + 2 * o, r.Height + 2 * o);
                using (GraphicsPath p = Look.Round(ring, Look.S(7)))
                using (var pen = new Pen(selected ? Look.Accent : Color.FromArgb(190, 190, 190), Math.Max(1.5f, 1.6f * Look.Scale)))
                    g.DrawPath(pen, p);
            }
        }

        /// <summary>A round handle: the colour inside a white ring with a soft dark edge.</summary>
        private static void DrawKnob(Graphics g, PointF c, float r, Color fill)
        {
            using (var b = new SolidBrush(Color.FromArgb(70, 0, 0, 0))) g.FillEllipse(b, c.X - r - 1, c.Y - r - 1, 2 * r + 2, 2 * r + 2);
            g.FillEllipse(Brushes.White, c.X - r, c.Y - r, 2 * r, 2 * r);
            float inner = r - Math.Max(2f, 2.5f * Look.Scale);
            using (var b = new SolidBrush(fill)) g.FillEllipse(b, c.X - inner, c.Y - inner, 2 * inner, 2 * inner);
        }

        /// <summary>White to the full hue from left to right, darkening to black from top to bottom.</summary>
        private static Bitmap MakeSvBitmap(Size size, float hue)
        {
            var bmp = new Bitmap(Math.Max(1, size.Width), Math.Max(1, size.Height), System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                var r = new Rectangle(0, 0, bmp.Width, bmp.Height);
                using (var b = new SolidBrush(FromHsv(hue, 1, 1))) g.FillRectangle(b, r);
                using (var b = new LinearGradientBrush(new Rectangle(0, 0, bmp.Width, bmp.Height), Color.White, Color.FromArgb(0, 255, 255, 255), LinearGradientMode.Horizontal))
                {
                    b.WrapMode = WrapMode.TileFlipXY;
                    g.FillRectangle(b, r);
                }
                using (var b = new LinearGradientBrush(new Rectangle(0, 0, bmp.Width, bmp.Height), Color.FromArgb(0, 0, 0, 0), Color.Black, LinearGradientMode.Vertical))
                {
                    b.WrapMode = WrapMode.TileFlipXY;
                    g.FillRectangle(b, r);
                }
            }
            return bmp;
        }

        // ------------------------------------------------------------------ colour maths

        internal static Color FromHsv(float h, float s, float v)
        {
            h = ((h % 360) + 360) % 360;
            float c = v * s, x = c * (1 - Math.Abs((h / 60f) % 2 - 1)), m = v - c;
            float r, g, b;
            if (h < 60) { r = c; g = x; b = 0; }
            else if (h < 120) { r = x; g = c; b = 0; }
            else if (h < 180) { r = 0; g = c; b = x; }
            else if (h < 240) { r = 0; g = x; b = c; }
            else if (h < 300) { r = x; g = 0; b = c; }
            else { r = c; g = 0; b = x; }
            return Color.FromArgb(Byte(r + m), Byte(g + m), Byte(b + m));
        }

        internal static void ToHsv(Color c, out float h, out float s, out float v)
        {
            float r = c.R / 255f, g = c.G / 255f, b = c.B / 255f;
            float max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b)), d = max - min;
            v = max;
            s = max <= 0 ? 0 : d / max;
            if (d <= 0) h = 0;
            else if (max == r) h = 60 * (((g - b) / d) % 6);
            else if (max == g) h = 60 * ((b - r) / d + 2);
            else h = 60 * ((r - g) / d + 4);
            if (h < 0) h += 360;
        }

        private static int Byte(float f) { return Math.Max(0, Math.Min(255, (int)Math.Round(f * 255))); }
        private static float Clamp01(float f) { return f < 0 ? 0 : f > 1 ? 1 : f; }

        // ------------------------------------------------------------------ Win32

        [StructLayout(LayoutKind.Sequential)] private struct RECT { public int L, T, R, B; }
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr h, out RECT r);
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr h);
        [DllImport("user32.dll")] private static extern bool ReleaseCapture();
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr w, IntPtr l);
        [DllImport("dwmapi.dll")] private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
    }
}
