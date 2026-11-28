using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace PptSelectionExport
{
    /// <summary>A slider in the look of the add-in's windows: a thin track, filled up to a round knob.</summary>
    internal sealed class SlimSlider : Control
    {
        public int Minimum = 0, Maximum = 100;
        private int _value;
        private bool _drag;
        public event EventHandler ValueChanged;

        public SlimSlider()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            BackColor = Color.White;
            Cursor = Cursors.Hand;
            TabStop = true;
        }

        public int Value
        {
            get { return _value; }
            set
            {
                int v = Math.Max(Minimum, Math.Min(Maximum, value));
                if (v == _value) return;
                _value = v;
                Invalidate();
                if (ValueChanged != null) ValueChanged(this, EventArgs.Empty);
            }
        }

        private int Knob { get { return Look.S(8); } }

        private void SetFromX(int x)
        {
            double f = (x - Knob) / (double)Math.Max(1, Width - 2 * Knob);
            Value = Minimum + (int)Math.Round(f * (Maximum - Minimum));
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left) { _drag = true; Focus(); SetFromX(e.X); }
            base.OnMouseDown(e);
        }
        protected override void OnMouseMove(MouseEventArgs e) { if (_drag) SetFromX(e.X); base.OnMouseMove(e); }
        protected override void OnMouseUp(MouseEventArgs e) { _drag = false; base.OnMouseUp(e); }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }
        protected override bool IsInputKey(Keys keyData)
        {
            return keyData == Keys.Left || keyData == Keys.Right || keyData == Keys.Up || keyData == Keys.Down || base.IsInputKey(keyData);
        }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            int step = e.KeyCode == Keys.Left || e.KeyCode == Keys.Down ? -1 : e.KeyCode == Keys.Right || e.KeyCode == Keys.Up ? 1
                     : e.KeyCode == Keys.PageDown ? -10 : e.KeyCode == Keys.PageUp ? 10 : 0;
            if (step != 0) { Value = _value + step; e.Handled = true; }
            base.OnKeyDown(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Look.Quality(g);
            g.Clear(BackColor);
            float cy = Height / 2f, t = Math.Max(3f, Look.S(4));
            float x0 = Knob, x1 = Width - Knob, x = x0 + (x1 - x0) * (_value - Minimum) / (float)Math.Max(1, Maximum - Minimum);
            using (GraphicsPath track = Look.Round(new RectangleF(x0, cy - t / 2, x1 - x0, t), t / 2))
            using (var b = new SolidBrush(Look.Line)) g.FillPath(b, track);
            if (x - x0 > t)
                using (GraphicsPath done = Look.Round(new RectangleF(x0, cy - t / 2, x - x0, t), t / 2))
                using (var b = new SolidBrush(Look.Accent)) g.FillPath(b, done);
            float r = Knob - 1f;
            using (var b = new SolidBrush(Color.White)) g.FillEllipse(b, x - r, cy - r, 2 * r, 2 * r);
            using (var pen = new Pen(Look.Accent, Math.Max(1.5f, Look.Scale * (Focused && ShowFocusCues ? 2.5f : 1.8f)))) g.DrawEllipse(pen, x - r, cy - r, 2 * r, 2 * r);
        }
    }

    /// <summary>A check box with its text, in the look of the add-in's windows.</summary>
    internal sealed class CheckLine : Control
    {
        private bool _checked;
        public event EventHandler CheckedChanged;

        public CheckLine()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            BackColor = Color.White;
            Cursor = Cursors.Hand;
            TabStop = true;
            AccessibleRole = AccessibleRole.CheckButton;
        }

        public bool Checked
        {
            get { return _checked; }
            set
            {
                if (value == _checked) return;
                _checked = value;
                Invalidate();
                if (CheckedChanged != null) CheckedChanged(this, EventArgs.Empty);
            }
        }

        protected override void OnClick(EventArgs e) { Focus(); Checked = !_checked; base.OnClick(e); }
        protected override void OnTextChanged(EventArgs e) { Invalidate(); base.OnTextChanged(e); }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }
        protected override bool IsInputKey(Keys keyData) { return keyData == Keys.Space || base.IsInputKey(keyData); }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Space) { Checked = !_checked; e.Handled = true; }
            base.OnKeyDown(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Look.Quality(g);
            g.Clear(BackColor);
            int size = Look.S(16);
            var box = new RectangleF(0.5f, (Height - size) / 2f, size, size);
            using (GraphicsPath p = Look.Round(box, Look.S(3)))
            {
                using (var b = new SolidBrush(_checked ? Look.Accent : Color.White)) g.FillPath(b, p);
                using (var pen = new Pen(_checked ? Look.Accent : (Focused && ShowFocusCues ? Look.TextSoft : Look.Border), 1f)) g.DrawPath(pen, p);
            }
            if (_checked)
                using (var pen = new Pen(Color.White, Math.Max(1.6f, Look.Scale * 1.8f)) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                    g.DrawLines(pen, new[]
                    {
                        new PointF(box.X + size * 0.24f, box.Y + size * 0.52f), new PointF(box.X + size * 0.43f, box.Y + size * 0.70f),
                        new PointF(box.X + size * 0.77f, box.Y + size * 0.31f)
                    });
            TextRenderer.DrawText(g, Text, Font, new Rectangle(size + Look.S(8), 0, Width - size - Look.S(8), Height), Look.Text,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.SingleLine | TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis);
        }
    }

    /// <summary>
    /// "Slim down": shows what takes up the space in the presentation, lets the user choose how
    /// far to shrink the pictures, compresses a temporary copy in the background and, when the
    /// user likes the result, saves it under a new name. The open presentation is never touched.
    /// </summary>
    internal sealed class SlimForm : Form
    {
        private const int W = 600, H = 556, BarH = 64, Pad = 24;
        private const int PpiTop = 166, QualityTop = 256, CheckTop = 316, ResultTop = 358;
        private static readonly int[] PpiChoices = { 0, 330, 220, 150, 96 };

        private readonly IWin32Window _owner;
        private readonly Settings _settings;
        private readonly string _input, _output, _name, _folder, _original, _extension;
        private readonly SlimOverview _overview;

        private readonly CloseButton _close = new CloseButton();
        private readonly FlatButton[] _ppi = new FlatButton[PpiChoices.Length];
        private readonly SlimSlider _quality = new SlimSlider { Minimum = 30, Maximum = 100 };
        private readonly CheckLine _toJpeg = new CheckLine();
        private readonly FlatButton _start = new FlatButton(), _save = new FlatButton();
        private Font _titleFont, _labelFont, _valueFont, _bigFont, _buttonFont, _smallFont;

        private int _ppiIndex;
        private SlimReport _report;         // the result that is waiting to be saved; null when there is none (or the settings changed since)
        private bool _hadReport;
        private volatile bool _cancel;
        private bool _working, _closeWhenDone;
        private int _done, _total;
        private string _error;

        /// <summary>Where the slimmed copy was saved; null when it was not.</summary>
        public string SavedPath { get; private set; }

        /// <param name="input">A copy of the presentation to work on (a temporary file).</param>
        /// <param name="workFolder">A temporary folder for the result until it is saved.</param>
        /// <param name="name">The presentation's name without extension, for the suggested file name.</param>
        /// <param name="folder">Where the save dialog starts.</param>
        /// <param name="original">The open presentation's own file, which must not be overwritten ("" when it has none).</param>
        public SlimForm(string input, string workFolder, string name, string folder, string original, Settings settings, IWin32Window owner)
        {
            _input = input;
            _extension = Path.GetExtension(input).ToLowerInvariant();      // .pptx, or .pptm for a presentation with macros
            _output = Path.Combine(workFolder, "slim" + _extension);
            _name = name;
            _folder = folder;
            _original = original ?? "";
            _settings = settings;
            _owner = owner;
            _overview = Slimmer.Inspect(input);

            Text = Strings.T("slim.title");
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            BackColor = Color.White;
            KeyPreview = true;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);

            _close.AccessibleName = Strings.T("picker.close");
            _close.Click += delegate { Close(); };
            Controls.Add(_close);

            _ppiIndex = Math.Max(0, Array.IndexOf(PpiChoices, settings.SlimPpi));
            for (int i = 0; i < _ppi.Length; i++)
            {
                int index = i;
                _ppi[i] = new FlatButton { Text = PpiChoices[i] == 0 ? Strings.T("slim.ppi.keep") : PpiChoices[i] + " ppi", Primary = i == _ppiIndex, BackColor = Color.White };
                _ppi[i].Click += delegate { ChoosePpi(index); };
                Controls.Add(_ppi[i]);
            }
            _quality.Value = settings.SlimQuality;
            _quality.AccessibleName = Strings.T("slim.quality.caption");
            _quality.ValueChanged += delegate { SettingsChanged(); };
            _toJpeg.Text = Strings.T("slim.tojpeg");
            _toJpeg.Checked = settings.SlimToJpeg;
            _toJpeg.CheckedChanged += delegate { SettingsChanged(); };
            _start.Click += delegate { StartOrCancel(); };
            _save.Text = Strings.T("slim.saveas");
            _save.Click += delegate { Save(); };
            Controls.AddRange(new Control[] { _quality, _toJpeg, _save, _start });
            UpdateButtons();
        }

        // ------------------------------------------------------------ what the user does

        private void ChoosePpi(int index)
        {
            if (_working || index == _ppiIndex) return;
            _ppiIndex = index;
            for (int i = 0; i < _ppi.Length; i++) { _ppi[i].Primary = i == index; _ppi[i].Invalidate(); }
            SettingsChanged();
        }

        /// <summary>A result made with other settings is not offered for saving any more.</summary>
        private void SettingsChanged()
        {
            if (_report != null) _hadReport = true;
            _report = null;
            _error = null;
            UpdateButtons();
            Invalidate();
        }

        private SlimOptions Options()
        {
            return new SlimOptions { Ppi = PpiChoices[_ppiIndex], Quality = _quality.Value, PhotosToJpeg = _toJpeg.Checked };
        }

        private void StartOrCancel()
        {
            if (_working) { _cancel = true; return; }

            SlimOptions options = Options();
            _settings.SlimPpi = options.Ppi;
            _settings.SlimQuality = options.Quality;
            _settings.SlimToJpeg = options.PhotosToJpeg;
            _settings.Save();

            _report = null; _error = null; _hadReport = false; SavedPath = null;
            _cancel = false; _working = true; _done = 0; _total = 0;
            UpdateButtons();
            Invalidate();

            // GDI+ and the zip only: nothing of PowerPoint is touched from this thread
            var worker = new Thread(delegate()
            {
                SlimReport report = null;
                string error = null;
                try
                {
                    report = Slimmer.Slim(_input, _output, options,
                        delegate(int done, int total) { Post(delegate { _done = done; _total = total; Invalidate(); }); },
                        delegate { return _cancel; });
                }
                catch (OperationCanceledException) { }
                catch (Exception ex) { error = ex.Message; }
                Post(delegate { Finished(report, error); });
            });
            worker.IsBackground = true;
            worker.Name = "Litchi slim";
            worker.Start();
        }

        private void Post(MethodInvoker action)
        {
            try { if (IsHandleCreated && !IsDisposed) BeginInvoke(action); }
            catch { }
        }

        private void Finished(SlimReport report, string error)
        {
            _working = false;
            _report = report;
            _error = error;
            UpdateButtons();
            Invalidate();
            if (_closeWhenDone) Close();
        }

        private bool WorthSaving
        {
            get { return _report != null && _report.NewFileBytes < _report.OldFileBytes * 0.99; }
        }

        private void UpdateButtons()
        {
            _start.Text = Strings.T(_working ? "slim.cancel" : "slim.start");
            _save.Enabled = !_working && WorthSaving;
            _save.Primary = _save.Enabled;
            _start.Primary = !_save.Enabled && !_working;
            foreach (FlatButton b in _ppi) b.Enabled = !_working;
            _quality.Enabled = _toJpeg.Enabled = !_working;
            _start.Invalidate(); _save.Invalidate();
        }

        private void Save()
        {
            if (_working || !WorthSaving) return;
            string path;
            using (var dlg = new SaveFileDialog())
            {
                dlg.Title = Strings.T("slim.dlg");
                dlg.Filter = Strings.T(_extension == ".pptm" ? "filter.pptm" : "filter.pptx");
                dlg.DefaultExt = _extension.TrimStart('.');
                dlg.AddExtension = true;
                dlg.OverwritePrompt = true;
                if (!string.IsNullOrEmpty(_folder) && Directory.Exists(_folder)) dlg.InitialDirectory = _folder;
                dlg.FileName = _name + Strings.T("slim.suffix");
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                path = dlg.FileName;
            }
            if (!path.EndsWith(_extension, StringComparison.OrdinalIgnoreCase)) path += _extension;
            if (_original.Length > 0 && string.Equals(Path.GetFullPath(path), _original, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(this, Strings.T("msg.slim.samefile"), Strings.T("slim.title"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try { File.Copy(_output, path, true); }
            catch (Exception ex)
            {
                MessageBox.Show(this, Strings.T("msg.slim.failed", ex.Message) + Strings.T("msg.slim.locked"), Strings.T("slim.title"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            SavedPath = path;
            Invalidate();
            if (_settings.OpenFolder)
            {
                try { Process.Start("explorer.exe", "/select,\"" + path + "\""); } catch { }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_working)
            {
                // let the worker stop first: it is writing into the temporary folder
                e.Cancel = true;
                _cancel = true;
                _closeWhenDone = true;
            }
            base.OnFormClosing(e);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (_save.Enabled) Save(); else StartOrCancel();
                e.Handled = e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Escape) { Close(); e.Handled = e.SuppressKeyPress = true; }
            base.OnKeyDown(e);
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
            _valueFont = Look.Font(13.5f, FontStyle.Regular);
            _bigFont = Look.Font(21, FontStyle.Bold);
            _buttonFont = Look.Font(13.5f, FontStyle.Regular);
            _smallFont = Look.Font(12, FontStyle.Regular);
            Size = new Size(Look.S(W), Look.S(H));

            _close.SetBounds(Look.S(W - 46) - 1, 1, Look.S(46), Look.S(34));
            const int gap = 8;
            double each = (W - 2 * Pad - gap * (_ppi.Length - 1)) / (double)_ppi.Length;
            for (int i = 0; i < _ppi.Length; i++)
            {
                _ppi[i].Font = _buttonFont;
                _ppi[i].SetBounds(Look.S(Pad + i * (each + gap)), Look.S(PpiTop + 24), Look.S(each), Look.S(32));
            }
            _quality.SetBounds(Look.S(Pad) - Look.S(8), Look.S(QualityTop + 22), Look.S(W - 2 * Pad - 64) + Look.S(16), Look.S(26));
            _toJpeg.Font = _valueFont;
            _toJpeg.SetBounds(Look.S(Pad), Look.S(CheckTop), Look.S(W - 2 * Pad), Look.S(26));

            int by = H - BarH + (BarH - 36) / 2;
            _save.SetBounds(Look.S(W - Pad - 112), Look.S(by), Look.S(112), Look.S(36));
            _start.SetBounds(Look.S(W - Pad - 112 - 10 - 112), Look.S(by), Look.S(112), Look.S(36));
            _save.BackColor = _start.BackColor = Look.Bar;
            _save.Font = _start.Font = _buttonFont;
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
            ActiveControl = _start;
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

        // ------------------------------------------------------------ painting

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

            // ---- what is in the file
            Draw(g, Strings.T("slim.overview", Bytes(_overview.FileBytes)), _valueFont, Pad, 62, Look.Text);
            var bar = new RectangleF(Look.S(Pad), Look.S(92), Look.S(W - 2 * Pad), Look.S(8));
            Color[] colors = { Look.Accent, Color.FromArgb(120, 120, 120), Color.FromArgb(205, 205, 205) };
            long[] parts = { _overview.ImageBytes, _overview.MediaBytes, _overview.OtherBytes };
            using (GraphicsPath clip = Look.Round(bar, bar.Height / 2))
            {
                g.SetClip(clip);
                float x = bar.X;
                for (int i = 0; i < parts.Length; i++)
                {
                    float w = _overview.FileBytes > 0 ? bar.Width * parts[i] / _overview.FileBytes : 0;
                    if (i == parts.Length - 1) w = bar.Right - x;
                    if (w > 0) using (var b = new SolidBrush(colors[i])) g.FillRectangle(b, x, bar.Y, w + 1, bar.Height);
                    x += w;
                }
                g.ResetClip();
            }
            string[] legend =
            {
                Strings.T("slim.part.images", _overview.ImageCount, Bytes(_overview.ImageBytes)),
                Strings.T("slim.part.media", _overview.MediaCount == 0 ? Strings.T("slim.none") : Bytes(_overview.MediaBytes)),
                Strings.T("slim.part.other", Bytes(_overview.OtherBytes))
            };
            int lx = Look.S(Pad), dot = Look.S(8);
            for (int i = 0; i < legend.Length; i++)
            {
                using (var b = new SolidBrush(colors[i])) g.FillEllipse(b, lx, Look.S(112) + Look.S(4), dot, dot);
                lx += dot + Look.S(6);
                TextRenderer.DrawText(g, legend[i], _labelFont, new Point(lx, Look.S(112)), Look.TextSoft, Plain);
                lx += TextRenderer.MeasureText(g, legend[i], _labelFont, Size.Empty, Plain).Width + Look.S(18);
            }
            using (var pen = new Pen(Look.Line, 1f)) g.DrawLine(pen, Look.S(Pad), Look.S(148), Look.S(W - Pad), Look.S(148));

            // ---- the settings
            Draw(g, Strings.T("slim.ppi.caption"), _labelFont, Pad, PpiTop, Look.TextSoft);
            Draw(g, Strings.T("slim.ppi.hint." + PpiChoices[_ppiIndex]), _smallFont, Pad, PpiTop + 62, Color.FromArgb(150, 150, 150));
            Draw(g, Strings.T("slim.quality.caption"), _labelFont, Pad, QualityTop, Look.TextSoft);
            TextRenderer.DrawText(g, _quality.Value + "%", _valueFont, new Rectangle(Look.S(W - Pad - 60), Look.S(QualityTop + 22), Look.S(60), Look.S(26)), Look.Text,
                TextFormatFlags.Right | TextFormatFlags.VerticalCenter | Plain);
            using (var pen = new Pen(Look.Line, 1f)) g.DrawLine(pen, Look.S(Pad), Look.S(ResultTop - 8), Look.S(W - Pad), Look.S(ResultTop - 8));

            PaintResult(g);

            var bottom = new Rectangle(0, ClientSize.Height - Look.S(BarH), ClientSize.Width, Look.S(BarH));
            using (var b = new SolidBrush(Look.Bar)) g.FillRectangle(b, bottom);
            using (var pen = new Pen(Look.Line, 1f)) g.DrawLine(pen, 0, bottom.Top, bottom.Right, bottom.Top);
            using (var pen = new Pen(Color.FromArgb(200, 200, 200), 1f)) g.DrawRectangle(pen, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
        }

        private void PaintResult(Graphics g)
        {
            int y = ResultTop;
            var line = new Rectangle(Look.S(Pad), 0, Look.S(W - 2 * Pad), Look.S(20));
            if (_working)
            {
                Draw(g, _total > 0 ? Strings.T("slim.working", Math.Min(_done + 1, _total), _total) : Strings.T("slim.preparing"), _valueFont, Pad, y + 4, Look.Text);
                var track = new RectangleF(Look.S(Pad), Look.S(y + 36), Look.S(W - 2 * Pad), Look.S(6));
                using (GraphicsPath p = Look.Round(track, track.Height / 2))
                using (var b = new SolidBrush(Look.Line)) g.FillPath(b, p);
                float done = _total > 0 ? track.Width * _done / _total : 0;
                if (done > track.Height)
                    using (GraphicsPath p = Look.Round(new RectangleF(track.X, track.Y, done, track.Height), track.Height / 2))
                    using (var b = new SolidBrush(Look.Accent)) g.FillPath(b, p);
                return;
            }
            if (_error != null)
            {
                line.Y = Look.S(y + 4); line.Height = Look.S(60);
                TextRenderer.DrawText(g, Strings.T("msg.slim.failed", _error), _valueFont, line, Look.Danger, TextFormatFlags.NoPadding | TextFormatFlags.WordBreak);
                return;
            }
            if (_report == null)
            {
                line.Y = Look.S(y + 4); line.Height = Look.S(60);
                TextRenderer.DrawText(g, Strings.T(_hadReport ? "slim.changed" : "slim.idle"), _valueFont, line, Look.TextSoft, TextFormatFlags.NoPadding | TextFormatFlags.WordBreak);
                return;
            }

            // ---- the result: before, after, and the pictures that gave the most
            string sizes = Bytes(_report.OldFileBytes) + "  →  " + Bytes(_report.NewFileBytes);
            Draw(g, sizes, _bigFont, Pad, y - 2, Look.Text);
            if (WorthSaving)
            {
                int percent = (int)Math.Round(100.0 * (_report.OldFileBytes - _report.NewFileBytes) / _report.OldFileBytes);
                int sw = TextRenderer.MeasureText(g, sizes, _bigFont, Size.Empty, Plain).Width;
                TextRenderer.DrawText(g, "−" + percent + "%", _bigFont, new Point(Look.S(Pad) + sw + Look.S(16), Look.S(y - 2)), Look.Accent, Plain);
            }
            int changed = _report.ChangedCount;
            Draw(g, WorthSaving ? Strings.T("slim.result.detail", changed, _report.Images.Count - changed) : Strings.T("slim.result.none"), _labelFont, Pad, y + 32, Look.TextSoft);

            var top = new List<SlimImage>();
            foreach (SlimImage i in _report.Images) if (i.Changed) top.Add(i);
            top.Sort(delegate(SlimImage a, SlimImage b) { return (b.OldBytes - b.NewBytes).CompareTo(a.OldBytes - a.NewBytes); });
            Color faint = Color.FromArgb(140, 140, 140);
            for (int i = 0; i < Math.Min(3, top.Count) && WorthSaving; i++)
            {
                SlimImage im = top[i];
                int ry = y + 56 + i * 17;
                TextRenderer.DrawText(g, im.Name, _smallFont, new Rectangle(Look.S(Pad), Look.S(ry), Look.S(150), Look.S(17)), faint, Plain | TextFormatFlags.EndEllipsis);
                Draw(g, im.OldPixels.Width + "×" + im.OldPixels.Height + " → " + im.NewPixels.Width + "×" + im.NewPixels.Height, _smallFont, Pad + 158, ry, faint);
                TextRenderer.DrawText(g, Bytes(im.OldBytes) + " → " + Bytes(im.NewBytes), _smallFont, new Rectangle(Look.S(W - Pad - 190), Look.S(ry), Look.S(190), Look.S(17)), faint, Plain | TextFormatFlags.Right);
            }
            if (SavedPath != null)
            {
                line.Y = Look.S(H - BarH - 24);
                TextRenderer.DrawText(g, Strings.T("slim.saved", SavedPath), _labelFont, line, Color.FromArgb(30, 136, 90), Plain | TextFormatFlags.PathEllipsis);
            }
        }

        /// <summary>"23.4 MB", "212 KB".</summary>
        internal static string Bytes(long bytes)
        {
            if (bytes >= 100L * 1024 * 1024) return (bytes / 1048576.0).ToString("0", CultureInfo.InvariantCulture) + " MB";
            if (bytes >= 1024 * 1024) return (bytes / 1048576.0).ToString("0.0", CultureInfo.InvariantCulture) + " MB";
            if (bytes >= 1024) return (bytes / 1024.0).ToString("0", CultureInfo.InvariantCulture) + " KB";
            return bytes.ToString(CultureInfo.InvariantCulture) + " B";
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                foreach (Font f in new[] { _titleFont, _labelFont, _valueFont, _bigFont, _buttonFont, _smallFont }) if (f != null) f.Dispose();
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
