using System;
using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PptSelectionExport
{
    /// <summary>
    /// What the add-in's windows have in common: no frame, white, a title in the top left corner
    /// that the window can be dragged by, a close button, and a bar for the buttons at the bottom.
    /// </summary>
    internal abstract class LitchiForm : Form
    {
        protected const int BarH = 64, Pad = 24;

        private readonly IWin32Window _owner;
        private readonly CloseButton _close = new CloseButton();
        private Font _titleFont;

        protected LitchiForm(string title, IWin32Window owner)
        {
            _owner = owner;
            Text = title;
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            BackColor = Color.White;
            KeyPreview = true;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            _close.AccessibleName = Strings.T("picker.close");
            _close.Click += delegate { DialogResult = DialogResult.Cancel; };
            Controls.Add(_close);
        }

        /// <summary>The window's size in logical pixels (at 96 dpi).</summary>
        protected abstract Size LogicalSize { get; }

        /// <summary>Places the controls; called once the display scaling is known (Look.Scale).</summary>
        protected abstract void LayoutControls();

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
            _titleFont = Look.Font(17, FontStyle.Bold);
            Size logical = LogicalSize;
            Size = new Size(Look.S(logical.Width), Look.S(logical.Height));
            _close.SetBounds(Look.S(logical.Width - 46) - 1, 1, Look.S(46), Look.S(34));
            LayoutControls();
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

        /// <summary>Where the owner window is on the screen; empty when there is none.</summary>
        protected Rectangle OwnerBounds()
        {
            RECT r;
            return _owner != null && _owner.Handle != IntPtr.Zero && GetWindowRect(_owner.Handle, out r) ? Rectangle.FromLTRB(r.L, r.T, r.R, r.B) : Rectangle.Empty;
        }

        protected const TextFormatFlags Plain = TextFormatFlags.NoPadding | TextFormatFlags.SingleLine;

        protected static void Draw(Graphics g, string text, Font font, int x, int y, Color color)
        {
            TextRenderer.DrawText(g, text, font, new Point(Look.S(x), Look.S(y)), color, Plain);
        }

        /// <summary>Background, title, button bar and the hairline around the window.</summary>
        protected void PaintChrome(Graphics g)
        {
            Look.Quality(g);
            g.Clear(Color.White);
            Draw(g, Text, _titleFont, Pad, 20, Look.Text);
            var bar = new Rectangle(0, ClientSize.Height - Look.S(BarH), ClientSize.Width, Look.S(BarH));
            using (var b = new SolidBrush(Look.Bar)) g.FillRectangle(b, bar);
            using (var pen = new Pen(Look.Line, 1f)) g.DrawLine(pen, 0, bar.Top, bar.Right, bar.Top);
            using (var pen = new Pen(Color.FromArgb(200, 200, 200), 1f)) g.DrawRectangle(pen, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
        }

        /// <summary>The usual two buttons at the right end of the bar: Cancel, then the primary one.</summary>
        protected void PlaceBarButtons(FlatButton cancel, FlatButton ok, Font font)
        {
            Size logical = LogicalSize;
            int by = logical.Height - BarH + (BarH - 36) / 2;
            ok.SetBounds(Look.S(logical.Width - Pad - 96), Look.S(by), Look.S(96), Look.S(36));
            cancel.SetBounds(Look.S(logical.Width - Pad - 96 - 10 - 96), Look.S(by), Look.S(96), Look.S(36));
            ok.BackColor = cancel.BackColor = Look.Bar;
            ok.Font = cancel.Font = font;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && _titleFont != null) _titleFont.Dispose();
            base.Dispose(disposing);
        }

        [StructLayout(LayoutKind.Sequential)] private struct RECT { public int L, T, R, B; }
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr h, out RECT r);
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr h);
        [DllImport("user32.dll")] private static extern bool ReleaseCapture();
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hwnd, int msg, IntPtr w, IntPtr l);
        [DllImport("dwmapi.dll")] private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
    }

    /// <summary>
    /// Lets PowerPoint bring its window up to date while a modal window of ours is open.
    ///
    /// PowerPoint redraws a slide that an add-in has changed when it next gets idle time. While
    /// an add-in's modal window runs the message loop, Office's component manager hands out idle
    /// time only through a timer that fires 0.9 seconds after things went quiet, so a live
    /// preview would trail every keystroke by almost a second. This class watches that timer
    /// message go by (it is addressed to a window of the class "MsoStdCompMgr") and, a moment
    /// after a change, posts the same message again: the idle time comes at once and the slide
    /// follows within a few hundredths of a second.
    ///
    /// Nothing here is needed for correctness. Until the message has been seen once in a
    /// PowerPoint session, and on a version of Office that does this differently, nothing is
    /// posted and the slide follows after those 0.9 seconds.
    /// </summary>
    internal sealed class HostIdle : IMessageFilter
    {
        private const int WM_TIMER = 0x0113;
        private const string ManagerClass = "MsoStdCompMgr";
        private const int Delay = 30, Times = 3;        // the queue has to be empty when the message arrives: not at once, and again to be sure

        private static readonly HostIdle Watcher = new HostIdle();
        private static IntPtr _window, _id, _callback;  // the component manager's timer as last seen
        private static int _watching, _left;
        private static Timer _soon;

        private HostIdle() { }

        /// <summary>To be called when a modal window opens (true) and when it has closed (false).</summary>
        public static void Watch(bool on)
        {
            if (on)
            {
                if (_watching++ == 0) Application.AddMessageFilter(Watcher);
            }
            else if (_watching > 0 && --_watching == 0)
            {
                Application.RemoveMessageFilter(Watcher);
                if (_soon != null) _soon.Stop();
            }
        }

        public bool PreFilterMessage(ref Message m)
        {
            if (m.Msg == WM_TIMER && m.LParam != IntPtr.Zero && (m.HWnd == _window || ClassOf(m.HWnd) == ManagerClass))
            {
                _window = m.HWnd;
                _id = m.WParam;
                _callback = m.LParam;
            }
            return false;
        }

        /// <summary>The slide has just been changed: have PowerPoint show it in a moment rather than in a second.</summary>
        public static void Soon()
        {
            if (_watching == 0 || _callback == IntPtr.Zero) return;
            if (_soon == null)
            {
                _soon = new Timer();
                _soon.Interval = Delay;
                _soon.Tick += delegate { Post(); };
            }
            _left = Times;
            _soon.Stop();
            _soon.Start();
        }

        private static void Post()
        {
            if (--_left <= 0) _soon.Stop();
            try
            {
                uint process;
                if (_callback != IntPtr.Zero && IsWindow(_window) && GetWindowThreadProcessId(_window, out process) == GetCurrentThreadId() && ClassOf(_window) == ManagerClass)
                {
                    PostMessage(_window, WM_TIMER, _id, _callback);
                    return;
                }
            }
            catch { }
            _callback = IntPtr.Zero;                    // that window is gone: wait until the message is seen again
            _soon.Stop();
        }

        private static string ClassOf(IntPtr window)
        {
            var name = new System.Text.StringBuilder(64);
            GetClassName(window, name, name.Capacity);
            return name.ToString();
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr h, System.Text.StringBuilder name, int max);
        [DllImport("user32.dll", EntryPoint = "PostMessageW")] private static extern bool PostMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr h);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr h, out uint process);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    }

    internal enum MatrixMode { Layout, Cut }

    /// <summary>
    /// The window of "matrix layout" and "matrix cut": how many across and how many down, and for
    /// the layout the gaps between them. While it is open as a layout window it tells its owner
    /// about every change, so that the shapes can follow on the slide right away.
    /// </summary>
    internal sealed class MatrixForm : LitchiForm
    {
        private const int W = 440, FieldW = 180, Cross = 32;
        private const double MinGap = -10, MaxGap = 50;         // centimetres

        private readonly MatrixMode _mode;
        private readonly int _count;                // layout: how many shapes there are to place
        private readonly string _note;              // cut: a remark under the hint, or null
        private readonly Settings _settings;
        private readonly FieldBox _columns = new FieldBox(), _rows = new FieldBox(), _gapX = new FieldBox(), _gapY = new FieldBox();
        private readonly CheckLine _byPosition = new CheckLine();
        private readonly FlatButton _ok = new FlatButton(), _cancel = new FlatButton();
        private Font _labelFont, _fieldFont, _buttonFont, _smallFont;
        private bool _filling;

        /// <summary>Raised whenever the values change and are usable (layout only).</summary>
        public event EventHandler OptionsChanged;

        /// <param name="count">Layout: the number of shapes to lay out.</param>
        /// <param name="note">Cut: a remark shown under the hint, or null.</param>
        public MatrixForm(MatrixMode mode, int count, string note, Settings settings, IWin32Window owner)
            : base(Strings.T(mode == MatrixMode.Layout ? "matrix.layout.title" : "matrix.cut.title"), owner)
        {
            _mode = mode;
            _count = count;
            _note = note;
            _settings = settings;

            _filling = true;
            if (mode == MatrixMode.Layout)
            {
                int columns = Math.Max(1, (int)Math.Ceiling(Math.Sqrt(count)));
                _columns.Box.Text = columns.ToString(CultureInfo.InvariantCulture);
                _rows.Box.Text = GridLayout.RowsFor(count, columns).ToString(CultureInfo.InvariantCulture);
                _gapX.Box.Text = Number(settings.MatrixGapX);
                _gapY.Box.Text = Number(settings.MatrixGapY);
                _byPosition.Checked = settings.MatrixByPosition;
            }
            else
            {
                _columns.Box.Text = settings.CutColumns.ToString(CultureInfo.InvariantCulture);
                _rows.Box.Text = settings.CutRows.ToString(CultureInfo.InvariantCulture);
            }
            _filling = false;

            _columns.Box.AccessibleName = Strings.T("matrix.columns");
            _rows.Box.AccessibleName = Strings.T("matrix.rows");
            _gapX.Box.AccessibleName = Strings.T("matrix.gapx");
            _gapY.Box.AccessibleName = Strings.T("matrix.gapy");
            foreach (FieldBox f in new[] { _columns, _rows, _gapX, _gapY }) f.Box.MaxLength = 8;
            _columns.Box.TextChanged += delegate { Typed(_columns); };
            _rows.Box.TextChanged += delegate { Typed(_rows); };
            _gapX.Box.TextChanged += delegate { Typed(_gapX); };
            _gapY.Box.TextChanged += delegate { Typed(_gapY); };
            foreach (FieldBox f in new[] { _columns, _rows })
            {
                FieldBox field = f;
                f.Box.KeyDown += delegate(object s, KeyEventArgs e) { Step(field, e, 1, 0); };
            }
            foreach (FieldBox f in new[] { _gapX, _gapY })
            {
                FieldBox field = f;
                f.Box.KeyDown += delegate(object s, KeyEventArgs e) { Step(field, e, 0.1, 2); };
            }
            _byPosition.Text = Strings.T("matrix.byposition");
            _byPosition.CheckedChanged += delegate { Changed(); };
            _ok.Primary = true;
            _ok.Text = Strings.T("picker.ok");
            _ok.Click += delegate { Accept(); };
            _cancel.Text = Strings.T("picker.cancel");
            _cancel.Click += delegate { DialogResult = DialogResult.Cancel; };

            Controls.AddRange(new Control[] { _columns, _rows });
            if (mode == MatrixMode.Layout) Controls.AddRange(new Control[] { _gapX, _gapY, _byPosition });
            Controls.AddRange(new Control[] { _cancel, _ok });
            _ok.Enabled = Valid;
        }

        // ------------------------------------------------------------ values

        private static string Number(double v) { return v.ToString("0.##", CultureInfo.InvariantCulture); }

        private int MaxCount { get { return _mode == MatrixMode.Layout ? _count : MatrixCut.MaxCount; } }

        /// <summary>A whole number from 1 to the largest count allowed here; 0 when the field holds something else.</summary>
        private int CountIn(FieldBox field)
        {
            int v;
            return int.TryParse(field.Box.Text.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out v) && v >= 1 && v <= MaxCount ? v : 0;
        }

        /// <summary>A gap in centimetres; NaN when the field holds something else.</summary>
        private static double GapIn(FieldBox field)
        {
            string t = field.Box.Text.Trim().ToLowerInvariant().Replace("厘米", "").Replace("cm", "").Replace(',', '.').Trim();
            double v;
            return double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out v) && !double.IsNaN(v) && v >= MinGap && v <= MaxGap ? v : double.NaN;
        }

        public int Columns { get { return CountIn(_columns); } }

        /// <summary>The rows there will be: for a layout what the columns give, for a cut what the field says.</summary>
        public int Rows { get { return _mode == MatrixMode.Layout ? (Columns > 0 ? GridLayout.RowsFor(_count, Columns) : 0) : CountIn(_rows); } }

        public bool Valid
        {
            get
            {
                if (Columns == 0 || CountIn(_rows) == 0) return false;
                if (_mode == MatrixMode.Cut) return Columns * Rows > 1;
                return !double.IsNaN(GapIn(_gapX)) && !double.IsNaN(GapIn(_gapY));
            }
        }

        /// <summary>The layout the fields describe, gaps in points.</summary>
        public GridOptions Grid()
        {
            return new GridOptions
            {
                Columns = Columns,
                GapX = GapIn(_gapX) * SlideResizer.PointsPerCm,
                GapY = GapIn(_gapY) * SlideResizer.PointsPerCm,
                ByPosition = _byPosition.Checked
            };
        }

        private void Typed(FieldBox field)
        {
            if (_filling) return;
            if (_mode == MatrixMode.Layout && (field == _columns || field == _rows))
            {
                // across times down has to hold all the shapes: the other field follows
                int typed = CountIn(field);
                if (typed > 0)
                {
                    FieldBox other = field == _columns ? _rows : _columns;
                    _filling = true;
                    other.Box.Text = GridLayout.RowsFor(_count, typed).ToString(CultureInfo.InvariantCulture);     // the same sum both ways
                    _filling = false;
                }
            }
            Changed();
        }

        /// <summary>Up and Down change the number in a field by one step.</summary>
        private void Step(FieldBox field, KeyEventArgs e, double step, int decimals)
        {
            if (e.KeyCode != Keys.Up && e.KeyCode != Keys.Down) return;
            double v;
            string t = field.Box.Text.Trim().Replace(',', '.');
            if (double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out v))
            {
                v = Math.Round(v + (e.KeyCode == Keys.Up ? step : -step), decimals);
                double min = decimals == 0 ? 1 : MinGap, max = decimals == 0 ? MaxCount : MaxGap;
                field.Box.Text = Number(Math.Max(min, Math.Min(max, v)));
                field.Box.SelectAll();
            }
            e.Handled = e.SuppressKeyPress = true;
        }

        private void Changed()
        {
            _ok.Enabled = Valid;
            Invalidate();
            if (_mode != MatrixMode.Layout || !Valid || OptionsChanged == null) return;
            OptionsChanged(this, EventArgs.Empty);
            HostIdle.Soon();                    // the shapes have moved: have the slide show it now
        }

        private void Accept()
        {
            if (!Valid) return;
            if (_mode == MatrixMode.Layout)
            {
                _settings.MatrixGapX = GapIn(_gapX);
                _settings.MatrixGapY = GapIn(_gapY);
                _settings.MatrixByPosition = _byPosition.Checked;
            }
            else
            {
                _settings.CutColumns = Columns;
                _settings.CutRows = Rows;
            }
            _settings.Save();
            DialogResult = DialogResult.OK;
        }

        // ------------------------------------------------------------ window

        private int HintTop { get { return _mode == MatrixMode.Layout ? 268 : 164; } }

        protected override Size LogicalSize
        {
            get { return new Size(W, HintTop + (_mode == MatrixMode.Cut && _note != null ? 58 : 36) + BarH); }
        }

        protected override void LayoutControls()
        {
            _labelFont = Look.Font(12.5f, FontStyle.Regular);
            _fieldFont = Look.Font(13.5f, FontStyle.Regular);
            _buttonFont = Look.Font(13.5f, FontStyle.Regular);
            _smallFont = Look.Font(12, FontStyle.Regular);
            foreach (FieldBox f in new[] { _columns, _rows, _gapX, _gapY }) f.Font = _fieldFont;
            _columns.SetBounds(Look.S(Pad), Look.S(110), Look.S(FieldW), Look.S(34));
            _rows.SetBounds(Look.S(Pad + FieldW + Cross), Look.S(110), Look.S(FieldW), Look.S(34));
            _gapX.SetBounds(Look.S(Pad), Look.S(178), Look.S(FieldW), Look.S(34));
            _gapY.SetBounds(Look.S(Pad + FieldW + Cross), Look.S(178), Look.S(FieldW), Look.S(34));
            _byPosition.Font = _fieldFont;
            _byPosition.SetBounds(Look.S(Pad), Look.S(228), Look.S(W - 2 * Pad), Look.S(26));
            PlaceBarButtons(_cancel, _ok, _buttonFont);
        }

        /// <summary>
        /// Layout: the part of the screen the laid-out shapes take up. The window opens beside it
        /// rather than in the middle of its owner, where it would cover what it is arranging.
        /// </summary>
        public Rectangle KeepClear;

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            if (!KeepClear.IsEmpty) MoveBeside(KeepClear);
            ActiveControl = _columns.Box;
            _columns.Box.SelectAll();
        }

        private bool _watching;

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (_mode != MatrixMode.Layout || _watching) return;
            _watching = true;
            HostIdle.Watch(true);
            HostIdle.Soon();                    // the proposal is on the slide already: show it
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            if (_watching) HostIdle.Watch(false);
            _watching = false;
            base.OnFormClosed(e);
        }

        /// <summary>To the right of the area, else to its left, below it or above it: the first place with room, inside the owner if possible.</summary>
        private void MoveBeside(Rectangle keep)
        {
            Rectangle work = Screen.FromPoint(new Point(Left + Width / 2, Top + Height / 2)).WorkingArea;
            Rectangle inside = OwnerBounds();
            inside.Intersect(work);
            int gap = Look.S(36);
            foreach (Rectangle area in new[] { inside, work })
            {
                if (area.Width < Width || area.Height < Height) continue;
                int x = Math.Max(area.Left, Math.Min(keep.Left, area.Right - Width));
                int y = Math.Max(area.Top, Math.Min(keep.Top, area.Bottom - Height));
                Point[] spots =
                {
                    new Point(keep.Right + gap, y), new Point(keep.Left - gap - Width, y),
                    new Point(x, keep.Bottom + gap), new Point(x, keep.Top - gap - Height)
                };
                foreach (Point spot in spots)
                {
                    if (!area.Contains(new Rectangle(spot, Size))) continue;
                    Location = spot;
                    return;
                }
            }
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) { Accept(); e.Handled = e.SuppressKeyPress = true; }
            else if (e.KeyCode == Keys.Escape) { DialogResult = DialogResult.Cancel; e.Handled = e.SuppressKeyPress = true; }
            base.OnKeyDown(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            PaintChrome(g);
            Draw(g, _mode == MatrixMode.Layout ? Strings.T("matrix.selected", _count) : Strings.T("matrix.cut.sub"), _labelFont, Pad, 58, Look.TextSoft);

            Draw(g, Strings.T("matrix.columns"), _labelFont, Pad, 88, Look.TextSoft);
            Draw(g, Strings.T("matrix.rows"), _labelFont, Pad + FieldW + Cross, 88, Look.TextSoft);
            TextRenderer.DrawText(g, "×", _fieldFont, new Rectangle(Look.S(Pad + FieldW), Look.S(110), Look.S(Cross), Look.S(34)), Look.TextSoft,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | Plain);
            if (_mode == MatrixMode.Layout)
            {
                Draw(g, Strings.T("matrix.gapx"), _labelFont, Pad, 156, Look.TextSoft);
                Draw(g, Strings.T("matrix.gapy"), _labelFont, Pad + FieldW + Cross, 156, Look.TextSoft);
            }

            string hint;
            Color color = Look.Text;
            if (Columns == 0 || CountIn(_rows) == 0) { hint = Strings.T("matrix.hint.count", MaxCount); color = Look.Danger; }
            else if (_mode == MatrixMode.Cut)
            {
                if (Columns * Rows == 1) { hint = Strings.T("matrix.hint.one"); color = Look.TextSoft; }
                else hint = Strings.T("matrix.hint.cut", Columns, Rows, Columns * Rows);
            }
            else if (!Valid) { hint = Strings.T("matrix.hint.gap"); color = Look.Danger; }
            else hint = Strings.T("matrix.hint.layout", _count, Columns, Rows);
            Draw(g, hint, _fieldFont, Pad, HintTop, color);
            if (_mode == MatrixMode.Cut && _note != null) Draw(g, _note, _smallFont, Pad, HintTop + 26, Color.FromArgb(150, 150, 150));
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_watching) HostIdle.Watch(false);
                _watching = false;
                foreach (Font f in new[] { _labelFont, _fieldFont, _buttonFont, _smallFont }) if (f != null) f.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
