using System;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptSelectionExport
{
    /// <summary>How the selected shapes are laid out as a grid.</summary>
    internal sealed class GridOptions
    {
        /// <summary>Shapes per row; the number of rows follows from how many shapes there are.</summary>
        public int Columns = 1;
        /// <summary>Space between neighbouring columns and rows, in points.</summary>
        public double GapX, GapY;
        /// <summary>Fill the grid by where the shapes are now (top to bottom, left to right) instead of in the order they were selected.</summary>
        public bool ByPosition;
    }

    /// <summary>
    /// "Matrix layout": puts the selected shapes into rows and columns. A column is as wide as its
    /// widest shape and a row as high as its highest; every shape sits in the middle of its cell,
    /// and the gaps are the space between the columns and between the rows. The grid starts at
    /// the top left corner of the area the shapes occupied.
    ///
    /// The layout is always computed from where the shapes were when this object was made, so it
    /// can be applied again and again while the user tries values, and taken back.
    /// </summary>
    internal sealed class GridLayout
    {
        private sealed class Item
        {
            public PowerPoint.Shape Shape;
            public float Left, Top;         // where it was
            public Aligner.Box Bounds;      // what it covered, turned shapes with their turned box
            public int Order;
        }

        private readonly List<Item> _items = new List<Item>();
        private readonly double _left, _top;

        /// <summary>How many shapes take part (connectors glued to shapes follow those and are not counted).</summary>
        public int Count { get { return _items.Count; } }

        public GridLayout(PowerPoint.ShapeRange range)
        {
            int n = range.Count;
            var glued = new List<PowerPoint.Shape>();
            for (int i = 1; i <= n; i++)
            {
                PowerPoint.Shape s = range[i];
                if (Aligner.IsGluedConnector(s)) { glued.Add(s); continue; }
                _items.Add(new Item { Shape = s, Left = s.Left, Top = s.Top, Bounds = Aligner.Box.Of(s), Order = _items.Count });
            }
            if (_items.Count == 0)                  // nothing but connectors selected: lay those out
                foreach (PowerPoint.Shape s in glued)
                    _items.Add(new Item { Shape = s, Left = s.Left, Top = s.Top, Bounds = Aligner.Box.Of(s), Order = _items.Count });

            _left = double.MaxValue; _top = double.MaxValue;
            foreach (Item it in _items)
            {
                _left = Math.Min(_left, it.Bounds.L);
                _top = Math.Min(_top, it.Bounds.T);
            }
        }

        /// <summary>Rows that <paramref name="columns"/> columns give for this many shapes.</summary>
        public static int RowsFor(int count, int columns)
        {
            columns = Math.Max(1, Math.Min(Math.Max(1, count), columns));
            return (count + columns - 1) / columns;
        }

        /// <summary>A grid worked out: which shape goes into which cell, and where the columns and rows are.</summary>
        private sealed class Plan
        {
            public List<Item> Order;
            public int Columns;
            public double[] X, Y, Width, Height;
        }

        private Plan PlanFor(GridOptions options)
        {
            int n = _items.Count;
            int columns = Math.Max(1, Math.Min(n, options.Columns)), rows = RowsFor(n, columns);

            // ---- which shape goes into which cell
            var order = new List<Item>(_items);
            if (options.ByPosition)
            {
                // the topmost shapes are the first row, the next ones the second, ...; each row from left to right
                order.Sort(delegate(Item a, Item b)
                {
                    int c = a.Bounds.CenterY.CompareTo(b.Bounds.CenterY);
                    return c != 0 ? c : a.Order.CompareTo(b.Order);
                });
                for (int r = 0; r < rows; r++)
                {
                    int start = r * columns, length = Math.Min(columns, n - start);
                    order.Sort(start, length, Comparer<Item>.Create(delegate(Item a, Item b)
                    {
                        int c = a.Bounds.CenterX.CompareTo(b.Bounds.CenterX);
                        return c != 0 ? c : a.Order.CompareTo(b.Order);
                    }));
                }
            }

            // ---- the columns and rows
            var width = new double[columns];
            var height = new double[rows];
            for (int i = 0; i < n; i++)
            {
                Aligner.Box b = order[i].Bounds;
                width[i % columns] = Math.Max(width[i % columns], b.R - b.L);
                height[i / columns] = Math.Max(height[i / columns], b.B - b.T);
            }
            var x = new double[columns];
            var y = new double[rows];
            for (int c = 0; c < columns; c++) x[c] = c == 0 ? _left : x[c - 1] + width[c - 1] + options.GapX;
            for (int r = 0; r < rows; r++) y[r] = r == 0 ? _top : y[r - 1] + height[r - 1] + options.GapY;
            return new Plan { Order = order, Columns = columns, X = x, Y = y, Width = width, Height = height };
        }

        /// <summary>Moves the shapes into the grid. Returns how many of them refused to move (locked ones, for example).</summary>
        public int Apply(GridOptions options)
        {
            if (_items.Count == 0) return 0;
            Plan p = PlanFor(options);

            // ---- every shape into the middle of its cell
            int unmoved = 0;
            for (int i = 0; i < p.Order.Count; i++)
            {
                Item it = p.Order[i];
                int c = i % p.Columns, r = i / p.Columns;
                double dx = p.X[c] + p.Width[c] / 2 - it.Bounds.CenterX, dy = p.Y[r] + p.Height[r] / 2 - it.Bounds.CenterY;
                if (!MoveTo(it, (float)(it.Left + dx), (float)(it.Top + dy))) unmoved++;
            }
            return unmoved;
        }

        /// <summary>The area the grid takes up on the slide, in points (left, top, width, height); null when there is nothing to lay out.</summary>
        public double[] Extent(GridOptions options)
        {
            if (_items.Count == 0) return null;
            Plan p = PlanFor(options);
            double l = double.MaxValue, t = double.MaxValue, r = double.MinValue, b = double.MinValue;
            for (int c = 0; c < p.X.Length; c++) { l = Math.Min(l, p.X[c]); r = Math.Max(r, p.X[c] + p.Width[c]); }
            for (int i = 0; i < p.Y.Length; i++) { t = Math.Min(t, p.Y[i]); b = Math.Max(b, p.Y[i] + p.Height[i]); }
            return new[] { l, t, r - l, b - t };
        }

        /// <summary>Puts every shape back where it was.</summary>
        public void Restore()
        {
            foreach (Item it in _items) MoveTo(it, it.Left, it.Top);
        }

        private static bool MoveTo(Item it, float left, float top)
        {
            try
            {
                if (Math.Abs(it.Shape.Left - left) > 0.0005f) it.Shape.Left = left;
                if (Math.Abs(it.Shape.Top - top) > 0.0005f) it.Shape.Top = top;
                return true;
            }
            catch { return false; }
        }
    }

    internal enum CutResult
    {
        Done,
        /// <summary>Not one picture or one closed shape: a line, a group, a table, several objects ...</summary>
        Unsupported,
    }

    /// <summary>
    /// "Matrix cut": cuts one picture or one shape into rows and columns of separate pieces that
    /// together look like the original.
    ///
    /// A piece of a picture is the whole picture again, cropped to its cell; all pieces share the
    /// one image file. A piece of a shape is the shape intersected with the cell (PowerPoint's
    /// Merge Shapes): it keeps the fill and the outline; cells the shape does not reach give no
    /// piece. Text on a shape cannot be cut and is left out. A turned original is cut upright
    /// and the pieces are turned back around its centre.
    /// </summary>
    internal static class MatrixCut
    {
        public const int MaxCount = 20;

        private const Office.MsoTriState True = Office.MsoTriState.msoTrue;

        private enum Kind { Unsupported, Picture, Shape }

        private static Kind KindOf(PowerPoint.Shape shape)
        {
            try
            {
                if (shape.Connector == True) return Kind.Unsupported;
                Office.MsoShapeType type = shape.Type;
                if (type == Office.MsoShapeType.msoPlaceholder)
                {
                    // a picture dropped into a placeholder cuts like any picture; other placeholders are left alone
                    if (shape.PlaceholderFormat.ContainedType != Office.MsoShapeType.msoPicture) return Kind.Unsupported;
                    type = Office.MsoShapeType.msoPicture;
                }
                switch (type)
                {
                    case Office.MsoShapeType.msoPicture:
                    case Office.MsoShapeType.msoLinkedPicture:
                        return Kind.Picture;
                    case Office.MsoShapeType.msoAutoShape:
                    case Office.MsoShapeType.msoFreeform:
                    case Office.MsoShapeType.msoTextBox:
                        return Kind.Shape;
                }
            }
            catch { }
            return Kind.Unsupported;
        }

        /// <summary>Whether this is something that can be cut: one picture, or one closed shape.</summary>
        public static bool CanCut(PowerPoint.Shape shape) { return KindOf(shape) != Kind.Unsupported; }

        /// <summary>Whether cutting would leave text out (a shape with text on it).</summary>
        public static bool HasText(PowerPoint.Shape shape)
        {
            try { return KindOf(shape) == Kind.Shape && shape.HasTextFrame == True && shape.TextFrame.HasText == True; }
            catch { return false; }
        }

        /// <summary>
        /// Cuts the shape into <paramref name="columns"/> x <paramref name="rows"/> pieces, which take its
        /// place in the stacking order, and removes it. <paramref name="pieces"/> are the new shapes.
        /// </summary>
        public static CutResult Cut(PowerPoint.Shapes shapes, PowerPoint.Shape shape, int columns, int rows, out List<PowerPoint.Shape> pieces)
        {
            pieces = new List<PowerPoint.Shape>();
            Kind kind = KindOf(shape);
            if (kind == Kind.Unsupported) return CutResult.Unsupported;
            columns = Math.Max(1, Math.Min(MaxCount, columns));
            rows = Math.Max(1, Math.Min(MaxCount, rows));

            float rotation = shape.Rotation;
            string name = shape.Name;
            double centerX = shape.Left + shape.Width / 2.0, centerY = shape.Top + shape.Height / 2.0;

            // an upright copy to cut from; the original stays untouched until the pieces are there
            PowerPoint.Shape work = Copy(shape, shape.Left, shape.Top);
            try
            {
                if (Math.Abs(rotation) > 0.001f) work.Rotation = 0;         // turns about its centre
                if (kind == Kind.Shape)
                {
                    // Merge Shapes would hand the text to every piece
                    try { if (work.HasTextFrame == True && work.TextFrame.HasText == True) work.TextFrame.TextRange.Text = ""; } catch { }
                }
                float left = work.Left, top = work.Top, width = work.Width, height = work.Height;

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < columns; c++)
                    {
                        // cell edges from the whole, so that neighbours share them exactly
                        float x0 = left + width * c / columns, x1 = left + width * (c + 1) / columns;
                        float y0 = top + height * r / rows, y1 = top + height * (r + 1) / rows;
                        PowerPoint.Shape piece = kind == Kind.Picture
                            ? PicturePiece(work, left, top, x0, y0, x1 - x0, y1 - y0)
                            : ShapePiece(shapes, work, left, top, x0, y0, x1 - x0, y1 - y0);
                        if (piece == null) continue;                        // the shape does not reach this cell
                        try { piece.Name = name + " " + (r + 1) + "-" + (c + 1); } catch { }
                        pieces.Add(piece);
                    }
                }
            }
            finally
            {
                try { work.Delete(); } catch { }
            }
            if (pieces.Count == 0) return CutResult.Done;       // nothing came out: the original stays

            // ---- turn the pieces back around the original's centre
            if (Math.Abs(rotation) > 0.001f)
            {
                double a = rotation * Math.PI / 180, cos = Math.Cos(a), sin = Math.Sin(a);
                foreach (PowerPoint.Shape piece in pieces)
                {
                    double px = piece.Left + piece.Width / 2.0 - centerX, py = piece.Top + piece.Height / 2.0 - centerY;
                    double nx = centerX + px * cos - py * sin, ny = centerY + px * sin + py * cos;
                    piece.Rotation = rotation;                              // about its own centre, which stays put
                    piece.Left = (float)(nx - piece.Width / 2.0);
                    piece.Top = (float)(ny - piece.Height / 2.0);
                }
            }

            // ---- into the original's place in the stacking order, then the original goes
            pieces = TakePlaceOf(shapes, shape, pieces);
            shape.Delete();
            return CutResult.Done;
        }

        /// <summary>A duplicate lying exactly on the given position (Duplicate puts its copy a little to the side, on top of everything).</summary>
        private static PowerPoint.Shape Copy(PowerPoint.Shape shape, float left, float top)
        {
            PowerPoint.Shape copy = shape.Duplicate()[1];
            copy.Left = left;
            copy.Top = top;
            return copy;
        }

        /// <summary>The picture again, cropped to one cell. The picture itself stays where it is; only the frame changes.</summary>
        private static PowerPoint.Shape PicturePiece(PowerPoint.Shape work, float left, float top, float x, float y, float w, float h)
        {
            PowerPoint.Shape piece = Copy(work, left, top);
            Office.Crop crop = piece.PictureFormat.Crop;
            crop.ShapeWidth = w;
            crop.ShapeHeight = h;
            crop.ShapeLeft = x;
            crop.ShapeTop = y;
            return piece;
        }

        /// <summary>The shape intersected with one cell; null when they do not overlap.</summary>
        private static PowerPoint.Shape ShapePiece(PowerPoint.Shapes shapes, PowerPoint.Shape work, float left, float top, float x, float y, float w, float h)
        {
            PowerPoint.Shape copy = Copy(work, left, top);
            shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, x, y, w, h);
            int count = shapes.Count;                       // the copy and the cell are the two topmost shapes
            shapes.Range(new object[] { count - 1, count }).MergeShapes(Office.MsoMergeCmd.msoMergeIntersect, copy);
            // one shape instead of two: the piece, on top. None at all: nothing in this cell.
            return shapes.Count == count - 1 ? shapes[shapes.Count] : null;
        }

        /// <summary>
        /// Moves the pieces (which are on top of everything) to just above <paramref name="original"/>,
        /// keeping their order. Several pieces travel as one group, so that the walk through the
        /// stacking order is made once.
        /// </summary>
        private static List<PowerPoint.Shape> TakePlaceOf(PowerPoint.Shapes shapes, PowerPoint.Shape original, List<PowerPoint.Shape> pieces)
        {
            try
            {
                PowerPoint.Shape mover = pieces[0];
                bool grouped = pieces.Count > 1;
                if (grouped)
                {
                    var indices = new object[pieces.Count];
                    for (int i = 0; i < pieces.Count; i++) indices[i] = pieces[i].ZOrderPosition;
                    mover = shapes.Range(indices).Group();
                }
                int above = mover.ZOrderPosition - original.ZOrderPosition - 1;       // shapes between the original and the pieces
                if (above > original.ZOrderPosition)
                {
                    // shorter from below: to the very back, then up past the original
                    mover.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                    for (int guard = 0; mover.ZOrderPosition < original.ZOrderPosition && guard < 5000; guard++) mover.ZOrder(Office.MsoZOrderCmd.msoBringForward);
                }
                else
                {
                    for (int guard = 0; mover.ZOrderPosition > original.ZOrderPosition + 1 && guard < 5000; guard++) mover.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                }
                if (!grouped) return pieces;

                PowerPoint.ShapeRange freed = mover.Ungroup();
                var result = new List<PowerPoint.Shape>();
                for (int i = 1; i <= freed.Count; i++) result.Add(freed[i]);
                return result;
            }
            catch { return pieces; }        // they stay on top; everything else is done
        }
    }
}
