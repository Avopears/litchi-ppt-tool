using System;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptSelectionExport
{
    /// <summary>
    /// Changes the slide size and scales what is already on the slides so that it still fits:
    /// PowerPoint's "Ensure Fit". Everything is scaled by the smaller of the two ratios
    /// (new width / old width, new height / old height) about the middle of the slide.
    ///
    /// PowerPoint's object model can only set the width and the height one after the other, and
    /// scales the content at each step on its own: by that step's ratio when the slide shrinks,
    /// not at all when it grows. A slide made smaller in both directions would have its content
    /// shrunk twice, one made larger would keep it small. So this remembers where everything is
    /// and how large its text is, lets PowerPoint resize (which also adapts masters and layouts),
    /// and then puts every object and every text size exactly where the one factor says.
    ///
    /// What is scaled is what PowerPoint scales: position and size, font sizes, text margins,
    /// paragraph spacing and indents. Line weights and effects keep their size, as in PowerPoint.
    /// </summary>
    internal static class SlideResizer
    {
        public const double PointsPerCm = 72.0 / 2.54;
        /// <summary>PowerPoint's limits for a slide: 1 inch to 56 inches.</summary>
        public const double MinCm = 2.54, MaxCm = 142.24;

        private const Office.MsoTriState True = Office.MsoTriState.msoTrue, False = Office.MsoTriState.msoFalse;

        public static double ToCm(double points) { return points / PointsPerCm; }
        public static float ToPoints(double cm) { return (float)(cm * PointsPerCm); }

        /// <summary>The factor the content gets: with it, what fitted on the old slide fits on the new one.</summary>
        public static double FitScale(double oldWidth, double oldHeight, double newWidth, double newHeight)
        {
            return Math.Min(newWidth / oldWidth, newHeight / oldHeight);
        }

        /// <summary>Sets the slide size (in points) and returns the factor the content was scaled by.</summary>
        public static double Resize(PowerPoint.Presentation pres, float newWidth, float newHeight)
        {
            float min = ToPoints(MinCm) - 0.01f, max = ToPoints(MaxCm) + 0.01f;
            if (newWidth < min || newWidth > max || newHeight < min || newHeight > max) throw new ArgumentOutOfRangeException();

            PowerPoint.PageSetup setup = pres.PageSetup;
            float oldWidth = setup.SlideWidth, oldHeight = setup.SlideHeight;
            if (Math.Abs(newWidth - oldWidth) < 0.005f && Math.Abs(newHeight - oldHeight) < 0.005f) return 1;

            // ---- remember: masters and layouts first, they are also put right first (slides inherit from them)
            var layoutText = new List<TextState>();
            var slideText = new List<TextState>();
            var frames = new List<FrameState>();
            PowerPoint.Designs designs = pres.Designs;
            for (int d = 1; d <= designs.Count; d++)
            {
                PowerPoint.Master master = designs[d].SlideMaster;
                CollectText(master.Shapes, layoutText);
                PowerPoint.CustomLayouts layouts = master.CustomLayouts;
                for (int i = 1; i <= layouts.Count; i++) CollectText(layouts[i].Shapes, layoutText);
            }
            PowerPoint.Slides slides = pres.Slides;
            for (int i = 1; i <= slides.Count; i++)
            {
                PowerPoint.Shapes shapes = slides[i].Shapes;
                CollectText(shapes, slideText);
                for (int j = 1; j <= shapes.Count; j++)
                {
                    try { frames.Add(new FrameState(shapes[j])); } catch { }
                }
            }

            // ---- PowerPoint's own resize: the slide, the masters and the layouts
            setup.SlideWidth = newWidth;
            setup.SlideHeight = newHeight;
            // PowerPoint keeps slide sizes in eighths of a point: go by the size it took
            newWidth = setup.SlideWidth;
            newHeight = setup.SlideHeight;
            double rw = newWidth / (double)oldWidth, rh = newHeight / (double)oldHeight;
            double scale = Math.Min(rw, rh);
            double byPowerPoint = Math.Min(rw, 1) * Math.Min(rh, 1);      // what the two steps above did to the content on their own

            // ---- and now exactly
            foreach (TextState t in layoutText) t.Apply(scale);
            foreach (TextState t in slideText) t.Apply(scale);
            if (Math.Abs(scale / byPowerPoint - 1) > 1e-6)
                foreach (FrameState f in frames) f.Apply(scale, byPowerPoint, oldWidth, oldHeight, newWidth, newHeight);
            return scale;
        }

        // ------------------------------------------------------------------ position and size

        /// <summary>Where a shape on a slide was before the resize.</summary>
        private sealed class FrameState
        {
            private readonly PowerPoint.Shape _shape;
            private readonly float _left, _top, _width, _height;
            private readonly bool _placeholder, _glued;

            public FrameState(PowerPoint.Shape shape)
            {
                _shape = shape;
                _left = shape.Left; _top = shape.Top; _width = shape.Width; _height = shape.Height;
                _placeholder = shape.Type == Office.MsoShapeType.msoPlaceholder;
                try
                {
                    if (shape.Connector == True)
                    {
                        PowerPoint.ConnectorFormat c = shape.ConnectorFormat;
                        _glued = c.BeginConnected == True && c.EndConnected == True;
                    }
                }
                catch { }
            }

            public void Apply(double scale, double byPowerPoint, float oldWidth, float oldHeight, float newWidth, float newHeight)
            {
                try
                {
                    if (_glued) return;         // it follows the shapes it is glued to

                    double cx = _left + _width / 2.0, cy = _top + _height / 2.0;
                    if (_placeholder)
                    {
                        // A placeholder that takes its place from the layout was moved along with the
                        // layout, and stays there. One that was moved by hand was treated like any
                        // other shape: scaled about the middle of the slide.
                        double w = _shape.Width, h = _shape.Height, x = _shape.Left + w / 2, y = _shape.Top + h / 2;
                        double ex = newWidth / 2.0 + (cx - oldWidth / 2.0) * byPowerPoint, ey = newHeight / 2.0 + (cy - oldHeight / 2.0) * byPowerPoint;
                        bool likeAShape = Near(w, _width * byPowerPoint) && Near(h, _height * byPowerPoint) && Near(x, ex) && Near(y, ey);
                        if (!likeAShape) return;
                    }

                    bool locked = false;
                    try { locked = _shape.LockAspectRatio == True; } catch { }
                    _shape.Width = (float)(_width * scale);
                    if (!locked) _shape.Height = (float)(_height * scale);      // with the proportions locked the height has followed
                    // a text box that sizes itself to its text takes the height the text needs: centre whatever it is
                    double nx = newWidth / 2.0 + (cx - oldWidth / 2.0) * scale, ny = newHeight / 2.0 + (cy - oldHeight / 2.0) * scale;
                    _shape.Left = (float)(nx - _shape.Width / 2.0);
                    _shape.Top = (float)(ny - _shape.Height / 2.0);
                }
                catch { }       // a locked or unusual object: PowerPoint's own result stays
            }

            private static bool Near(double a, double b) { return Math.Abs(a - b) <= 0.5 + 0.01 * Math.Abs(b); }
        }

        // ------------------------------------------------------------------ text

        private static void CollectText(PowerPoint.Shapes shapes, List<TextState> list)
        {
            int n = shapes.Count;
            for (int i = 1; i <= n; i++)
            {
                try { CollectText(shapes[i], list); } catch { }
            }
        }

        private static void CollectText(PowerPoint.Shape shape, List<TextState> list)
        {
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                PowerPoint.GroupShapes items = shape.GroupItems;        // flat: the shapes of groups inside it too
                for (int i = 1; i <= items.Count; i++)
                {
                    try { CollectText(items[i], list); } catch { }
                }
                return;
            }
            if (shape.HasTable == True)
            {
                PowerPoint.Table table = shape.Table;
                int rows = table.Rows.Count, columns = table.Columns.Count;
                for (int r = 1; r <= rows; r++)
                    for (int c = 1; c <= columns; c++)
                    {
                        try { list.Add(new TextState(table.Cell(r, c).Shape)); } catch { }
                    }
                return;
            }
            if (shape.HasTextFrame == True) list.Add(new TextState(shape));
        }

        /// <summary>The sizes in one text frame before the resize: margins, and per paragraph spacing, indents and font sizes.</summary>
        private sealed class TextState
        {
            private readonly PowerPoint.Shape _shape;
            private readonly float _left, _top, _right, _bottom;
            private readonly List<ParagraphState> _paragraphs = new List<ParagraphState>();

            private sealed class ParagraphState
            {
                public float Before = float.NaN, After = float.NaN, Within = float.NaN;     // NaN: counted in lines, which need no scaling
                public float LeftIndent, FirstLineIndent;
                public float[] FontSizes;
            }

            public TextState(PowerPoint.Shape shape)
            {
                _shape = shape;
                PowerPoint.TextFrame2 frame = shape.TextFrame2;
                _left = frame.MarginLeft; _top = frame.MarginTop; _right = frame.MarginRight; _bottom = frame.MarginBottom;
                Office.TextRange2 text = frame.TextRange;
                int count = text.get_Paragraphs(-1, -1).Count;
                for (int i = 1; i <= count; i++)
                {
                    Office.TextRange2 paragraph = text.get_Paragraphs(i, 1);
                    Office.ParagraphFormat2 format = paragraph.ParagraphFormat;
                    var p = new ParagraphState();
                    if (format.LineRuleBefore == False) p.Before = format.SpaceBefore;
                    if (format.LineRuleAfter == False) p.After = format.SpaceAfter;
                    if (format.LineRuleWithin == False) p.Within = format.SpaceWithin;
                    p.LeftIndent = format.LeftIndent;
                    p.FirstLineIndent = format.FirstLineIndent;
                    int runs = paragraph.get_Runs(-1, -1).Count;
                    p.FontSizes = new float[runs];
                    for (int j = 1; j <= runs; j++) p.FontSizes[j - 1] = paragraph.get_Runs(j, 1).Font.Size;
                    _paragraphs.Add(p);
                }
            }

            /// <summary>
            /// Gives every size its value times <paramref name="scale"/>. A value that is there already
            /// is not written: text that takes its format from the master has it as soon as the master
            /// is right, and should go on taking it from there.
            /// </summary>
            public void Apply(double scale)
            {
                try
                {
                    PowerPoint.TextFrame2 frame = _shape.TextFrame2;
                    if (Differs(frame.MarginLeft, _left * scale)) frame.MarginLeft = (float)(_left * scale);
                    if (Differs(frame.MarginTop, _top * scale)) frame.MarginTop = (float)(_top * scale);
                    if (Differs(frame.MarginRight, _right * scale)) frame.MarginRight = (float)(_right * scale);
                    if (Differs(frame.MarginBottom, _bottom * scale)) frame.MarginBottom = (float)(_bottom * scale);

                    Office.TextRange2 text = frame.TextRange;
                    int count = Math.Min(_paragraphs.Count, text.get_Paragraphs(-1, -1).Count);
                    for (int i = 1; i <= count; i++)
                    {
                        ParagraphState p = _paragraphs[i - 1];
                        Office.TextRange2 paragraph = text.get_Paragraphs(i, 1);
                        try
                        {
                            Office.ParagraphFormat2 format = paragraph.ParagraphFormat;
                            if (!float.IsNaN(p.Before) && Differs(format.SpaceBefore, p.Before * scale)) format.SpaceBefore = (float)(p.Before * scale);
                            if (!float.IsNaN(p.After) && Differs(format.SpaceAfter, p.After * scale)) format.SpaceAfter = (float)(p.After * scale);
                            if (!float.IsNaN(p.Within) && Differs(format.SpaceWithin, p.Within * scale)) format.SpaceWithin = (float)(p.Within * scale);
                            // the left indent first: PowerPoint refuses a hanging indent that reaches left of the text box
                            if (Differs(format.LeftIndent, p.LeftIndent * scale)) format.LeftIndent = (float)(p.LeftIndent * scale);
                            if (Differs(format.FirstLineIndent, p.FirstLineIndent * scale)) format.FirstLineIndent = (float)(p.FirstLineIndent * scale);
                        }
                        catch { }

                        int runs = Math.Min(p.FontSizes.Length, paragraph.get_Runs(-1, -1).Count);
                        for (int j = 1; j <= runs; j++)
                        {
                            double size = Math.Max(1, Math.Min(4000, p.FontSizes[j - 1] * scale));
                            if (p.FontSizes[j - 1] <= 0) continue;
                            Office.Font2 font = paragraph.get_Runs(j, 1).Font;
                            if (Differs(font.Size, size)) font.Size = (float)size;
                        }
                    }
                }
                catch { }
            }

            private static bool Differs(float current, double target)
            {
                return Math.Abs(current - target) > 0.006 + 0.0005 * Math.Abs(target);
            }
        }
    }
}
