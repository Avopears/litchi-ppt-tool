using System;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptSelectionExport
{
    internal enum TransformCommand { RotateLeft, RotateRight, FlipHorizontal, FlipVertical }

    internal enum MergeCommand { Union, Combine, Subtract, Fragment }

    internal enum MergeResult
    {
        Done,
        /// <summary>Fewer than two shapes selected.</summary>
        NeedTwo,
        /// <summary>A group, or shapes inside a group: PowerPoint cannot merge those.</summary>
        Grouped,
    }

    /// <summary>Rotate / flip and the boolean operations (PowerPoint's "Merge Shapes").</summary>
    internal static class ShapeTools
    {
        /// <summary>Turns or mirrors each selected shape in place, like PowerPoint's own commands.</summary>
        public static void Transform(PowerPoint.ShapeRange range, TransformCommand cmd)
        {
            switch (cmd)
            {
                case TransformCommand.RotateLeft: range.IncrementRotation(-90f); break;
                case TransformCommand.RotateRight: range.IncrementRotation(90f); break;
                case TransformCommand.FlipHorizontal: range.Flip(Office.MsoFlipCmd.msoFlipHorizontal); break;
                case TransformCommand.FlipVertical: range.Flip(Office.MsoFlipCmd.msoFlipVertical); break;
            }
        }

        /// <summary>
        /// Merges the shapes. The first selected shape is the primary one: the result takes its
        /// format, and "subtract" cuts the others out of it (as in PowerPoint's Merge Shapes menu).
        /// </summary>
        /// <param name="isChild">True when the shapes were selected inside a group.</param>
        /// <param name="shapes">The collection the shapes live in, used to select the result; may be null.</param>
        public static MergeResult Merge(PowerPoint.ShapeRange range, bool isChild, MergeCommand cmd, PowerPoint.Shapes shapes)
        {
            if (range.Count < 2) return MergeResult.NeedTwo;
            if (isChild) return MergeResult.Grouped;
            for (int i = 1; i <= range.Count; i++)
                if (range[i].Type == Office.MsoShapeType.msoGroup) return MergeResult.Grouped;

            var before = new HashSet<int>();
            if (shapes != null)
                for (int i = 1; i <= shapes.Count; i++) before.Add(shapes[i].Id);

            range.MergeShapes(ToMso(cmd), range[1]);

            // select what came out, like PowerPoint does
            if (shapes != null)
            {
                try
                {
                    var created = new List<object>();
                    for (int i = 1; i <= shapes.Count; i++)
                        if (!before.Contains(shapes[i].Id)) created.Add(i);
                    if (created.Count > 0) shapes.Range(created.ToArray()).Select(Office.MsoTriState.msoTrue);
                }
                catch { }   // no window to select in (a test, for example)
            }
            return MergeResult.Done;
        }

        private static Office.MsoMergeCmd ToMso(MergeCommand cmd)
        {
            switch (cmd)
            {
                case MergeCommand.Union: return Office.MsoMergeCmd.msoMergeUnion;
                case MergeCommand.Combine: return Office.MsoMergeCmd.msoMergeCombine;
                case MergeCommand.Subtract: return Office.MsoMergeCmd.msoMergeSubtract;
                default: return Office.MsoMergeCmd.msoMergeFragment;
            }
        }
    }
}
