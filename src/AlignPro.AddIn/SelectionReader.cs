using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using AlignPro.Geometry;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AlignPro.AddIn
{
    /// <summary>What the solver needs, read out of PowerPoint in one pass.</summary>
    internal sealed class SelectionSnapshot
    {
        public SelectionSnapshot(
            IReadOnlyList<ShapeSnapshot> shapes,
            SlideMetrics slide,
            int slideId,
            ShapeKey? anchor,
            IReadOnlyList<ShapeKey>? slideOrder = null,
            CurveDefinition? anchorCurve = null,
            string? anchorCurveProblem = null)
        {
            Shapes = shapes;
            Slide = slide;
            SlideId = slideId;
            Anchor = anchor;
            SlideOrder = slideOrder;
            AnchorCurve = anchorCurve;
            AnchorCurveProblem = anchorCurveProblem;
        }

        /// <summary>
        /// The curve the anchor defines, when the caller asked for it and the anchor is a shape that
        /// has one. Null otherwise, with the reason in <see cref="AnchorCurveProblem"/>.
        /// </summary>
        public CurveDefinition? AnchorCurve { get; }

        /// <summary>Why <see cref="AnchorCurve"/> is null, when it was asked for.</summary>
        public string? AnchorCurveProblem { get; }

        public IReadOnlyList<ShapeSnapshot> Shapes { get; }

        public SlideMetrics Slide { get; }

        public int SlideId { get; }

        /// <summary>
        /// The last shape in the selection. PowerPoint preserves selection order (measured, probe 2),
        /// so this is the shape the user selected last.
        /// </summary>
        public ShapeKey? Anchor { get; }

        /// <summary>
        /// Every top-level shape on the slide, back to front, or null when the caller did not ask for
        /// it. Only the ordering verbs need it, and it costs a second pass over the slide, so align
        /// and distribute do not pay for it.
        /// </summary>
        public IReadOnlyList<ShapeKey>? SlideOrder { get; }
    }

    /// <summary>
    /// Turns the live PowerPoint selection into immutable snapshots. This is the only place that reads
    /// shape geometry, so the solver never sees a COM object.
    /// </summary>
    internal static class SelectionReader
    {
        // MsoShapeType values used here, named to keep the intent readable.
        private const int MsoAutoShape = 1;
        private const int MsoFreeform = 5;
        private const int MsoGroup = 6;
        private const int MsoLine = 9;
        private const int MsoPlaceholder = 14;

        /// <summary>
        /// Reads the current selection. Returns null and sets <paramref name="problem"/> when there is
        /// nothing usable selected - that is an ordinary outcome, not an error.
        /// </summary>
        public static SelectionSnapshot? TryRead(
            PowerPoint.Application app,
            double margin,
            out string? problem,
            bool includeSlideOrder = false,
            bool includeAnchorCurve = false)
        {
            problem = null;

            PowerPoint.DocumentWindow? window = null;
            PowerPoint.Selection? selection = null;
            PowerPoint.ShapeRange? range = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Presentation? presentation = null;

            try
            {
                if (app.Windows.Count == 0)
                {
                    problem = "Open a presentation first.";
                    return null;
                }

                window = app.ActiveWindow;
                selection = window.Selection;

                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    problem = "Select one or more shapes on a slide.";
                    return null;
                }

                range = selection.ShapeRange;
                if (range.Count == 0)
                {
                    problem = "Select one or more shapes on a slide.";
                    return null;
                }

                slide = window.View.Slide as PowerPoint.Slide;
                if (slide == null)
                {
                    problem = "AlignPro works on slides, not on the master or notes pages.";
                    return null;
                }

                presentation = slide.Parent as PowerPoint.Presentation;
                if (presentation == null)
                {
                    problem = "Could not read the presentation's slide size.";
                    return null;
                }

                var slideId = slide.SlideID;
                var metrics = new SlideMetrics(
                    presentation.PageSetup.SlideWidth,
                    presentation.PageSetup.SlideHeight,
                    margin,
                    ReadPlaceholderBounds(slide));

                var shapes = new List<ShapeSnapshot>(range.Count);
                ShapeKey? anchor = null;

                for (var i = 1; i <= range.Count; i++)
                {
                    PowerPoint.Shape? shape = null;
                    try
                    {
                        shape = range[i];
                        var snapshot = ReadShape(shape, slideId);
                        shapes.Add(snapshot);
                        anchor = snapshot.Key;   // ends up holding the last one
                    }
                    finally
                    {
                        Com.Release(shape);
                    }
                }

                var slideOrder = includeSlideOrder ? ReadSlideOrder(slide, slideId) : null;

                CurveDefinition? curve = null;
                string? curveProblem = null;
                if (includeAnchorCurve)
                {
                    PowerPoint.Shape? last = null;
                    try
                    {
                        last = range[range.Count];
                        curve = ReadCurve(last, out curveProblem);
                    }
                    finally
                    {
                        Com.Release(last);
                    }
                }

                return new SelectionSnapshot(shapes, metrics, slideId, anchor, slideOrder, curve, curveProblem);
            }
            finally
            {
                Com.Release(presentation);
                Com.Release(slide);
                Com.Release(range);
                Com.Release(selection);
                Com.Release(window);
            }
        }

        /// <summary>
        /// Every top-level shape on the slide, back to front. The <c>Shapes</c> collection is indexed
        /// in z-order, so its position is the stacking order - no need to read <c>ZOrderPosition</c>
        /// per shape, which would be a COM call each.
        /// </summary>
        /// <remarks>
        /// Top-level only, deliberately. A shape inside a group is not in this collection, and its own
        /// z-position is measured within the group rather than against the slide. The solver refuses
        /// such a selection rather than silently restacking the wrong things.
        /// </remarks>
        private static IReadOnlyList<ShapeKey> ReadSlideOrder(PowerPoint.Slide slide, int slideId)
        {
            PowerPoint.Shapes? shapes = null;
            try
            {
                shapes = slide.Shapes;
                var order = new List<ShapeKey>(shapes.Count);

                for (var i = 1; i <= shapes.Count; i++)
                {
                    PowerPoint.Shape? shape = null;
                    try
                    {
                        shape = shapes[i];
                        order.Add(new ShapeKey(slideId, shape.Id));
                    }
                    finally
                    {
                        Com.Release(shape);
                    }
                }

                return order;
            }
            finally
            {
                Com.Release(shapes);
            }
        }

        /// <summary>
        /// The curve a shape defines, converted out of COM and nothing more - the geometry of turning
        /// it into points lives in <see cref="CurveSolver"/>, where it can be tested.
        /// </summary>
        /// <remarks>
        /// Four kinds are understood: an oval, PowerPoint's Arc, a straight line, and a freeform
        /// (which is also what the Curve and Scribble tools draw). Anything else is refused by name
        /// rather than approximated by its frame, which would be a guess the user could not see.
        /// </remarks>
        private static CurveDefinition? ReadCurve(PowerPoint.Shape shape, out string? problem)
        {
            const int msoShapeOval = 9;
            const int msoShapeArc = 25;

            problem = null;
            try
            {
                var type = (int)shape.Type;

                if (type == MsoLine)
                {
                    // A line reports no nodes (probe 9); it runs corner to corner of its frame, and
                    // its flips say which diagonal. The snapshot already carries the frame and flips.
                    return CurveDefinition.Line();
                }

                if (type == MsoFreeform)
                {
                    return ReadNodes(shape, out problem);
                }

                if (type == MsoAutoShape || type == MsoPlaceholder)
                {
                    var autoShape = (int)shape.AutoShapeType;
                    if (autoShape == msoShapeOval) return CurveDefinition.Ellipse();
                    if (autoShape == msoShapeArc) return ReadArc(shape, out problem);
                }

                problem = "The curve (the shape selected last) must be an oval, an arc, a line or a freeform path.";
                return null;
            }
            catch (COMException ex)
            {
                problem = "Could not read the curve from the shape selected last: " + ex.Message;
                return null;
            }
        }

        /// <summary>
        /// The Arc autoshape's two adjustments are its start and end angles, in degrees clockwise
        /// from three o'clock, as directions from the ellipse's centre. See
        /// <c>docs/object-model-findings.md</c>, probe 10 - the frame is not what it seems.
        /// </summary>
        private static CurveDefinition? ReadArc(PowerPoint.Shape shape, out string? problem)
        {
            problem = null;
            PowerPoint.Adjustments? adjustments = null;
            try
            {
                adjustments = shape.Adjustments;
                if (adjustments.Count < 2)
                {
                    problem = "This arc does not report its start and end angles.";
                    return null;
                }

                return CurveDefinition.Arc(adjustments[1], adjustments[2]);
            }
            finally
            {
                Com.Release(adjustments);
            }
        }

        /// <summary>
        /// A freeform's nodes as plain points. Control points are nodes in their own right; which
        /// ones they are is read from the segment types by <see cref="CurveSolver"/>. PowerPoint
        /// reports them already rotated and flipped, exactly where they are drawn (probe 9).
        /// </summary>
        private static CurveDefinition? ReadNodes(PowerPoint.Shape shape, out string? problem)
        {
            const int msoSegmentCurve = 1;

            problem = null;
            PowerPoint.ShapeNodes? nodes = null;
            try
            {
                nodes = shape.Nodes;
                var count = nodes.Count;
                if (count < 2)
                {
                    problem = "The path has fewer than two points.";
                    return null;
                }

                var result = new List<PathNode>(count);
                for (var i = 1; i <= count; i++)
                {
                    PowerPoint.ShapeNode? node = null;
                    try
                    {
                        node = nodes[i];

                        // A one-by-two SAFEARRAY. It can arrive with a lower bound of one rather than
                        // zero, so it is read through Array rather than cast to float[,].
                        var points = (Array)node.Points;
                        var row = points.GetLowerBound(0);
                        var column = points.GetLowerBound(1);
                        var x = Convert.ToDouble(points.GetValue(row, column));
                        var y = Convert.ToDouble(points.GetValue(row, column + 1));

                        var segment = (int)node.SegmentType == msoSegmentCurve
                            ? PathSegmentKind.Curve
                            : PathSegmentKind.Line;
                        result.Add(new PathNode(x, y, segment));
                    }
                    finally
                    {
                        Com.Release(node);
                    }
                }

                return CurveDefinition.Path(result);
            }
            finally
            {
                Com.Release(nodes);
            }
        }

        private static ShapeSnapshot ReadShape(PowerPoint.Shape shape, int slideId)
        {
            // Read each property exactly once. Every access is a COM call, and chained expressions
            // leave RCWs we cannot release.
            var id = shape.Id;
            var left = shape.Left;
            var top = shape.Top;
            var width = shape.Width;
            var height = shape.Height;
            var rotation = shape.Rotation;
            var type = (int)shape.Type;
            var name = shape.Name;

            var flipH = SafeFlip(() => shape.HorizontalFlip);
            var flipV = SafeFlip(() => shape.VerticalFlip);

            return new ShapeSnapshot(
                new ShapeKey(slideId, id),
                // Guard the size: RectD rejects negatives, and a degenerate shape should not throw.
                new RectD(left, top, Math.Max(0, width), Math.Max(0, height)),
                rotation,
                flipH,
                flipV,
                ReadTextBounds(shape),
                isGroup: type == MsoGroup,
                isPlaceholder: type == MsoPlaceholder,
                name: name);
        }

        /// <summary>
        /// Text bounds from <c>TextRange2.Bound*</c>, already in slide coordinates. Returns null for a
        /// shape with no text frame or no text, which the solver treats as "fall back to the frame".
        /// </summary>
        private static RectD? ReadTextBounds(PowerPoint.Shape shape)
        {
            // A mixed pair: PowerPoint defines its own TextFrame2, but its TextRange is the shared
            // Office TextRange2. Getting either namespace wrong here fails to compile.
            PowerPoint.TextFrame2? frame = null;
            Office.TextRange2? text = null;
            try
            {
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue) return null;

                frame = shape.TextFrame2;
                if (frame.HasText != Office.MsoTriState.msoTrue) return null;

                text = frame.TextRange;
                var width = text.BoundWidth;
                var height = text.BoundHeight;
                if (width <= 0 || height <= 0) return null;

                return new RectD(text.BoundLeft, text.BoundTop, width, height);
            }
            catch (COMException)
            {
                // Some shape types claim a text frame then refuse to measure it. Not worth failing over.
                return null;
            }
            finally
            {
                Com.Release(text);
                Com.Release(frame);
            }
        }

        /// <summary>
        /// The body placeholder rectangle from the slide's layout, for the content-placeholder
        /// reference. Null when the layout has no body placeholder.
        /// </summary>
        private static RectD? ReadPlaceholderBounds(PowerPoint.Slide slide)
        {
            const int ppPlaceholderBody = 2;
            const int ppPlaceholderObject = 7;

            PowerPoint.CustomLayout? layout = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Placeholders? placeholders = null;
            try
            {
                layout = slide.CustomLayout;
                shapes = layout.Shapes;
                placeholders = shapes.Placeholders;

                for (var i = 1; i <= placeholders.Count; i++)
                {
                    PowerPoint.Shape? placeholder = null;
                    PowerPoint.PlaceholderFormat? format = null;
                    try
                    {
                        placeholder = placeholders[i];
                        format = placeholder.PlaceholderFormat;
                        var type = (int)format.Type;
                        if (type != ppPlaceholderBody && type != ppPlaceholderObject) continue;

                        return new RectD(
                            placeholder.Left,
                            placeholder.Top,
                            Math.Max(0, placeholder.Width),
                            Math.Max(0, placeholder.Height));
                    }
                    finally
                    {
                        Com.Release(format);
                        Com.Release(placeholder);
                    }
                }

                return null;
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                Com.Release(placeholders);
                Com.Release(shapes);
                Com.Release(layout);
            }
        }

        private static bool SafeFlip(Func<Office.MsoTriState> read)
        {
            try
            {
                return read() == Office.MsoTriState.msoTrue;
            }
            catch (COMException)
            {
                // Lines and a few other shape types have no flip state.
                return false;
            }
        }
    }
}
