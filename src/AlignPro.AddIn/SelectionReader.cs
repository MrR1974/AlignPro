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
            string? anchorCurveProblem = null,
            bool insideGroup = false,
            double groupRotation = 0)
        {
            Shapes = shapes;
            Slide = slide;
            SlideId = slideId;
            Anchor = anchor;
            SlideOrder = slideOrder;
            AnchorCurve = anchorCurve;
            AnchorCurveProblem = anchorCurveProblem;
            InsideGroup = insideGroup;
            GroupRotation = groupRotation;
        }

        /// <summary>
        /// The angle of the group the shapes are inside, in degrees clockwise, or 0 at slide level.
        /// The solver works along the group's own axes when it is turned; see <see cref="GroupSpace"/>.
        /// </summary>
        public double GroupRotation { get; }

        /// <summary>
        /// True when the shapes were selected inside a group rather than on the slide. They then all
        /// belong to one group (probe 11), whose frame is <see cref="SlideMetrics.GroupBounds"/>.
        /// </summary>
        public bool InsideGroup { get; }

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
        /// True when no shape is selected on the slide in view - nothing at all, or only slides in the
        /// thumbnail pane. Tidy then works on the whole slide.
        /// </summary>
        public static bool NothingSelected(PowerPoint.Application app)
        {
            PowerPoint.DocumentWindow? window = null;
            PowerPoint.Selection? selection = null;
            try
            {
                if (app.Windows.Count == 0) return false;

                window = app.ActiveWindow;
                selection = window.Selection;
                return selection.Type == PowerPoint.PpSelectionType.ppSelectionNone ||
                       selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides;
            }
            catch (COMException)
            {
                return false;
            }
            finally
            {
                Com.Release(selection);
                Com.Release(window);
            }
        }

        /// <summary>
        /// Every visible top-level shape on the slide in view, for a verb that works on the whole slide
        /// when nothing is selected. Hidden shapes are left out: moving what nobody can see is never what
        /// was meant. A group comes back as one shape, as it does when selected.
        /// </summary>
        public static SelectionSnapshot? TryReadSlide(PowerPoint.Application app, double margin, out string? problem)
        {
            problem = null;

            PowerPoint.DocumentWindow? window = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Presentation? presentation = null;
            PowerPoint.Shapes? all = null;

            try
            {
                if (app.Windows.Count == 0)
                {
                    problem = "Open a presentation first.";
                    return null;
                }

                window = app.ActiveWindow;
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

                all = slide.Shapes;
                var shapes = new List<ShapeSnapshot>(all.Count);
                for (var i = 1; i <= all.Count; i++)
                {
                    PowerPoint.Shape? shape = null;
                    try
                    {
                        shape = all[i];
                        if (shape.Visible == Office.MsoTriState.msoFalse) continue;
                        shapes.Add(ReadShape(shape, slideId));
                    }
                    finally
                    {
                        Com.Release(shape);
                    }
                }

                return new SelectionSnapshot(shapes, metrics, slideId, anchor: null);
            }
            finally
            {
                Com.Release(all);
                Com.Release(presentation);
                Com.Release(slide);
                Com.Release(window);
            }
        }

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
            PowerPoint.ShapeRange? groupRange = null;
            PowerPoint.Shape? group = null;
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

                // Shapes picked inside a group: ShapeRange then holds the group itself, and the shapes
                // actually picked are the child range, in selection order (probe 11).
                RectD? groupBounds = null;
                var groupRotation = 0.0;
                var insideGroup = selection.HasChildShapeRange;
                if (insideGroup)
                {
                    groupRange = range;
                    group = groupRange[1];
                    range = selection.ChildShapeRange;

                    // The group's frame is its unrotated one, like any shape's, turned about its centre.
                    groupRotation = group.Rotation;
                    groupBounds = new RectD(group.Left, group.Top, Math.Max(0, group.Width), Math.Max(0, group.Height));
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
                    ReadPlaceholderBounds(slide),
                    groupBounds);

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

                IReadOnlyList<ShapeKey>? slideOrder = null;
                if (includeSlideOrder)
                {
                    slideOrder = insideGroup ? ReadGroupOrder(group!, slideId, out problem) : ReadSlideOrder(slide, slideId);
                    if (slideOrder == null) return null;
                }

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

                return new SelectionSnapshot(
                    shapes, metrics, slideId, anchor, slideOrder, curve, curveProblem, insideGroup, groupRotation);
            }
            finally
            {
                Com.Release(presentation);
                Com.Release(slide);
                Com.Release(range);
                Com.Release(group);
                Com.Release(groupRange);
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
        /// The shapes inside a group, back to front - the stacking order a selection inside it is
        /// restacked within. Null, with the reason, when the group holds another group.
        /// </summary>
        /// <remarks>
        /// <para>
        /// <c>GroupItems</c> is not in z-order, so the order is read from <c>ZOrderPosition</c>, which
        /// is (probe 18).
        /// </para>
        /// <para>
        /// A nested group is flattened in <c>GroupItems</c> and every leaf reports the outer group as
        /// its parent, yet bringing a leaf to the front only restacks it within its inner group (probe
        /// 19). Nothing says which inner group a leaf is in, so no ordering across them can be trusted.
        /// The one trace an inner group leaves is a z-position of its own that no leaf holds, so a gap
        /// in the leaves' positions means the group holds another, and the selection is refused.
        /// </para>
        /// </remarks>
        private static IReadOnlyList<ShapeKey>? ReadGroupOrder(PowerPoint.Shape group, int slideId, out string? problem)
        {
            problem = null;
            PowerPoint.GroupShapes? items = null;
            try
            {
                items = group.GroupItems;
                var stack = new List<(int Z, int Id)>(items.Count);
                for (var i = 1; i <= items.Count; i++)
                {
                    PowerPoint.Shape? child = null;
                    try
                    {
                        child = items[i];
                        stack.Add((child.ZOrderPosition, child.Id));
                    }
                    finally
                    {
                        Com.Release(child);
                    }
                }

                stack.Sort((a, b) => a.Z.CompareTo(b.Z));
                if (stack.Count > 0 && stack[stack.Count - 1].Z - stack[0].Z + 1 != stack.Count)
                {
                    problem = "This group has another group inside it, and PowerPoint only restacks a shape among the shapes of its own inner group. AlignPro cannot reorder inside it yet - ungroup the inner group first.";
                    return null;
                }

                var order = new List<ShapeKey>(stack.Count);
                foreach (var entry in stack) order.Add(new ShapeKey(slideId, entry.Id));
                return order;
            }
            finally
            {
                Com.Release(items);
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
            var isConnector = SafeFlip(() => shape.Connector);

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
                name: name,
                isConnector: isConnector);
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

        /// <summary>
        /// Reads a tri-state that some shape types do not have - flips on a line, the connector flag
        /// on a few others - as false rather than failing over it.
        /// </summary>
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
