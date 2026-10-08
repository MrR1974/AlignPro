using System.Collections.Generic;
using System.Runtime.InteropServices;
using AlignPro.Geometry;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AlignPro.AddIn
{
    /// <summary>How much of an operation actually reached the document.</summary>
    internal sealed class ApplyOutcome
    {
        public ApplyOutcome(int applied, int missing)
        {
            Applied = applied;
            Missing = missing;
        }

        public int Applied { get; }

        /// <summary>Shapes that could not be written - deleted since the snapshot, or locked.</summary>
        public int Missing { get; }
    }

    /// <summary>
    /// Writes new frames and angles back to PowerPoint, and restacks. With <see cref="ShapeCreator"/>,
    /// the only place in the add-in that mutates the document.
    /// </summary>
    internal static class ChangeApplier
    {
        /// <summary>
        /// Applies changes to the slide with the given id. Shapes are resolved by
        /// <see cref="ShapeKey.ShapeId"/> rather than by name, among the slide's shapes and the shapes
        /// inside its groups, and a shape that has since been deleted is skipped rather than throwing.
        /// </summary>
        public static ApplyOutcome Apply(
            PowerPoint.Application app, int slideId, IReadOnlyList<GeometryChange> changes)
        {
            if (changes.Count == 0) return new ApplyOutcome(0, 0);

            // Close PowerPoint's open undo entry first, so the writes below form one entry of their
            // own rather than joining whatever automation has already accumulated. See UndoBoundary.
            UndoBoundary.TryClose(app);

            PowerPoint.Presentation? presentation = null;
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;

            var applied = 0;
            var missing = 0;

            try
            {
                presentation = app.ActivePresentation;
                slides = presentation.Slides;
                slide = slides.FindBySlideID(slideId);
                shapes = slide.Shapes;

                // One pass over the slide to index shapes by id, rather than a lookup per change.
                var byId = IndexById(shapes);

                foreach (var change in changes)
                {
                    if (!byId.TryGetValue(change.Key.ShapeId, out var place))
                    {
                        missing++;
                        continue;
                    }

                    PowerPoint.Shape? shape = null;
                    try
                    {
                        shape = Resolve(shapes, place);
                        ApplyChange(shape, change);
                        applied++;
                    }
                    catch (COMException)
                    {
                        // A locked or otherwise unwritable shape should not abort the rest.
                        missing++;
                    }
                    finally
                    {
                        Com.Release(shape);
                    }
                }

                return new ApplyOutcome(applied, missing);
            }
            finally
            {
                Com.Release(shapes);
                Com.Release(slide);
                Com.Release(slides);
                Com.Release(presentation);
            }
        }

        /// <summary>
        /// Restacks the slide, or the shapes inside one group, into the given order, back to front.
        /// </summary>
        /// <remarks>
        /// <para>
        /// <c>ZOrderPosition</c> is read-only, so the only lever PowerPoint offers is "bring this one
        /// to the front". Applied to each shape in turn from the back of the target order forwards,
        /// that lands on exactly the order asked for: every shape brought to the front lands above the
        /// one before it, so after the last call the stack reads as the list does.
        /// </para>
        /// <para>
        /// Shapes deleted since the snapshot are skipped. The survivors still end up in the right
        /// order relative to one another, because their relative order in the target list is
        /// unaffected by a missing entry.
        /// </para>
        /// </remarks>
        public static ApplyOutcome ApplyOrder(
            PowerPoint.Application app, int slideId, IReadOnlyList<ShapeKey> targetOrder)
        {
            if (targetOrder.Count == 0) return new ApplyOutcome(0, 0);

            // As in Apply: one native undo entry per AlignPro operation, restacking included.
            UndoBoundary.TryClose(app);

            const int msoBringToFront = 0;

            PowerPoint.Presentation? presentation = null;
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;

            var applied = 0;
            var missing = 0;

            try
            {
                presentation = app.ActivePresentation;
                slides = presentation.Slides;
                slide = slides.FindBySlideID(slideId);
                shapes = slide.Shapes;

                // Index by id up front: the set of ids on the slide does not change, so one pass to
                // learn them is enough. The order is either the slide's or one group's, never a mix.
                var present = IndexById(shapes);

                foreach (var key in targetOrder)
                {
                    if (!present.TryGetValue(key.ShapeId, out var place))
                    {
                        missing++;
                        continue;
                    }

                    PowerPoint.Shape? shape = null;
                    try
                    {
                        // Re-scanned per call, never taken from the index: every BringToFront renumbers
                        // the collection it acts in, so a remembered position points at the wrong shape
                        // by the second call. That holds inside a group too - a ribbon run that trusted
                        // GroupItems' positions stacked the wrong shapes. A shape inside a group is
                        // looked for only in its own group, whose place on the slide does not change,
                        // since a BringToFront there restacks within the group and leaves the slide.
                        shape = place.Child == 0
                            ? FindById(shapes, key.ShapeId)
                            : FindChildById(shapes, place.Index, key.ShapeId);
                        if (shape == null)
                        {
                            missing++;
                            continue;
                        }

                        shape.ZOrder((Office.MsoZOrderCmd)msoBringToFront);
                        applied++;
                    }
                    catch (COMException)
                    {
                        // A locked shape should not abort the rest of the stack.
                        missing++;
                    }
                    finally
                    {
                        Com.Release(shape);
                    }
                }

                return new ApplyOutcome(applied, missing);
            }
            finally
            {
                Com.Release(shapes);
                Com.Release(slide);
                Com.Release(slides);
                Com.Release(presentation);
            }
        }

        /// <summary>
        /// Where each shape on the slide is, by id: its index among the slide's shapes, and for a shape
        /// inside a group, its index within that group's <c>GroupItems</c> as well. A shape inside a
        /// group is not in <c>slide.Shapes</c> at all (probe 12), and <c>GroupItems</c> is already
        /// flattened, so one level covers nested groups too (probe 16).
        /// </summary>
        private static Dictionary<int, (int Index, int Child)> IndexById(PowerPoint.Shapes shapes)
        {
            const int msoGroup = 6;

            var byId = new Dictionary<int, (int, int)>(shapes.Count);
            for (var i = 1; i <= shapes.Count; i++)
            {
                PowerPoint.Shape? shape = null;
                PowerPoint.GroupShapes? items = null;
                try
                {
                    shape = shapes[i];
                    byId[shape.Id] = (i, 0);
                    if ((int)shape.Type != msoGroup) continue;

                    items = shape.GroupItems;
                    for (var j = 1; j <= items.Count; j++)
                    {
                        PowerPoint.Shape? child = null;
                        try
                        {
                            child = items[j];
                            byId[child.Id] = (i, j);
                        }
                        finally
                        {
                            Com.Release(child);
                        }
                    }
                }
                finally
                {
                    Com.Release(items);
                    Com.Release(shape);
                }
            }

            return byId;
        }

        /// <summary>The shape at a place <see cref="IndexById"/> recorded. Released by the caller.</summary>
        private static PowerPoint.Shape Resolve(PowerPoint.Shapes shapes, (int Index, int Child) place)
        {
            if (place.Child == 0) return shapes[place.Index];

            PowerPoint.Shape? group = null;
            PowerPoint.GroupShapes? items = null;
            try
            {
                group = shapes[place.Index];
                items = group.GroupItems;
                return items[place.Child];
            }
            finally
            {
                Com.Release(items);
                Com.Release(group);
            }
        }

        /// <summary>
        /// The shape with this id inside the group at <paramref name="groupIndex"/> among the slide's
        /// shapes, or null when it is gone. Released by the caller.
        /// </summary>
        private static PowerPoint.Shape? FindChildById(PowerPoint.Shapes shapes, int groupIndex, int id)
        {
            PowerPoint.Shape? group = null;
            PowerPoint.GroupShapes? items = null;
            try
            {
                group = shapes[groupIndex];
                items = group.GroupItems;
                for (var j = 1; j <= items.Count; j++)
                {
                    var child = items[j];
                    if (child.Id == id) return child;
                    Com.Release(child);
                }

                return null;
            }
            finally
            {
                Com.Release(items);
                Com.Release(group);
            }
        }

        /// <summary>
        /// The shape with this id at slide level or inside any group, or null when it is gone.
        /// Released by the caller.
        /// </summary>
        internal static PowerPoint.Shape? FindAnywhere(PowerPoint.Shapes shapes, int id)
        {
            const int msoGroup = 6;

            var top = FindById(shapes, id);
            if (top != null) return top;

            for (var i = 1; i <= shapes.Count; i++)
            {
                PowerPoint.Shape? shape = null;
                try
                {
                    shape = shapes[i];
                    if ((int)shape.Type != msoGroup) continue;

                    var child = FindChildById(shapes, i, id);
                    if (child != null) return child;
                }
                finally
                {
                    Com.Release(shape);
                }
            }

            return null;
        }

        /// <summary>
        /// The top-level shape with this id, or null when it is gone. Released by the caller.
        /// </summary>
        internal static PowerPoint.Shape? FindById(PowerPoint.Shapes shapes, int id)
        {
            for (var i = 1; i <= shapes.Count; i++)
            {
                PowerPoint.Shape? shape = null;
                try
                {
                    shape = shapes[i];
                    if (shape.Id == id)
                    {
                        var found = shape;
                        shape = null;   // hand ownership to the caller
                        return found;
                    }
                }
                finally
                {
                    Com.Release(shape);
                }
            }

            return null;
        }

        /// <summary>
        /// Writes one change: the angle when the change carries one, then the frame. The order does
        /// not matter to PowerPoint - it rotates about the frame's centre and reports the unrotated
        /// frame whatever the angle - but a change with no angle must leave the angle alone.
        /// </summary>
        internal static void ApplyChange(PowerPoint.Shape shape, GeometryChange change)
        {
            if (change.NewRotation.HasValue &&
                !GeometryChange.SameAngle(shape.Rotation, change.NewRotation.Value))
            {
                shape.Rotation = (float)GeometryChange.NormaliseAngle(change.NewRotation.Value);
            }

            ApplyFrame(shape, change.NewFrame);
        }

        /// <summary>
        /// Writes only the properties that actually differ. Size goes first: setting Width or Height
        /// holds the top-left corner, so position is applied afterwards and wins either way.
        /// </summary>
        internal static void ApplyFrame(PowerPoint.Shape shape, RectD frame)
        {
            const float tolerance = 1e-4f;

            var width = (float)frame.Width;
            var height = (float)frame.Height;
            var left = (float)frame.X;
            var top = (float)frame.Y;

            if (Differs(shape.Width, width, tolerance)) shape.Width = width;
            if (Differs(shape.Height, height, tolerance)) shape.Height = height;
            if (Differs(shape.Left, left, tolerance)) shape.Left = left;
            if (Differs(shape.Top, top, tolerance)) shape.Top = top;
        }

        private static bool Differs(float current, float target, float tolerance) =>
            current > target + tolerance || current < target - tolerance;
    }
}
