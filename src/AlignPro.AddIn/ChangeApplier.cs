using System.Collections.Generic;
using System.Runtime.InteropServices;
using AlignPro.Geometry;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AlignPro.AddIn
{
    /// <summary>How much of a transaction actually reached the document.</summary>
    internal sealed class ApplyOutcome
    {
        public ApplyOutcome(int applied, int missing)
        {
            Applied = applied;
            Missing = missing;
        }

        public int Applied { get; }

        /// <summary>Shapes that no longer exist - deleted since the snapshot, or since the undo.</summary>
        public int Missing { get; }
    }

    /// <summary>
    /// Writes new frames back to PowerPoint. The only place in the add-in that mutates the document.
    /// </summary>
    internal static class ChangeApplier
    {
        /// <summary>
        /// Applies changes to the slide with the given id. Shapes are resolved by
        /// <see cref="ShapeKey.ShapeId"/> rather than by name, and a shape that has since been deleted
        /// is skipped rather than throwing - which is what makes undo safe after an edit.
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
                var byId = new Dictionary<int, int>(shapes.Count);
                for (var i = 1; i <= shapes.Count; i++)
                {
                    PowerPoint.Shape? shape = null;
                    try
                    {
                        shape = shapes[i];
                        byId[shape.Id] = i;
                    }
                    finally
                    {
                        Com.Release(shape);
                    }
                }

                foreach (var change in changes)
                {
                    if (!byId.TryGetValue(change.Key.ShapeId, out var index))
                    {
                        missing++;
                        continue;
                    }

                    PowerPoint.Shape? shape = null;
                    try
                    {
                        shape = shapes[index];
                        ApplyFrame(shape, change.NewFrame);
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
        /// Restacks the slide into the given order, back to front.
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

                // Index by id up front. Positions shift with every BringToFront, so the shape has to
                // be fetched by identity each time rather than by a remembered index - but the set of
                // ids on the slide does not change, so one pass to learn them is enough.
                var present = new HashSet<int>();
                for (var i = 1; i <= shapes.Count; i++)
                {
                    PowerPoint.Shape? shape = null;
                    try
                    {
                        shape = shapes[i];
                        present.Add(shape.Id);
                    }
                    finally
                    {
                        Com.Release(shape);
                    }
                }

                foreach (var key in targetOrder)
                {
                    if (!present.Contains(key.ShapeId))
                    {
                        missing++;
                        continue;
                    }

                    PowerPoint.Shape? shape = null;
                    try
                    {
                        // Re-scanned per shape rather than indexed once: every BringToFront renumbers
                        // the collection, so a remembered index would point at the wrong shape by the
                        // second call. The Shapes collection has no by-id accessor to use instead.
                        shape = FindById(shapes, key.ShapeId);
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
        /// The shape with this id, or null when it is gone. Released by the caller.
        /// </summary>
        private static PowerPoint.Shape? FindById(PowerPoint.Shapes shapes, int id)
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
        /// Writes only the properties that actually differ. Size goes first: setting Width or Height
        /// holds the top-left corner, so position is applied afterwards and wins either way.
        /// </summary>
        private static void ApplyFrame(PowerPoint.Shape shape, RectD frame)
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
