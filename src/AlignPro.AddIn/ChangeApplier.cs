using System.Collections.Generic;
using System.Runtime.InteropServices;
using AlignPro.Geometry;
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
                        Release(shape);
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
                        Release(shape);
                    }
                }

                return new ApplyOutcome(applied, missing);
            }
            finally
            {
                Release(shapes);
                Release(slide);
                Release(slides);
                Release(presentation);
            }
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

        private static void Release(object? comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
    }
}
