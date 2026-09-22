using System.Collections.Generic;
using System.Runtime.InteropServices;
using AlignPro.Geometry;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AlignPro.AddIn
{
    /// <summary>What a creation pass made, and what it could not.</summary>
    internal sealed class CreateOutcome
    {
        public CreateOutcome(IReadOnlyList<ShapeKey> created, int missing)
        {
            Created = created;
            Missing = missing;
        }

        /// <summary>Every shape made, copy by copy, in selection order within each copy.</summary>
        public IReadOnlyList<ShapeKey> Created { get; }

        /// <summary>Originals that no longer exist, so their copies could not be made.</summary>
        public int Missing { get; }
    }

    /// <summary>
    /// Makes and removes the shapes behind a duplicate. The only code in the add-in that adds shapes
    /// to a slide or deletes them.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Each original is duplicated on its own with <c>Shape.Duplicate</c> rather than the selection
    /// all at once with <c>ShapeRange.Duplicate</c>. The copies must be matched to their placements,
    /// and one call per shape makes that match certain instead of resting on the order a range
    /// happens to come back in. A group duplicates as a group either way.
    /// </para>
    /// <para>
    /// The originals are duplicated back to front, so each copy lands above the one made before it
    /// and the copies keep the originals' stacking among themselves. See
    /// <c>docs/object-model-findings.md</c>, probe 8.
    /// </para>
    /// </remarks>
    internal static class ShapeCreator
    {
        /// <summary>
        /// Makes every copy in the recipe and leaves the originals and the copies selected.
        /// </summary>
        /// <param name="originals">The selection, in the order it was selected.</param>
        public static CreateOutcome Create(
            PowerPoint.Application app,
            int slideId,
            IReadOnlyList<ShapeKey> originals,
            IReadOnlyList<DuplicateCopy> copies)
        {
            // One native undo entry for the whole operation, as for every other verb.
            UndoBoundary.TryClose(app);

            PowerPoint.Presentation? presentation = null;
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;

            var created = new List<ShapeKey>();
            var missing = 0;

            try
            {
                presentation = app.ActivePresentation;
                slides = presentation.Slides;
                slide = slides.FindBySlideID(slideId);
                shapes = slide.Shapes;

                var stacking = StackingOrder(shapes);
                var backToFront = new List<ShapeKey>(originals);
                backToFront.Sort((a, b) => IndexIn(stacking, a).CompareTo(IndexIn(stacking, b)));

                foreach (var original in originals)
                {
                    if (IndexIn(stacking, original) == int.MaxValue) missing++;
                }

                foreach (var copy in copies)
                {
                    // Made back to front for stacking, but recorded in selection order so the
                    // selection that follows - and so the next verb's anchor - reads as the user
                    // would expect.
                    var made = new Dictionary<ShapeKey, ShapeKey>();

                    foreach (var original in backToFront)
                    {
                        var placement = PlacementFor(copy, original);
                        if (placement == null) continue;

                        var madeKey = DuplicateOne(shapes, slideId, original, placement);
                        if (madeKey.HasValue) made[original] = madeKey.Value;
                    }

                    foreach (var original in originals)
                    {
                        if (made.TryGetValue(original, out var key)) created.Add(key);
                    }
                }

                SelectAll(shapes, originals, created);
                return new CreateOutcome(created, missing);
            }
            finally
            {
                Com.Release(shapes);
                Com.Release(slide);
                Com.Release(slides);
                Com.Release(presentation);
            }
        }

        /// <summary>Deletes shapes a duplicate made. Shapes already gone are counted, not fatal.</summary>
        public static ApplyOutcome Remove(PowerPoint.Application app, int slideId, IReadOnlyList<ShapeKey> created)
        {
            if (created.Count == 0) return new ApplyOutcome(0, 0);

            UndoBoundary.TryClose(app);

            PowerPoint.Presentation? presentation = null;
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;

            var removed = 0;
            var missing = 0;

            try
            {
                presentation = app.ActivePresentation;
                slides = presentation.Slides;
                slide = slides.FindBySlideID(slideId);
                shapes = slide.Shapes;

                foreach (var key in created)
                {
                    PowerPoint.Shape? shape = null;
                    try
                    {
                        // Looked up each time: every deletion renumbers the collection.
                        shape = ChangeApplier.FindById(shapes, key.ShapeId);
                        if (shape == null)
                        {
                            missing++;
                            continue;
                        }

                        shape.Delete();
                        removed++;
                    }
                    catch (COMException)
                    {
                        missing++;
                    }
                    finally
                    {
                        Com.Release(shape);
                    }
                }

                return new ApplyOutcome(removed, missing);
            }
            finally
            {
                Com.Release(shapes);
                Com.Release(slide);
                Com.Release(slides);
                Com.Release(presentation);
            }
        }

        private static ShapeKey? DuplicateOne(
            PowerPoint.Shapes shapes, int slideId, ShapeKey original, ShapePlacement placement)
        {
            PowerPoint.Shape? source = null;
            PowerPoint.ShapeRange? duplicated = null;
            PowerPoint.Shape? copy = null;
            try
            {
                source = ChangeApplier.FindById(shapes, original.ShapeId);
                if (source == null) return null;

                duplicated = source.Duplicate();
                copy = duplicated[1];

                // Duplicate lands the copy at a small offset from the original. The placement is
                // absolute, so writing it simply overrides that offset - nothing to undo first.
                if (!GeometryChange.SameAngle(copy.Rotation, placement.Rotation))
                {
                    copy.Rotation = (float)placement.Rotation;
                }

                ChangeApplier.ApplyFrame(copy, placement.Frame);
                return new ShapeKey(slideId, copy.Id);
            }
            catch (COMException ex)
            {
                Diagnostics.Log("Duplicate of " + original + " failed: " + ex.Message);
                return null;
            }
            finally
            {
                Com.Release(copy);
                Com.Release(duplicated);
                Com.Release(source);
            }
        }

        /// <summary>
        /// Leaves the originals and every copy selected, originals first - so the copies can go
        /// straight into the next verb, and the last copy made is the anchor.
        /// </summary>
        private static void SelectAll(
            PowerPoint.Shapes shapes, IReadOnlyList<ShapeKey> originals, IReadOnlyList<ShapeKey> created)
        {
            var replace = true;
            foreach (var key in Concat(originals, created))
            {
                PowerPoint.Shape? shape = null;
                try
                {
                    shape = ChangeApplier.FindById(shapes, key.ShapeId);
                    if (shape == null) continue;

                    shape.Select(replace ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse);
                    replace = false;
                }
                catch (COMException ex)
                {
                    // Selection is a convenience; the shapes exist either way.
                    Diagnostics.Log("Selecting " + key + " after duplicate failed: " + ex.Message);
                }
                finally
                {
                    Com.Release(shape);
                }
            }
        }

        private static IEnumerable<ShapeKey> Concat(IReadOnlyList<ShapeKey> first, IReadOnlyList<ShapeKey> second)
        {
            foreach (var key in first) yield return key;
            foreach (var key in second) yield return key;
        }

        private static ShapePlacement? PlacementFor(DuplicateCopy copy, ShapeKey original)
        {
            foreach (var placement in copy.Placements)
            {
                if (placement.Source == original) return placement;
            }

            return null;
        }

        /// <summary>Shape ids back to front, as the collection holds them.</summary>
        private static List<int> StackingOrder(PowerPoint.Shapes shapes)
        {
            var order = new List<int>(shapes.Count);
            for (var i = 1; i <= shapes.Count; i++)
            {
                PowerPoint.Shape? shape = null;
                try
                {
                    shape = shapes[i];
                    order.Add(shape.Id);
                }
                finally
                {
                    Com.Release(shape);
                }
            }

            return order;
        }

        /// <summary>Where a shape sits in the stack, or last of all when it is not on the slide.</summary>
        private static int IndexIn(List<int> stacking, ShapeKey key)
        {
            var index = stacking.IndexOf(key.ShapeId);
            return index < 0 ? int.MaxValue : index;
        }
    }
}
