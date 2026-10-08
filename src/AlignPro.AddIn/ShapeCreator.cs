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
    /// Makes the shapes behind a duplicate. The only code in the add-in that adds shapes to a slide.
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
        /// <remarks>
        /// The originals are the shapes the reader held, and each copy is kept as <c>Duplicate</c>
        /// hands it back, to be placed and then selected - nothing is looked up twice.
        /// </remarks>
        public static CreateOutcome Create(
            PowerPoint.Application app,
            int slideId,
            ShapeIndex index,
            IReadOnlyList<ShapeKey> originals,
            IReadOnlyList<DuplicateCopy> copies)
        {
            // One native undo entry for the whole operation, as for every other verb.
            UndoBoundary.TryClose(app);

            var created = new List<ShapeKey>();
            var missing = 0;

            {
                // Where each original sits in the stack, by ZOrderPosition - numbered across the whole
                // slide, so it orders shapes inside a group as well as shapes on the slide. A shape
                // that is gone sorts last of all, and is counted.
                var stacking = new Dictionary<ShapeKey, int>(originals.Count);
                foreach (var original in originals)
                {
                    var shape = index.Find(original.ShapeId);
                    stacking[original] = shape?.ZOrderPosition ?? int.MaxValue;
                    if (shape == null) missing++;
                }

                var backToFront = new List<ShapeKey>(originals);
                backToFront.Sort((a, b) => stacking[a].CompareTo(stacking[b]));

                foreach (var copy in copies)
                {
                    // Made back to front for stacking, but recorded in selection order so the
                    // selection that follows - and so the next verb's anchor - reads as the user
                    // would expect.
                    var made = new Dictionary<ShapeKey, ShapeKey>();

                    foreach (var original in backToFront)
                    {
                        var placement = PlacementFor(copy, original);
                        var source = index.Find(original.ShapeId);
                        if (placement == null || source == null) continue;

                        var madeKey = DuplicateOne(source, index, slideId, original, placement);
                        if (madeKey.HasValue) made[original] = madeKey.Value;
                    }

                    foreach (var original in originals)
                    {
                        if (made.TryGetValue(original, out var key)) created.Add(key);
                    }
                }

                SelectAll(index, originals, created);
            }

            return new CreateOutcome(created, missing);
        }

        /// <summary>
        /// Makes one copy and places it. A shape inside a group is copied into the same group, as
        /// PowerPoint's own Ctrl+D does (probes 20 to 22). The copy joins the index, so it can be
        /// selected afterwards without being looked for.
        /// </summary>
        private static ShapeKey? DuplicateOne(
            PowerPoint.Shape source, ShapeIndex index, int slideId, ShapeKey original, ShapePlacement placement)
        {
            PowerPoint.ShapeRange? duplicated = null;
            try
            {
                duplicated = source.Duplicate();
                var copy = duplicated[1];
                var id = copy.Id;
                index.Keep(copy, id);

                // Duplicate lands the copy at a small offset from the original, the same size. The
                // placement is absolute, so its angle and position are simply written - without
                // reading the copy first, which would wait on PowerPoint redrawing the slide.
                copy.Rotation = (float)placement.Rotation;
                copy.Left = (float)placement.Frame.X;
                copy.Top = (float)placement.Frame.Y;
                return new ShapeKey(slideId, id);
            }
            catch (COMException ex)
            {
                Diagnostics.Log("Duplicate of " + original + " failed: " + ex.Message);
                return null;
            }
            finally
            {
                Com.Release(duplicated);
            }
        }

        /// <summary>
        /// Leaves the originals and every copy selected, originals first - so the copies can go
        /// straight into the next verb, and the last copy made is the anchor.
        /// </summary>
        private static void SelectAll(ShapeIndex index, IReadOnlyList<ShapeKey> originals, IReadOnlyList<ShapeKey> created)
        {
            var replace = true;
            foreach (var key in Concat(originals, created))
            {
                var shape = index.Find(key.ShapeId);
                if (shape == null) continue;

                try
                {
                    shape.Select(replace ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse);
                    replace = false;
                }
                catch (COMException ex)
                {
                    // Selection is a convenience; the shapes exist either way.
                    Diagnostics.Log("Selecting " + key + " after duplicate failed: " + ex.Message);
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
    }
}
