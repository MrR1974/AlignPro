using System;
using System.Collections.Generic;

namespace AlignPro.Geometry
{
    /// <summary>What to do with the selection's stacking order. One verb per ribbon button.</summary>
    public enum OrderVerb
    {
        /// <summary>The first shape selected ends on top, the last selected at the bottom.</summary>
        StackFirstOnTop,

        /// <summary>The mirror image: the first shape selected ends at the bottom.</summary>
        StackFirstOnBottom
    }

    /// <summary>
    /// The outcome of an ordering solve. Mirrors <see cref="SolveResult"/>: a refusal is an ordinary
    /// result carrying a reason, not an exception.
    /// </summary>
    public sealed class OrderResult
    {
        private OrderResult(ZOrderChange? change, IReadOnlyList<string> diagnostics, bool succeeded)
        {
            Change = change;
            Diagnostics = diagnostics;
            Succeeded = succeeded;
        }

        /// <summary>The ordering to apply. Null when the solve was refused.</summary>
        public ZOrderChange? Change { get; }

        /// <summary>Human-readable notes: why nothing happened.</summary>
        public IReadOnlyList<string> Diagnostics { get; }

        public bool Succeeded { get; }

        public static OrderResult Ok(ZOrderChange change) =>
            new OrderResult(change, Array.Empty<string>(), true);

        public static OrderResult Refused(string reason) =>
            new OrderResult(null, new[] { reason }, false);
    }

    /// <summary>
    /// Works out the slide's new stacking order. Pure: no Office interop, no mutation of its inputs.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Kept apart from <see cref="AlignSolver"/> on purpose. That solver is defined by emitting frame
    /// coordinates and nothing else, and ordering emits no geometry at all, so folding it in would
    /// mean a verb its own contract could not describe.
    /// </para>
    /// <para>
    /// The ordering is <em>in place</em>: the selected shapes are redistributed across the z-positions
    /// they already occupy between them, so anything not selected keeps the layer it was on. Lifting
    /// the whole selection to the front would have been one call per shape and no arithmetic, but it
    /// silently drags shapes over the top of things the user never asked to disturb.
    /// </para>
    /// </remarks>
    public static class ZOrderSolver
    {
        /// <summary>
        /// Rewrites the selection's slots in the slide's stacking order, leaving every unselected
        /// shape exactly where it was.
        /// </summary>
        /// <param name="slideOrder">
        /// Every top-level shape on the slide, back to front, as PowerPoint currently has them.
        /// </param>
        /// <param name="selection">The selected shapes, in the order they were selected.</param>
        /// <param name="verb">Which end of the selection ends up on top.</param>
        public static OrderResult Solve(
            IReadOnlyList<ShapeKey> slideOrder, IReadOnlyList<ShapeKey> selection, OrderVerb verb)
        {
            if (slideOrder is null) throw new ArgumentNullException(nameof(slideOrder));
            if (selection is null) throw new ArgumentNullException(nameof(selection));

            if (selection.Count < 2) return OrderResult.Refused("Select at least two shapes to reorder.");

            // Where each selected shape currently sits. A shape that is not in the slide's own
            // collection is inside a group, and its z-position is measured within that group rather
            // than against the slide - so there is no honest answer, and guessing one would quietly
            // restack the wrong things.
            var slots = new List<int>(selection.Count);
            var selected = new HashSet<ShapeKey>();

            foreach (var key in selection)
            {
                if (!selected.Add(key)) continue;

                var index = IndexOf(slideOrder, key);
                if (index < 0)
                {
                    return OrderResult.Refused(
                        "Ordering works on shapes that sit directly on the slide. Select the group " +
                        "itself rather than a shape inside it.");
                }

                slots.Add(index);
            }

            if (slots.Count < 2) return OrderResult.Refused("Select at least two shapes to reorder.");

            // The slots keep their positions; only which shape occupies each one changes. Sorting them
            // detaches the slots from the selection order, which is the whole point - the lowest slot
            // gets whichever shape belongs at the back.
            slots.Sort();

            var target = new ShapeKey[slideOrder.Count];
            for (var i = 0; i < slideOrder.Count; i++) target[i] = slideOrder[i];

            // Slots run back to front, so the shape that belongs on top is filled in last. "First
            // selected on top" therefore walks the selection backwards.
            for (var i = 0; i < slots.Count; i++)
            {
                var pick = verb == OrderVerb.StackFirstOnTop ? slots.Count - 1 - i : i;
                target[slots[i]] = DistinctAt(selection, pick);
            }

            return OrderResult.Ok(new ZOrderChange(slideOrder, target));
        }

        private static int IndexOf(IReadOnlyList<ShapeKey> order, ShapeKey key)
        {
            for (var i = 0; i < order.Count; i++)
            {
                if (order[i] == key) return i;
            }

            return -1;
        }

        /// <summary>
        /// The nth distinct key in selection order. PowerPoint should not hand us the same shape
        /// twice, but a duplicate would otherwise shift every shape after it by one slot and scramble
        /// the result rather than merely repeating it.
        /// </summary>
        private static ShapeKey DistinctAt(IReadOnlyList<ShapeKey> selection, int n)
        {
            var seen = new HashSet<ShapeKey>();
            foreach (var key in selection)
            {
                if (!seen.Add(key)) continue;
                if (seen.Count - 1 == n) return key;
            }

            throw new ArgumentOutOfRangeException(nameof(n), n, "Fewer distinct shapes than slots.");
        }
    }
}
