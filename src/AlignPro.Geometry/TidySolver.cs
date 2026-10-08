using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace AlignPro.Geometry
{
    /// <summary>What Tidy is asked to do.</summary>
    public sealed class TidyRequest
    {
        /// <summary>The tolerance used when none is given, in points.</summary>
        public const double DefaultTolerance = 3;

        /// <summary>The smallest tolerance accepted: below it, rounding noise would count as intent.</summary>
        public const double MinTolerance = 0.5;

        /// <summary>The largest tolerance accepted: above it, deliberate staggers start to look sloppy.</summary>
        public const double MaxTolerance = 20;

        public TidyRequest(double tolerance = DefaultTolerance, BoundsModel boundsModel = BoundsModel.ShapeFrame)
        {
            Tolerance = tolerance;
            BoundsModel = boundsModel;
        }

        /// <summary>
        /// How far apart, in points, edges or centres can be and still count as meant to line up.
        /// </summary>
        public double Tolerance { get; }

        /// <summary>Which rectangle of each shape is measured, as for the align verbs.</summary>
        public BoundsModel BoundsModel { get; }
    }

    /// <summary>
    /// Finds shapes that are <em>nearly</em> lined up and makes them exact. It fixes a slide built by
    /// eye - the box 2pt left of its neighbours - and never invents a layout that is not already
    /// almost there.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A separate entry point beside <see cref="AlignSolver"/>, because it has no verb and no
    /// reference: it decides for itself what to line up. It returns an ordinary
    /// <see cref="SolveResult"/> of translations, so it gets the Measure setting, rigid groups and one
    /// native undo entry like every other verb. It only ever moves shapes - no change it makes resizes
    /// or turns one - and it leaves connectors out entirely, since they re-route when the shapes they
    /// join move.
    /// </para>
    /// <para>
    /// Each axis is solved on its own, because a move along one cannot disturb the other. On an axis,
    /// each of the three features - left, centre and right, or top, middle and bottom - is swept for
    /// clusters of values within the tolerance of the cluster's lowest, so a chain of values each 2pt
    /// apart never grows into one wide cluster. Alignments the user already has exactly are locked:
    /// Tidy never breaks one to make a near one. A cluster snaps to one of its own members' values,
    /// so at least one shape stays still, and a placeholder - whose position belongs to the layout -
    /// is never moved. A shape moves at most once per axis; where clusters compete for one, the
    /// larger and then tighter wins, and the other is skipped whole rather than partly applied.
    /// </para>
    /// </remarks>
    public static class TidySolver
    {
        /// <summary>
        /// How close two values must be, in points, to count as already equal. Far looser than
        /// <see cref="RectD.Epsilon"/>, because PowerPoint keeps positions to limited precision and a
        /// rotated group's coordinates turn and turn back: a layout Tidy has just made exact reads back a
        /// thousandth of a point off. Anything finer than this is noise, not a fix worth making.
        /// </summary>
        public const double Exact = 0.01;

        private enum Feature { Low, Centre, High }

        private sealed class Cluster
        {
            public Cluster(Feature feature, List<(ShapeSnapshot Shape, double Value)> members)
            {
                Feature = feature;
                Members = members;
            }

            public Feature Feature { get; }

            public List<(ShapeSnapshot Shape, double Value)> Members { get; }

            public double Spread => Members[Members.Count - 1].Value - Members[0].Value;
        }

        public static SolveResult Solve(TidyRequest request, IReadOnlyList<ShapeSnapshot> shapes)
        {
            if (request is null) throw new ArgumentNullException(nameof(request));
            if (shapes is null) throw new ArgumentNullException(nameof(shapes));

            var tolerance = request.Tolerance;
            if (double.IsNaN(tolerance) || tolerance < TidyRequest.MinTolerance || tolerance > TidyRequest.MaxTolerance)
            {
                return SolveResult.Refused(string.Format(
                    CultureInfo.CurrentCulture,
                    "The tolerance must be from {0} to {1} points.",
                    TidyRequest.MinTolerance, TidyRequest.MaxTolerance));
            }

            var connectors = shapes.Count(s => s.IsConnector);
            var candidates = shapes.Where(s => !s.IsConnector).ToList();
            if (candidates.Count < 2)
            {
                return SolveResult.Refused(connectors > 0
                    ? "Select at least two shapes to tidy. Connectors are left out, because they re-route when the shapes they join move."
                    : "Select at least two shapes to tidy.");
            }

            var diagnostics = new List<string>();
            var bounds = AlignSolver.BuildBoundsLookup(request.BoundsModel, candidates, diagnostics);

            var dx = new Dictionary<ShapeKey, double>();
            var dy = new Dictionary<ShapeKey, double>();
            var skipped = 0;

            var alignedX = SolveAxis(true, candidates, bounds, tolerance, dx, ref skipped);
            var alignedY = SolveAxis(false, candidates, bounds, tolerance, dy, ref skipped);

            // Spacing runs on the aligned positions. Rows - shapes sharing an exact top, middle or
            // bottom - are spaced along x; columns along y. Each axis moves only that axis.
            var aligned = candidates.ToDictionary(
                s => s.Key,
                s => bounds[s.Key].Offset(dx.TryGetValue(s.Key, out var x) ? x : 0, dy.TryGetValue(s.Key, out var y) ? y : 0));
            var rows = SpaceAxis(true, candidates, aligned, tolerance, dx, ref skipped);
            var columns = SpaceAxis(false, candidates, aligned, tolerance, dy, ref skipped);

            var changes = new List<GeometryChange>();
            foreach (var shape in candidates)
            {
                dx.TryGetValue(shape.Key, out var x);
                dy.TryGetValue(shape.Key, out var y);
                if (Math.Abs(x) <= Exact && Math.Abs(y) <= Exact) continue;

                changes.Add(new GeometryChange(shape.Key, shape.Frame, shape.Frame.Offset(x, y)));
            }

            // Always speaks: a 2pt fix is invisible, and a silent button that may or may not have done
            // something reads as broken.
            diagnostics.Insert(0, Summary(changes.Count, alignedX + alignedY, rows, columns, tolerance));
            if (skipped > 0)
            {
                diagnostics.Add(skipped == 1
                    ? "1 near-alignment or spacing was skipped, because it would have pulled a shape out of another."
                    : string.Format(CultureInfo.CurrentCulture,
                        "{0} near-alignments or spacings were skipped, because they would have pulled shapes out of others.", skipped));
            }
            if (connectors > 0)
            {
                diagnostics.Add(connectors == 1
                    ? "1 connector was left alone; it follows the shapes it joins."
                    : string.Format(CultureInfo.CurrentCulture,
                        "{0} connectors were left alone; they follow the shapes they join.", connectors));
            }

            return SolveResult.OkWithNotice(changes, diagnostics.ToArray());
        }

        private static string Summary(int moved, int alignments, int rows, int columns, double tolerance)
        {
            if (moved == 0)
            {
                return string.Format(CultureInfo.CurrentCulture,
                    "Nothing to tidy: no shapes are within {0}pt of lining up or of even spacing.", tolerance);
            }

            var parts = new List<string>();
            if (alignments > 0) parts.Add(Count(alignments, "alignment", "alignments"));
            if (rows > 0) parts.Add(Count(rows, "row", "rows"));
            if (columns > 0) parts.Add(Count(columns, "column", "columns"));

            return string.Format(CultureInfo.CurrentCulture,
                "Tidied {0}: {1}.", Count(moved, "shape", "shapes"), string.Join(", ", parts));
        }

        private static string Count(int n, string one, string many) =>
            string.Format(CultureInfo.CurrentCulture, "{0} {1}", n, n == 1 ? one : many);

        /// <summary>
        /// Evens out rows (<paramref name="horizontal"/>) or columns that are nearly evenly spaced,
        /// recording each move in <paramref name="offsets"/> on top of the alignment already there.
        /// </summary>
        /// <remarks>
        /// <para>
        /// A row is three or more shapes sharing an exact top, middle or bottom - whether step 11 made
        /// it so or it already was - that do not overlap along the row. Its gaps and its centre-to-centre
        /// pitches are measured; whichever varies less, if by no more than the tolerance, is made exact
        /// with the outer two shapes held still, as Distribute does. Leading and trailing edges are not
        /// tried: they only differ from centres when sizes differ, and then the gap is what the eye reads.
        /// </para>
        /// <para>
        /// Spacing never breaks an alignment. A shape is moved together with everything exactly lined up
        /// with it along the row's direction - its <em>unit</em> - so in a grid, evening a row slides
        /// whole columns and every column stays a column. A row is skipped if that would move a
        /// placeholder, or move a unit that an earlier row already moved.
        /// </para>
        /// </remarks>
        /// <returns>How many rows or columns were evened.</returns>
        private static int SpaceAxis(
            bool horizontal,
            IReadOnlyList<ShapeSnapshot> shapes,
            Dictionary<ShapeKey, RectD> current,
            double tolerance,
            Dictionary<ShapeKey, double> offsets,
            ref int skipped)
        {
            var units = UnitsAlong(horizontal, shapes, current);
            var unitMoved = new HashSet<int>();
            var evened = 0;

            // Lines run across the other axis: a row shares a y feature and is spaced along x.
            var lines = new List<List<ShapeSnapshot>>();
            var seen = new HashSet<string>();
            foreach (Feature feature in Enum.GetValues(typeof(Feature)))
            {
                var values = shapes
                    .Select(s => (Shape: s, Value: ValueOf(current[s.Key], !horizontal, feature)))
                    .OrderBy(v => v.Value)
                    .ThenBy(v => v.Shape.Key.ShapeId)
                    .ToList();

                foreach (var run in Runs(values, Exact))
                {
                    if (run.Count < 3) continue;
                    var key = string.Join(",", run.Select(m => m.Shape.Key.ShapeId).OrderBy(id => id));
                    if (seen.Add(key)) lines.Add(run.Select(m => m.Shape).ToList());
                }
            }

            foreach (var line in lines.OrderByDescending(l => l.Count))
            {
                var ordered = line.OrderBy(s => Low(current[s.Key], horizontal)).ToList();
                if (!Separate(ordered, current, horizontal)) continue;

                var targets = EvenTargets(ordered, current, horizontal, tolerance);
                if (targets == null) continue;

                // The move each unit needs, and whether every one of them is allowed.
                var deltas = new Dictionary<int, double>();
                var allowed = true;
                for (var i = 0; i < ordered.Count && allowed; i++)
                {
                    var delta = targets[i] - Low(current[ordered[i].Key], horizontal);
                    var unit = units[ordered[i].Key];

                    if (deltas.TryGetValue(unit, out var other))
                    {
                        allowed = Math.Abs(other - delta) <= Exact;
                        continue;
                    }
                    deltas[unit] = delta;

                    if (Math.Abs(delta) <= Exact) continue;
                    if (unitMoved.Contains(unit) || shapes.Any(s => units[s.Key] == unit && s.IsPlaceholder)) allowed = false;
                }

                if (!allowed)
                {
                    skipped++;
                    continue;
                }

                foreach (var pair in deltas)
                {
                    if (Math.Abs(pair.Value) <= Exact) continue;

                    unitMoved.Add(pair.Key);
                    foreach (var shape in shapes.Where(s => units[s.Key] == pair.Key))
                    {
                        offsets[shape.Key] = (offsets.TryGetValue(shape.Key, out var before) ? before : 0) + pair.Value;
                        current[shape.Key] = horizontal
                            ? current[shape.Key].Offset(pair.Value, 0)
                            : current[shape.Key].Offset(0, pair.Value);
                    }
                }
                evened++;
            }

            return evened;
        }

        /// <summary>
        /// Where each shape's leading edge should go for the line to be evenly spaced, or null when it
        /// already is, or is too uneven to have been meant as even.
        /// </summary>
        private static double[]? EvenTargets(
            List<ShapeSnapshot> ordered, Dictionary<ShapeKey, RectD> current, bool horizontal, double tolerance)
        {
            var rects = ordered.Select(s => current[s.Key]).ToList();
            var n = rects.Count;

            var gaps = new double[n - 1];
            var pitches = new double[n - 1];
            for (var i = 0; i < n - 1; i++)
            {
                gaps[i] = Low(rects[i + 1], horizontal) - High(rects[i], horizontal);
                pitches[i] = Mid(rects[i + 1], horizontal) - Mid(rects[i], horizontal);
            }

            var gapVariation = gaps.Max() - gaps.Min();
            var pitchVariation = pitches.Max() - pitches.Min();
            var useGaps = gapVariation <= pitchVariation;
            var variation = useGaps ? gapVariation : pitchVariation;

            if (variation <= Exact || variation > tolerance) return null;

            var targets = new double[n];
            targets[0] = Low(rects[0], horizontal);
            targets[n - 1] = Low(rects[n - 1], horizontal);

            if (useGaps)
            {
                var gap = gaps.Sum() / (n - 1);
                var cursor = High(rects[0], horizontal);
                for (var i = 1; i < n - 1; i++)
                {
                    targets[i] = cursor + gap;
                    cursor = targets[i] + Size(rects[i], horizontal);
                }
            }
            else
            {
                var first = Mid(rects[0], horizontal);
                var pitch = (Mid(rects[n - 1], horizontal) - first) / (n - 1);
                for (var i = 1; i < n - 1; i++)
                {
                    targets[i] = first + pitch * i - Size(rects[i], horizontal) / 2;
                }
            }

            return targets;
        }

        /// <summary>True when no shape in the line overlaps the next along it.</summary>
        private static bool Separate(List<ShapeSnapshot> ordered, Dictionary<ShapeKey, RectD> current, bool horizontal)
        {
            for (var i = 0; i < ordered.Count - 1; i++)
            {
                if (Low(current[ordered[i + 1].Key], horizontal) < High(current[ordered[i].Key], horizontal) - Exact) return false;
            }

            return true;
        }

        /// <summary>
        /// Groups shapes that are exactly lined up along <paramref name="horizontal"/> - on any of the
        /// three features, and transitively - so that moving one moves them all, keeping every
        /// alignment.
        /// </summary>
        private static Dictionary<ShapeKey, int> UnitsAlong(
            bool horizontal, IReadOnlyList<ShapeSnapshot> shapes, Dictionary<ShapeKey, RectD> current)
        {
            var parent = shapes.ToDictionary(s => s.Key, s => s.Key);

            ShapeKey Find(ShapeKey key)
            {
                while (!parent[key].Equals(key)) key = parent[key];
                return key;
            }

            foreach (Feature feature in Enum.GetValues(typeof(Feature)))
            {
                var values = shapes
                    .Select(s => (Shape: s, Value: ValueOf(current[s.Key], horizontal, feature)))
                    .OrderBy(v => v.Value)
                    .ToList();
                foreach (var run in Runs(values, Exact))
                {
                    for (var i = 1; i < run.Count; i++) parent[Find(run[i].Shape.Key)] = Find(run[0].Shape.Key);
                }
            }

            var ids = new Dictionary<ShapeKey, int>();
            var units = new Dictionary<ShapeKey, int>();
            foreach (var shape in shapes)
            {
                var root = Find(shape.Key);
                if (!ids.TryGetValue(root, out var id))
                {
                    id = ids.Count;
                    ids[root] = id;
                }
                units[shape.Key] = id;
            }

            return units;
        }

        private static double Low(RectD rect, bool horizontal) => horizontal ? rect.Left : rect.Top;

        private static double High(RectD rect, bool horizontal) => horizontal ? rect.Right : rect.Bottom;

        private static double Mid(RectD rect, bool horizontal) => horizontal ? rect.CentreX : rect.CentreY;

        private static double Size(RectD rect, bool horizontal) => horizontal ? rect.Width : rect.Height;

        /// <summary>
        /// Settles one axis, recording each moved shape's offset in <paramref name="offsets"/>.
        /// </summary>
        /// <returns>How many alignments were made.</returns>
        private static int SolveAxis(
            bool horizontal,
            IReadOnlyList<ShapeSnapshot> shapes,
            IReadOnlyDictionary<ShapeKey, RectD> bounds,
            double tolerance,
            Dictionary<ShapeKey, double> offsets,
            ref int skipped)
        {
            // Shapes that may not move on this axis: placeholders always, and anything already in an
            // exact alignment. Exact on any feature locks the shape for the whole axis, because moving
            // it would break that alignment whichever feature the move was for.
            var locked = new HashSet<ShapeKey>(shapes.Where(s => s.IsPlaceholder).Select(s => s.Key));

            var clusters = new List<Cluster>();
            foreach (Feature feature in Enum.GetValues(typeof(Feature)))
            {
                var values = shapes
                    .Select(s => (Shape: s, Value: ValueOf(bounds[s.Key], horizontal, feature)))
                    .OrderBy(v => v.Value)
                    .ThenBy(v => v.Shape.Key.ShapeId)
                    .ToList();

                foreach (var run in Runs(values, Exact))
                {
                    if (run.Count >= 2) foreach (var member in run) locked.Add(member.Shape.Key);
                }

                foreach (var run in Runs(values, tolerance))
                {
                    if (run.Count >= 2) clusters.Add(new Cluster(feature, run));
                }
            }

            var moved = new HashSet<ShapeKey>();
            var made = 0;

            var ordered = clusters
                .OrderByDescending(c => c.Members.Count)
                .ThenBy(c => c.Spread)
                .ThenBy(c => c.Feature)
                .ThenBy(c => c.Members[0].Value);

            foreach (var found in ordered)
            {
                // Judged where its shapes are now, after the clusters already accepted. Snapping left
                // edges of equal widths also lines up their centres and rights, and those clusters
                // must then read as done, not as conflicts.
                var cluster = new Cluster(found.Feature, found.Members
                    .Select(m => (m.Shape, Value: m.Value + (offsets.TryGetValue(m.Shape.Key, out var o) ? o : 0)))
                    .OrderBy(m => m.Value)
                    .ToList());

                if (!TryChooseTarget(cluster, out var target))
                {
                    skipped++;
                    continue;
                }

                // Already exact: nothing to do, and nothing to claim.
                if (cluster.Members.All(m => AtTarget(m.Value, target))) continue;

                var blocked = cluster.Members.Any(m =>
                    !AtTarget(m.Value, target) && (locked.Contains(m.Shape.Key) || moved.Contains(m.Shape.Key)));
                if (blocked)
                {
                    skipped++;
                    continue;
                }

                foreach (var member in cluster.Members)
                {
                    moved.Add(member.Shape.Key);
                    var offset = target - member.Value;
                    if (Math.Abs(offset) <= Exact) continue;
                    offsets[member.Shape.Key] = (offsets.TryGetValue(member.Shape.Key, out var before) ? before : 0) + offset;
                }
                made++;
            }

            return made;
        }

        /// <summary>
        /// The value a cluster snaps to: one of its members', never a computed one, so at least one
        /// shape stays still. The member that moves the others least - the median - with ties to the
        /// lower shape id. A placeholder's value wins outright; two placeholders that disagree leave
        /// nothing to choose, and the cluster is skipped.
        /// </summary>
        private static bool TryChooseTarget(Cluster cluster, out double target)
        {
            var placeholders = cluster.Members.Where(m => m.Shape.IsPlaceholder).ToList();
            if (placeholders.Count > 0)
            {
                target = placeholders[0].Value;
                var first = target;
                return placeholders.All(p => AtTarget(p.Value, first));
            }

            var best = cluster.Members
                .Select(candidate => (
                    candidate.Value,
                    candidate.Shape.Key.ShapeId,
                    Cost: cluster.Members.Sum(m => Math.Abs(m.Value - candidate.Value))))
                .OrderBy(c => c.Cost)
                .ThenBy(c => c.ShapeId)
                .First();

            target = best.Value;
            return true;
        }

        /// <summary>
        /// Sorted values swept into runs, each value joining while it is within
        /// <paramref name="width"/> of the run's lowest. So no run spans more than the width.
        /// </summary>
        private static IEnumerable<List<(ShapeSnapshot Shape, double Value)>> Runs(
            List<(ShapeSnapshot Shape, double Value)> sorted, double width)
        {
            var run = new List<(ShapeSnapshot Shape, double Value)>();
            foreach (var entry in sorted)
            {
                if (run.Count > 0 && entry.Value - run[0].Value > width)
                {
                    yield return run;
                    run = new List<(ShapeSnapshot Shape, double Value)>();
                }
                run.Add(entry);
            }

            if (run.Count > 0) yield return run;
        }

        private static double ValueOf(RectD rect, bool horizontal, Feature feature)
        {
            switch (feature)
            {
                case Feature.Low: return horizontal ? rect.Left : rect.Top;
                case Feature.High: return horizontal ? rect.Right : rect.Bottom;
                default: return horizontal ? rect.CentreX : rect.CentreY;
            }
        }

        private static bool AtTarget(double value, double target) => Math.Abs(value - target) <= Exact;
    }
}
