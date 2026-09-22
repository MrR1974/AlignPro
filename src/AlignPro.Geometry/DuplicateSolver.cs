using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace AlignPro.Geometry
{
    /// <summary>The point each duplicate step turns about.</summary>
    public enum DuplicatePivot
    {
        /// <summary>
        /// Each shape turns about its own centre, so the angle spins shapes in place and only the
        /// X and Y offsets move them.
        /// </summary>
        OwnCentre,

        /// <summary>The centre of the whole selection as it looks - its visual bounds.</summary>
        SelectionCentre,

        /// <summary>The centre of the anchor, the shape selected last.</summary>
        AnchorCentre,

        /// <summary>The centre of the slide.</summary>
        SlideCentre
    }

    /// <summary>What to duplicate, how far each step moves and turns, and how many copies to make.</summary>
    public sealed class DuplicateRequest
    {
        /// <summary>
        /// More copies than this is almost certainly a typo, and each one is a real shape on the
        /// slide that then has to be undone.
        /// </summary>
        public const int MaxCopies = 100;

        public DuplicateRequest(
            double offsetX = 0,
            double offsetY = 0,
            double angle = 0,
            int copies = 1,
            DuplicatePivot pivot = DuplicatePivot.OwnCentre,
            bool rotateShapes = true,
            ShapeKey? anchor = null)
        {
            OffsetX = offsetX;
            OffsetY = offsetY;
            Angle = angle;
            Copies = copies;
            Pivot = pivot;
            RotateShapes = rotateShapes;
            Anchor = anchor;
        }

        /// <summary>How far each step moves right, in points. Negative moves left.</summary>
        public double OffsetX { get; }

        /// <summary>How far each step moves down, in points. Negative moves up.</summary>
        public double OffsetY { get; }

        /// <summary>How far each step turns, in degrees clockwise - PowerPoint's own sense.</summary>
        public double Angle { get; }

        public int Copies { get; }

        public DuplicatePivot Pivot { get; }

        /// <summary>
        /// True turns each copy with the step. False keeps copies at the original's angle, so only
        /// their positions follow the rotation - a ring of upright labels rather than spokes.
        /// </summary>
        public bool RotateShapes { get; }

        /// <summary>Required by <see cref="DuplicatePivot.AnchorCentre"/>.</summary>
        public ShapeKey? Anchor { get; }
    }

    /// <summary>Where one copy of one original shape goes.</summary>
    public sealed class ShapePlacement
    {
        public ShapePlacement(ShapeKey source, RectD frame, double rotation)
        {
            Source = source;
            Frame = frame;
            Rotation = rotation;
        }

        /// <summary>The original this copy is made from.</summary>
        public ShapeKey Source { get; }

        /// <summary>The copy's frame, the same size as the original's.</summary>
        public RectD Frame { get; }

        /// <summary>The copy's clockwise angle in degrees, in [0, 360).</summary>
        public double Rotation { get; }

        public override string ToString() =>
            string.Format(CultureInfo.InvariantCulture, "from {0}: {1} rot={2:F1}", Source, Frame, Rotation);
    }

    /// <summary>One copy of the whole selection.</summary>
    public sealed class DuplicateCopy
    {
        public DuplicateCopy(int index, IReadOnlyList<ShapePlacement> placements)
        {
            if (index < 1) throw new ArgumentOutOfRangeException(nameof(index), index, "Copies count from one.");
            Index = index;
            Placements = placements ?? throw new ArgumentNullException(nameof(placements));
        }

        /// <summary>Which copy this is, from one: copy k has had the step applied k times.</summary>
        public int Index { get; }

        /// <summary>One placement per original, in the order the originals were selected.</summary>
        public IReadOnlyList<ShapePlacement> Placements { get; }
    }

    /// <summary>
    /// The outcome of a duplicate solve. Mirrors <see cref="SolveResult"/>: a refusal is an ordinary
    /// result with a reason, not an exception.
    /// </summary>
    public sealed class DuplicateResult
    {
        private DuplicateResult(
            IReadOnlyList<DuplicateCopy> copies, IReadOnlyList<string> diagnostics, bool succeeded, bool notable)
        {
            Copies = copies;
            Diagnostics = diagnostics;
            Succeeded = succeeded;
            Notable = notable;
        }

        public IReadOnlyList<DuplicateCopy> Copies { get; }

        public IReadOnlyList<string> Diagnostics { get; }

        public bool Succeeded { get; }

        /// <summary>True when the user needs telling something, as for <see cref="SolveResult.Notable"/>.</summary>
        public bool Notable { get; }

        public static DuplicateResult Ok(IReadOnlyList<DuplicateCopy> copies, params string[] diagnostics) =>
            new DuplicateResult(copies, diagnostics, true, notable: false);

        public static DuplicateResult OkWithNotice(IReadOnlyList<DuplicateCopy> copies, params string[] diagnostics) =>
            new DuplicateResult(copies, diagnostics, true, notable: true);

        public static DuplicateResult Refused(string reason) =>
            new DuplicateResult(Array.Empty<DuplicateCopy>(), new[] { reason }, false, notable: true);
    }

    /// <summary>
    /// Works out where each copy of a duplicated selection goes. Pure: no Office interop, no
    /// mutation of its inputs.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A separate entry point, like <see cref="ZOrderSolver"/>, because what it emits is not a change
    /// to an existing shape - it is a recipe for shapes that do not exist yet. Folding that into
    /// <see cref="AlignSolver"/> would stretch a contract that is defined by rewriting frames.
    /// </para>
    /// <para>
    /// One step is a rigid transform of the slide: turn by <see cref="DuplicateRequest.Angle"/>
    /// about the pivot, then move by the offsets. Copy <em>k</em> is the originals with that step
    /// applied <em>k</em> times, so a step that only moves gives a straight row, one that only turns
    /// about a far pivot gives a ring, and one that does both gives a spiral. A zero field is inert.
    /// </para>
    /// <para>
    /// The selection is duplicated as a unit: every shape gets the same step, so the internal layout
    /// of a multi-shape selection comes through intact - except under
    /// <see cref="DuplicatePivot.OwnCentre"/>, where each shape spins in place by design.
    /// </para>
    /// </remarks>
    public static class DuplicateSolver
    {
        public static DuplicateResult Solve(
            DuplicateRequest request, IReadOnlyList<ShapeSnapshot> shapes, SlideMetrics slide)
        {
            if (request is null) throw new ArgumentNullException(nameof(request));
            if (shapes is null) throw new ArgumentNullException(nameof(shapes));
            if (slide is null) throw new ArgumentNullException(nameof(slide));

            if (shapes.Count == 0) return DuplicateResult.Refused("Nothing is selected.");

            if (request.Copies < 1) return DuplicateResult.Refused("Make at least one copy.");
            if (request.Copies > DuplicateRequest.MaxCopies)
            {
                return DuplicateResult.Refused(string.Format(
                    CultureInfo.CurrentCulture,
                    "{0} copies is more than AlignPro makes at once. The limit is {1}.",
                    request.Copies, DuplicateRequest.MaxCopies));
            }

            if (!TryResolvePivot(request, shapes, slide, out var pivot, out var refusal))
            {
                return DuplicateResult.Refused(refusal!);
            }

            var theta = request.Angle * Math.PI / 180.0;
            var cos = Math.Cos(theta);
            var sin = Math.Sin(theta);

            // Walked one step at a time rather than computed in closed form per copy: k steps of a
            // rigid transform is exactly the step applied k times, and this reads as the definition.
            var centres = shapes.Select(s => (X: s.Frame.CentreX, Y: s.Frame.CentreY)).ToArray();
            var angles = shapes.Select(s => s.Rotation).ToArray();

            var copies = new List<DuplicateCopy>(request.Copies);
            for (var k = 1; k <= request.Copies; k++)
            {
                var placements = new List<ShapePlacement>(shapes.Count);

                for (var i = 0; i < shapes.Count; i++)
                {
                    var (x, y) = centres[i];

                    // Own centre has no fixed pivot: each shape's pivot is wherever it currently is,
                    // so the turn leaves its centre alone.
                    if (pivot.HasValue)
                    {
                        var (px, py) = pivot.Value;
                        var dx = x - px;
                        var dy = y - py;

                        // Y points down on a slide, so this standard rotation matrix turns clockwise
                        // on screen - the same sense as PowerPoint's Rotation property.
                        x = px + dx * cos - dy * sin;
                        y = py + dx * sin + dy * cos;
                    }

                    x += request.OffsetX;
                    y += request.OffsetY;
                    centres[i] = (x, y);

                    if (request.RotateShapes) angles[i] += request.Angle;

                    var frame = RectD.FromCentre(x, y, shapes[i].Frame.Width, shapes[i].Frame.Height);
                    placements.Add(new ShapePlacement(
                        shapes[i].Key, frame, GeometryChange.NormaliseAngle(angles[i])));
                }

                copies.Add(new DuplicateCopy(k, placements));
            }

            // A step that leaves everything where it was would stack every copy exactly on the
            // original, which looks like nothing happened while quietly adding shapes to the slide.
            if (LandsOnOriginals(shapes, copies[0]))
            {
                return DuplicateResult.Refused(
                    request.Angle != 0 && !request.RotateShapes && request.Pivot == DuplicatePivot.OwnCentre
                        ? "Turning each shape about its own centre with Rotate shapes off moves nothing. " +
                          "Set X or Y, pick another pivot, or turn Rotate shapes on."
                        : "Set X, Y or Angle - with all three at zero every copy lands on the original.");
            }

            var offSlide = CountOffSlide(copies, slide);
            if (offSlide > 0)
            {
                return DuplicateResult.OkWithNotice(copies, offSlide == 1
                    ? "1 copy lands entirely off the slide."
                    : string.Format(CultureInfo.CurrentCulture, "{0} copies land entirely off the slide.", offSlide));
            }

            return DuplicateResult.Ok(copies);
        }

        private static bool TryResolvePivot(
            DuplicateRequest request,
            IReadOnlyList<ShapeSnapshot> shapes,
            SlideMetrics slide,
            out (double X, double Y)? pivot,
            out string? refusal)
        {
            refusal = null;
            pivot = null;

            switch (request.Pivot)
            {
                case DuplicatePivot.OwnCentre:
                    return true;

                case DuplicatePivot.SlideCentre:
                    pivot = (slide.Width / 2, slide.Height / 2);
                    return true;

                case DuplicatePivot.AnchorCentre:
                    var anchor = request.Anchor.HasValue
                        ? shapes.FirstOrDefault(s => s.Key == request.Anchor.Value)
                        : null;
                    if (anchor is null)
                    {
                        refusal = "No anchor shape to turn about. Select the pivot shape last.";
                        return false;
                    }

                    pivot = (anchor.Frame.CentreX, anchor.Frame.CentreY);
                    return true;

                default:
                    // What the selection looks like, not its unrotated frames: the centre of a ring
                    // should be where the eye puts it.
                    var union = shapes[0].VisualBounds;
                    for (var i = 1; i < shapes.Count; i++) union = union.Union(shapes[i].VisualBounds);
                    pivot = (union.CentreX, union.CentreY);
                    return true;
            }
        }

        private static bool LandsOnOriginals(IReadOnlyList<ShapeSnapshot> shapes, DuplicateCopy first)
        {
            for (var i = 0; i < shapes.Count; i++)
            {
                var placement = first.Placements[i];
                if (!placement.Frame.ApproximatelyEquals(shapes[i].Frame, 1e-3)) return false;
                if (!GeometryChange.SameAngle(placement.Rotation, shapes[i].Rotation)) return false;
            }

            return true;
        }

        /// <summary>Copies whose every shape sits wholly outside the slide.</summary>
        private static int CountOffSlide(IReadOnlyList<DuplicateCopy> copies, SlideMetrics slide)
        {
            var count = 0;
            foreach (var copy in copies)
            {
                var allOff = copy.Placements.All(p =>
                    p.Frame.Right < 0 || p.Frame.Left > slide.Width ||
                    p.Frame.Bottom < 0 || p.Frame.Top > slide.Height);
                if (allOff) count++;
            }

            return count;
        }
    }
}
