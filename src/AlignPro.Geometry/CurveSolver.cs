using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace AlignPro.Geometry
{
    /// <summary>A point in slide coordinates, in points, Y increasing downwards.</summary>
    public readonly struct PointD
    {
        public PointD(double x, double y)
        {
            X = x;
            Y = y;
        }

        public double X { get; }

        public double Y { get; }

        public double DistanceTo(PointD other)
        {
            var dx = other.X - X;
            var dy = other.Y - Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        public override string ToString() => string.Format(CultureInfo.InvariantCulture, "({0:F2}, {1:F2})", X, Y);
    }

    /// <summary>What kind of segment leaves a path node, as PowerPoint's <c>ShapeNode.SegmentType</c> says.</summary>
    public enum PathSegmentKind
    {
        /// <summary>A straight line to the next node.</summary>
        Line,

        /// <summary>
        /// A cubic Bezier: the next two nodes are its control points and the one after is where it
        /// ends. PowerPoint lists control points as nodes in their own right.
        /// </summary>
        Curve
    }

    /// <summary>One entry of a freeform's node list, converted out of COM and nothing more.</summary>
    public readonly struct PathNode
    {
        public PathNode(double x, double y, PathSegmentKind segment = PathSegmentKind.Line)
        {
            Point = new PointD(x, y);
            Segment = segment;
        }

        public PointD Point { get; }

        public PathSegmentKind Segment { get; }
    }

    /// <summary>Which kind of shape the anchor is, as far as the curve is concerned.</summary>
    public enum CurveKind
    {
        /// <summary>An ellipse or circle filling the frame. A closed loop.</summary>
        Ellipse,

        /// <summary>
        /// PowerPoint's Arc autoshape. Not the ellipse filling the frame: the frame is the bounding
        /// box of the arc and the ellipse's centre together (probe 10), so the ellipse is rebuilt
        /// from the frame and the two angles.
        /// </summary>
        Arc,

        /// <summary>A freeform or drawn curve, given as its nodes.</summary>
        Path,

        /// <summary>A straight line, corner to corner of its frame. Lines have no nodes to read.</summary>
        Line
    }

    /// <summary>
    /// The curve the anchor defines, in the anchor's own terms: angles for an arc, nodes for a path.
    /// Turned into slide coordinates by <see cref="CurveSolver"/> using the anchor's frame, flips and
    /// rotation, so the reader only has to convert what PowerPoint reports.
    /// </summary>
    public sealed class CurveDefinition
    {
        private CurveDefinition(CurveKind kind, double startAngle, double endAngle, IReadOnlyList<PathNode> nodes)
        {
            Kind = kind;
            StartAngle = startAngle;
            EndAngle = endAngle;
            Nodes = nodes;
        }

        public CurveKind Kind { get; }

        /// <summary>
        /// Where an arc starts, in degrees clockwise from three o'clock, measured as the direction
        /// from the ellipse's centre - not the ellipse's parameter, which differs unless it is a circle.
        /// </summary>
        public double StartAngle { get; }

        /// <summary>Where an arc ends. The arc runs clockwise from <see cref="StartAngle"/> to here.</summary>
        public double EndAngle { get; }

        /// <summary>
        /// A path's nodes in slide coordinates exactly as drawn: PowerPoint reports them with the
        /// shape's rotation and flips already applied (probe 9). Empty for the other kinds.
        /// </summary>
        public IReadOnlyList<PathNode> Nodes { get; }

        public static CurveDefinition Ellipse() =>
            new CurveDefinition(CurveKind.Ellipse, 0, 0, Array.Empty<PathNode>());

        /// <summary>
        /// A straight line from the frame's top-left to its bottom-right. Which diagonal it really
        /// runs along comes from the flips, which the solver applies.
        /// </summary>
        public static CurveDefinition Line() =>
            new CurveDefinition(CurveKind.Line, 0, 0, Array.Empty<PathNode>());

        public static CurveDefinition Arc(double startAngle, double endAngle) =>
            new CurveDefinition(CurveKind.Arc, startAngle, endAngle, Array.Empty<PathNode>());

        public static CurveDefinition Path(IReadOnlyList<PathNode> nodes)
        {
            if (nodes is null) throw new ArgumentNullException(nameof(nodes));
            return new CurveDefinition(CurveKind.Path, 0, 0, nodes);
        }
    }

    /// <summary>Everything the curve solver needs beyond the shapes themselves.</summary>
    public sealed class CurveRequest
    {
        public CurveRequest(ShapeKey anchor, CurveDefinition curve, double? exactSpacing = null, bool rotateShapes = true)
        {
            Anchor = anchor;
            Curve = curve ?? throw new ArgumentNullException(nameof(curve));
            ExactSpacing = exactSpacing;
            RotateShapes = rotateShapes;
        }

        /// <summary>The shape that defines the curve - by default the last one selected.</summary>
        public ShapeKey Anchor { get; }

        public CurveDefinition Curve { get; }

        /// <summary>
        /// Distance between neighbouring centres, measured along the curve. Null spreads the shapes
        /// evenly over the whole curve.
        /// </summary>
        public double? ExactSpacing { get; }

        /// <summary>Turn each shape to follow the curve, or leave its angle alone.</summary>
        public bool RotateShapes { get; }
    }

    /// <summary>A curve flattened into straight segments, measured by length.</summary>
    public sealed class Polyline
    {
        private readonly double[] _cumulative;

        public Polyline(IReadOnlyList<PointD> points, bool closed)
        {
            if (points is null) throw new ArgumentNullException(nameof(points));
            if (points.Count < 2) throw new ArgumentException("A polyline needs at least two points.", nameof(points));

            Points = points;
            Closed = closed;

            _cumulative = new double[points.Count];
            for (var i = 1; i < points.Count; i++)
            {
                _cumulative[i] = _cumulative[i - 1] + points[i - 1].DistanceTo(points[i]);
            }
        }

        /// <summary>The vertices in order. A closed polyline repeats its first point at the end.</summary>
        public IReadOnlyList<PointD> Points { get; }

        /// <summary>True for a loop, whose end is its start.</summary>
        public bool Closed { get; }

        public double Length => _cumulative[_cumulative.Length - 1];

        /// <summary>
        /// The point a given distance along the curve, and the direction of travel there in degrees
        /// clockwise from pointing right.
        /// </summary>
        public (PointD Point, double Tangent) At(double distance)
        {
            distance = Math.Max(0, Math.Min(Length, distance));
            var point = PointAt(distance);

            // The direction across a short span either side, not the one segment the point is on. A
            // flattened curve's segments each lean by half a step against the true tangent, which on
            // a ring shows as 89.8 degrees where 90 was meant; a span centred on the point cancels
            // that. At a sharp corner it gives the bisector, which is what following a path means.
            var before = distance - TangentSpan;
            var after = distance + TangentSpan;
            if (!Closed)
            {
                before = Math.Max(0, before);
                after = Math.Min(Length, after);
            }

            var a = PointAt(before);
            var b = PointAt(after);
            var tangent = Math.Atan2(b.Y - a.Y, b.X - a.X) * 180.0 / Math.PI;
            return (point, GeometryChange.NormaliseAngle(tangent));
        }

        /// <summary>Half the span the tangent is measured across, in points.</summary>
        private const double TangentSpan = 0.25;

        /// <summary>The point a distance along; a closed curve wraps round, an open one stops at its ends.</summary>
        private PointD PointAt(double distance)
        {
            if (Closed)
            {
                distance %= Length;
                if (distance < 0) distance += Length;
            }
            else
            {
                distance = Math.Max(0, Math.Min(Length, distance));
            }

            var segment = 0;
            while (segment < Points.Count - 2 && _cumulative[segment + 1] <= distance) segment++;

            var a = Points[segment];
            var b = Points[segment + 1];
            var span = _cumulative[segment + 1] - _cumulative[segment];
            var t = span > RectD.Epsilon ? (distance - _cumulative[segment]) / span : 0;
            return new PointD(a.X + (b.X - a.X) * t, a.Y + (b.Y - a.Y) * t);
        }
    }

    /// <summary>
    /// Places shapes along the curve the anchor defines: round an ellipse, along an arc, or along any
    /// path. Pure: no Office interop, no mutation of its inputs.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Every kind of curve is flattened to a <see cref="Polyline"/> and shapes are placed by distance
    /// along it, so one placement rule covers all three. Spacing by distance rather than by angle is
    /// the point: stepping an ellipse's parameter evenly bunches shapes at the narrow ends, which is
    /// the first thing anyone notices on a flattened ring.
    /// </para>
    /// <para>
    /// Emits ordinary <see cref="GeometryChange"/>s, frame plus optional angle, so undo and redo come
    /// for free through the same pipeline as every other verb.
    /// </para>
    /// </remarks>
    public static class CurveSolver
    {
        /// <summary>Straight segments per full turn of an ellipse. Well under a hundredth of a point off a circle.</summary>
        private const int EllipseSegments = 720;

        /// <summary>Straight segments per Bezier. Plenty for placement, where only the length matters.</summary>
        private const int BezierSegments = 32;

        public static SolveResult Solve(CurveRequest request, IReadOnlyList<ShapeSnapshot> shapes)
        {
            if (request is null) throw new ArgumentNullException(nameof(request));
            if (shapes is null) throw new ArgumentNullException(nameof(shapes));

            var anchor = shapes.FirstOrDefault(s => s.Key == request.Anchor);
            if (anchor is null) return SolveResult.Refused("The curve shape is not part of the selection.");

            var toPlace = shapes.Where(s => s.Key != anchor.Key).ToList();
            if (toPlace.Count == 0)
            {
                return SolveResult.Refused("Select the shapes to place first, then the curve to place them on last.");
            }

            Polyline curve;
            try
            {
                curve = Build(request.Curve, anchor);
            }
            catch (ArgumentException ex)
            {
                return SolveResult.Refused(ex.Message);
            }

            if (curve.Length <= RectD.Epsilon) return SolveResult.Refused("The curve has no length to place shapes along.");

            if (!TryDistances(request.ExactSpacing, curve, toPlace.Count, out var distances, out var refusal))
            {
                return SolveResult.Refused(refusal!);
            }

            var changes = new List<GeometryChange>(toPlace.Count);
            for (var i = 0; i < toPlace.Count; i++)
            {
                var shape = toPlace[i];
                var (point, tangent) = curve.At(distances[i]);
                var frame = RectD.FromCentre(point.X, point.Y, shape.Frame.Width, shape.Frame.Height);

                changes.Add(request.RotateShapes
                    ? new GeometryChange(shape.Key, shape.Frame, frame, shape.Rotation, tangent)
                    : new GeometryChange(shape.Key, shape.Frame, frame));
            }

            return SolveResult.Ok(changes);
        }

        /// <summary>
        /// Where along the curve each shape's centre goes. A closed curve gets a full circuit of even
        /// slots, so the last shape does not land on the first; an open one includes both ends.
        /// </summary>
        private static bool TryDistances(
            double? exactSpacing, Polyline curve, int count, out double[] distances, out string? refusal)
        {
            distances = new double[count];
            refusal = null;

            if (exactSpacing.HasValue)
            {
                var spacing = exactSpacing.Value;
                if (spacing <= 0)
                {
                    refusal = "Exact spacing along a curve must be more than zero.";
                    return false;
                }

                // A closed curve needs room for the gap back round to the first shape as well.
                var needed = spacing * (curve.Closed ? count : count - 1);
                if (needed > curve.Length + 1e-6)
                {
                    refusal = string.Format(
                        CultureInfo.CurrentCulture,
                        "{0} shapes {1:0.#}pt apart need {2:0.#}pt of curve, and this one is {3:0.#}pt long. " +
                        "Reduce Exact (pt), or clear it to spread them over the whole curve.",
                        count, spacing, needed, curve.Length);
                    return false;
                }

                for (var i = 0; i < count; i++) distances[i] = i * spacing;
                return true;
            }

            if (curve.Closed)
            {
                for (var i = 0; i < count; i++) distances[i] = curve.Length * i / count;
            }
            else if (count == 1)
            {
                distances[0] = curve.Length / 2;
            }
            else
            {
                for (var i = 0; i < count; i++) distances[i] = curve.Length * i / (count - 1);
            }

            return true;
        }

        /// <summary>
        /// Flattens the anchor's curve into slide coordinates: built in the anchor's unrotated frame,
        /// then flipped and turned about the frame's centre exactly as PowerPoint draws it.
        /// </summary>
        public static Polyline Build(CurveDefinition curve, ShapeSnapshot anchor)
        {
            if (curve is null) throw new ArgumentNullException(nameof(curve));
            if (anchor is null) throw new ArgumentNullException(nameof(anchor));

            if (curve.Kind == CurveKind.Path)
            {
                // Already where it is drawn - transforming it again would turn it twice.
                var points = FlattenPath(curve.Nodes);
                var loops = points.Count > 2 && points[0].DistanceTo(points[points.Count - 1]) < 1e-3;
                return new Polyline(points, loops);
            }

            List<PointD> local;
            bool closed;

            switch (curve.Kind)
            {
                case CurveKind.Ellipse:
                    // From twelve o'clock, clockwise: where a person starts a ring of anything.
                    local = SampleEllipse(
                        anchor.Frame.CentreX, anchor.Frame.CentreY, anchor.Frame.Width / 2, anchor.Frame.Height / 2,
                        -90, 270, byDirection: false);
                    closed = true;
                    break;

                case CurveKind.Arc:
                    // A rotated arc's frame is not its unrotated frame (probe 10), so there is nothing
                    // honest to rebuild the ellipse from.
                    if (!GeometryChange.SameAngle(anchor.Rotation, 0))
                    {
                        throw new ArgumentException(
                            "AlignPro cannot follow a rotated arc: PowerPoint reports its frame differently " +
                            "once it is turned. Set the arc's rotation to 0, or use an oval or a freeform.");
                    }

                    local = SampleArc(anchor.Frame, curve.StartAngle, curve.EndAngle);
                    closed = false;
                    break;

                default:
                    local = new List<PointD>
                    {
                        new PointD(anchor.Frame.Left, anchor.Frame.Top),
                        new PointD(anchor.Frame.Right, anchor.Frame.Bottom)
                    };
                    closed = false;
                    break;
            }

            var cx = anchor.Frame.CentreX;
            var cy = anchor.Frame.CentreY;
            var theta = anchor.Rotation * Math.PI / 180.0;
            var cos = Math.Cos(theta);
            var sin = Math.Sin(theta);

            var placed = new List<PointD>(local.Count);
            foreach (var p in local)
            {
                var dx = p.X - cx;
                var dy = p.Y - cy;

                // Flip first, then turn: PowerPoint mirrors the shape within its frame and rotates
                // the result.
                if (anchor.FlipH) dx = -dx;
                if (anchor.FlipV) dy = -dy;

                placed.Add(new PointD(cx + dx * cos - dy * sin, cy + dx * sin + dy * cos));
            }

            return new Polyline(placed, closed);
        }

        /// <summary>
        /// Points along an Arc autoshape, from its start angle clockwise to its end.
        /// </summary>
        /// <remarks>
        /// <para>
        /// PowerPoint's Arc keeps its ellipse fixed and sizes the frame to the bounding box of the arc
        /// together with the ellipse's centre - a pie slice's box, measured in probe 10. So the default
        /// quarter, twelve o'clock to three, has its centre at the frame's bottom-left corner, and the
        /// ellipse is twice the frame each way.
        /// </para>
        /// <para>
        /// Rebuilding the ellipse means finding radii whose slice fits the frame exactly. The angles
        /// are directions from the centre, so where the ends land along each axis depends on how
        /// stretched the ellipse is - which is the unknown. That leaves one equation in one unknown,
        /// the ratio of the radii, solved here by bisection; the rest follows directly.
        /// </para>
        /// </remarks>
        private static List<PointD> SampleArc(RectD frame, double startDegrees, double endDegrees)
        {
            while (endDegrees <= startDegrees) endDegrees += 360;

            if (frame.Width <= RectD.Epsilon || frame.Height <= RectD.Epsilon)
            {
                throw new ArgumentException("The arc is too thin to follow.");
            }

            // g(k) = k * spanX(k) / spanY(k) must equal the frame's aspect, with k = rx / ry.
            var target = frame.Width / frame.Height;
            double Aspect(double k)
            {
                var span = SliceSpan(startDegrees, endDegrees, k);
                var height = span.MaxY - span.MinY;
                return height <= 1e-12 ? double.PositiveInfinity : k * (span.MaxX - span.MinX) / height;
            }

            var lo = Math.Log(1e-4);
            var hi = Math.Log(1e4);
            for (var i = 0; i < 200; i++)
            {
                var mid = (lo + hi) / 2;
                if (Aspect(Math.Exp(mid)) < target) lo = mid; else hi = mid;
            }

            var ratio = Math.Exp((lo + hi) / 2);
            var slice = SliceSpan(startDegrees, endDegrees, ratio);
            var spanX = slice.MaxX - slice.MinX;
            var spanY = slice.MaxY - slice.MinY;
            if (spanX <= 1e-9 || spanY <= 1e-9 || Math.Abs(Aspect(ratio) - target) > 1e-4 * target)
            {
                throw new ArgumentException("Could not work out the ellipse this arc belongs to.");
            }

            var rx = frame.Width / spanX;
            var ry = frame.Height / spanY;
            var cx = frame.Left - rx * slice.MinX;
            var cy = frame.Top - ry * slice.MinY;

            return SampleEllipse(cx, cy, rx, ry, startDegrees, endDegrees, byDirection: true);
        }

        /// <summary>
        /// The extent of a pie slice of a unit-height ellipse whose width is <paramref name="ratio"/>
        /// times its height - the arc from one direction to the other, plus the centre - in units of
        /// each radius.
        /// </summary>
        private static (double MinX, double MaxX, double MinY, double MaxY) SliceSpan(
            double startDegrees, double endDegrees, double ratio)
        {
            // Direction to parameter: the point at direction theta on x = a cos t, y = b sin t has
            // tan theta = (b sin t) / (a cos t), so t = atan2(a sin theta, b cos theta). The
            // parameter stays in the same quadrant as the direction, so the axis crossings - where
            // the extremes are - are the same for both.
            double Parameter(double degrees)
            {
                var theta = degrees * Math.PI / 180.0;
                return Math.Atan2(ratio * Math.Sin(theta), Math.Cos(theta));
            }

            var ts = Parameter(startDegrees);
            var te = Parameter(endDegrees);

            double minX = 0, maxX = 0, minY = 0, maxY = 0;   // the centre is in the box
            void Include(double t)
            {
                var c = Math.Cos(t);
                var s = Math.Sin(t);
                minX = Math.Min(minX, c); maxX = Math.Max(maxX, c);
                minY = Math.Min(minY, s); maxY = Math.Max(maxY, s);
            }

            Include(ts);
            Include(te);

            // Every axis the sweep crosses is an extreme: 0, 90, 180 and 270 degrees, in any turn.
            var first = Math.Ceiling(startDegrees / 90.0) * 90.0;
            for (var axis = first; axis < endDegrees; axis += 90.0)
            {
                if (axis > startDegrees) Include(axis * Math.PI / 180.0);
            }

            return (minX, maxX, minY, maxY);
        }

        /// <summary>
        /// Points round an ellipse from one angle to another, clockwise.
        /// </summary>
        /// <param name="cx">The ellipse's centre, across.</param>
        /// <param name="cy">The ellipse's centre, down.</param>
        /// <param name="rx">The horizontal radius.</param>
        /// <param name="ry">The vertical radius.</param>
        /// <param name="fromDegrees">Where to start, clockwise from three o'clock.</param>
        /// <param name="toDegrees">Where to stop; more than <paramref name="fromDegrees"/>.</param>
        /// <param name="byDirection">
        /// True when the angles are directions from the centre - how an Arc autoshape states them -
        /// rather than the ellipse's parameter. On a circle the two agree.
        /// </param>
        private static List<PointD> SampleEllipse(
            double cx, double cy, double rx, double ry, double fromDegrees, double toDegrees, bool byDirection)
        {
            var sweep = toDegrees - fromDegrees;
            var steps = Math.Max(8, (int)Math.Ceiling(EllipseSegments * sweep / 360.0));

            var points = new List<PointD>(steps + 1);
            for (var i = 0; i <= steps; i++)
            {
                var angle = (fromDegrees + sweep * i / steps) * Math.PI / 180.0;
                var cos = Math.Cos(angle);
                var sin = Math.Sin(angle);

                double x, y;
                if (byDirection)
                {
                    // Where the ray at this angle from the centre meets the ellipse.
                    var r = rx * ry / Math.Sqrt(ry * ry * cos * cos + rx * rx * sin * sin);
                    x = r * cos;
                    y = r * sin;
                }
                else
                {
                    x = rx * cos;
                    y = ry * sin;
                }

                points.Add(new PointD(cx + x, cy + y));
            }

            return points;
        }

        /// <summary>
        /// Turns a freeform's node list into points. A node whose segment is a curve is followed by
        /// its two control points and then the node the curve ends on.
        /// </summary>
        private static List<PointD> FlattenPath(IReadOnlyList<PathNode> nodes)
        {
            if (nodes.Count < 2) throw new ArgumentException("The path has fewer than two points.");

            var points = new List<PointD> { nodes[0].Point };
            var i = 0;
            while (i < nodes.Count - 1)
            {
                if (nodes[i].Segment == PathSegmentKind.Curve && i + 3 < nodes.Count)
                {
                    var p0 = nodes[i].Point;
                    var c1 = nodes[i + 1].Point;
                    var c2 = nodes[i + 2].Point;
                    var p3 = nodes[i + 3].Point;

                    for (var step = 1; step <= BezierSegments; step++)
                    {
                        var t = step / (double)BezierSegments;
                        var u = 1 - t;
                        var a = u * u * u;
                        var b = 3 * u * u * t;
                        var c = 3 * u * t * t;
                        var d = t * t * t;
                        points.Add(new PointD(
                            a * p0.X + b * c1.X + c * c2.X + d * p3.X,
                            a * p0.Y + b * c1.Y + c * c2.Y + d * p3.Y));
                    }

                    i += 3;
                }
                else
                {
                    points.Add(nodes[i + 1].Point);
                    i++;
                }
            }

            return points;
        }
    }
}
