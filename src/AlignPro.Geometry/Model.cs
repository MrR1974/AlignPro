using System;
using System.Collections.Generic;
using System.Globalization;

namespace AlignPro.Geometry
{
    /// <summary>
    /// Identifies a shape stably across an operation and its undo. <c>Shape.Name</c> is deliberately
    /// not used: PowerPoint does not enforce uniqueness on it.
    /// </summary>
    public readonly struct ShapeKey : IEquatable<ShapeKey>
    {
        public ShapeKey(int slideId, int shapeId)
        {
            SlideId = slideId;
            ShapeId = shapeId;
        }

        /// <summary>PowerPoint's <c>Slide.SlideID</c>, stable for the life of the slide.</summary>
        public int SlideId { get; }

        /// <summary>PowerPoint's <c>Shape.Id</c>, unique within the slide.</summary>
        public int ShapeId { get; }

        public bool Equals(ShapeKey other) => SlideId == other.SlideId && ShapeId == other.ShapeId;

        public override bool Equals(object? obj) => obj is ShapeKey other && Equals(other);

        public override int GetHashCode() => unchecked((SlideId * 397) ^ ShapeId);

        public static bool operator ==(ShapeKey left, ShapeKey right) => left.Equals(right);

        public static bool operator !=(ShapeKey left, ShapeKey right) => !left.Equals(right);

        public override string ToString() =>
            string.Format(CultureInfo.InvariantCulture, "{0}:{1}", SlideId, ShapeId);
    }

    /// <summary>
    /// An immutable reading of one shape's geometry, taken from PowerPoint before the solver runs.
    /// The solver sees only this - never a live COM object.
    /// </summary>
    public sealed class ShapeSnapshot
    {
        public ShapeSnapshot(
            ShapeKey key,
            RectD frame,
            double rotation = 0,
            bool flipH = false,
            bool flipV = false,
            RectD? textBounds = null,
            bool isGroup = false,
            bool isPlaceholder = false,
            string? name = null)
        {
            Key = key;
            Frame = frame;
            Rotation = rotation;
            FlipH = flipH;
            FlipV = flipV;
            TextBounds = textBounds;
            IsGroup = isGroup;
            IsPlaceholder = isPlaceholder;
            Name = name;
        }

        public ShapeKey Key { get; }

        /// <summary>
        /// The object model's own rectangle, which ignores rotation entirely - measured, see
        /// <c>docs/object-model-findings.md</c> probe 1. The shape's centre is invariant under rotation.
        /// </summary>
        public RectD Frame { get; }

        /// <summary>Clockwise rotation in degrees about the frame's centre.</summary>
        public double Rotation { get; }

        public bool FlipH { get; }

        public bool FlipV { get; }

        /// <summary>
        /// Text bounds from <c>TextRange2.Bound*</c>, or null when the shape holds no text. Already
        /// in slide coordinates.
        /// </summary>
        public RectD? TextBounds { get; }

        /// <summary>
        /// True for a group. Groups are always treated as one object: the solver never descends into
        /// them, so translation preserves internal spacing for free. Resize verbs refuse them,
        /// because scaling a group scales the gaps between its children.
        /// </summary>
        public bool IsGroup { get; }

        public bool IsPlaceholder { get; }

        /// <summary>Diagnostic only - never used for identity.</summary>
        public string? Name { get; }

        /// <summary>
        /// The rotation-aware rectangle: what the eye sees. Derived purely from frame and rotation,
        /// since the centre holds. Exact for rectangles; for other geometry it is the bounding box of
        /// the rotated frame, which over-estimates. Flips do not affect it.
        /// </summary>
        public RectD VisualBounds
        {
            get
            {
                var theta = Rotation * Math.PI / 180.0;
                var cos = Math.Abs(Math.Cos(theta));
                var sin = Math.Abs(Math.Sin(theta));
                var width = Frame.Width * cos + Frame.Height * sin;
                var height = Frame.Width * sin + Frame.Height * cos;
                return RectD.FromCentre(Frame.CentreX, Frame.CentreY, width, height);
            }
        }

        public override string ToString() =>
            string.Format(CultureInfo.InvariantCulture, "{0} [{1}] {2}", Name ?? "(unnamed)", Key, Frame);
    }

    /// <summary>
    /// One shape's new frame, and optionally its new rotation. The solver always emits frame
    /// coordinates, whatever bounds model the request asked for, because the frame is the only
    /// rectangle PowerPoint lets us write.
    /// </summary>
    /// <remarks>
    /// Rotation is optional rather than always carried, and that is load-bearing: a change that
    /// carries no rotation leaves the shape's angle alone, so the align verbs cannot clobber a
    /// rotation they never read. PowerPoint rotates about the frame's centre, so a rotation-only
    /// change leaves the frame untouched and the two never interfere.
    /// </remarks>
    public sealed class GeometryChange
    {
        /// <summary>How close two angles must be, in degrees, to count as the same.</summary>
        public const double AngleEpsilon = 1e-4;

        public GeometryChange(
            ShapeKey key, RectD oldFrame, RectD newFrame, double? oldRotation = null, double? newRotation = null)
        {
            if (oldRotation.HasValue != newRotation.HasValue)
            {
                throw new ArgumentException(
                    "A rotation change needs both the old and the new angle, or neither.", nameof(newRotation));
            }

            Key = key;
            OldFrame = oldFrame;
            NewFrame = newFrame;
            OldRotation = oldRotation;
            NewRotation = newRotation;
        }

        public ShapeKey Key { get; }

        /// <summary>The frame as it was, so the change is its own undo record.</summary>
        public RectD OldFrame { get; }

        public RectD NewFrame { get; }

        /// <summary>The clockwise angle in degrees as it was, or null when rotation is not changing.</summary>
        public double? OldRotation { get; }

        /// <summary>The clockwise angle in degrees to write, or null to leave the angle alone.</summary>
        public double? NewRotation { get; }

        /// <summary>True when this change touches the angle at all.</summary>
        public bool ChangesRotation =>
            OldRotation.HasValue && NewRotation.HasValue && !SameAngle(OldRotation.Value, NewRotation.Value);

        /// <summary>True when nothing would actually move, so the apply step can skip it.</summary>
        public bool IsNoOp => OldFrame.ApproximatelyEquals(NewFrame) && !ChangesRotation;

        /// <summary>The same change running backwards, ready to apply as an undo.</summary>
        public GeometryChange Inverted() => new GeometryChange(Key, NewFrame, OldFrame, NewRotation, OldRotation);

        /// <summary>An angle folded into [0, 360), the range PowerPoint itself reports.</summary>
        public static double NormaliseAngle(double degrees)
        {
            var folded = degrees % 360.0;
            if (folded < 0) folded += 360.0;

            // 359.99999 and 0 are the same angle; report the one PowerPoint would.
            return folded >= 360.0 - AngleEpsilon ? 0 : folded;
        }

        /// <summary>True when two angles point the same way, however many turns apart.</summary>
        public static bool SameAngle(double a, double b)
        {
            var difference = Math.Abs(NormaliseAngle(a) - NormaliseAngle(b));
            return difference <= AngleEpsilon || difference >= 360.0 - AngleEpsilon;
        }

        public override string ToString() => NewRotation.HasValue
            ? string.Format(
                CultureInfo.InvariantCulture, "{0}: {1} rot={2:F1} -> {3} rot={4:F1}",
                Key, OldFrame, OldRotation, NewFrame, NewRotation)
            : string.Format(CultureInfo.InvariantCulture, "{0}: {1} -> {2}", Key, OldFrame, NewFrame);
    }

    /// <summary>
    /// A change to the slide's stacking order, recorded as the whole slide's ordering before and
    /// after, back to front.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The whole slide is recorded rather than just the selection because PowerPoint gives no way to
    /// write a shape's <c>ZOrderPosition</c> directly - the only lever is "bring this one to the
    /// front". Replaying a complete ordering front-wards with that lever lands on exactly the order
    /// asked for, and nothing else needs to be worked out. It also makes the inverse trivial: the
    /// ordering that was there before is itself a complete instruction for getting back to it.
    /// </para>
    /// <para>
    /// Only top-level shapes appear here. A shape inside a group has a z-position within its group
    /// rather than within the slide, so ordering one against slide-level shapes is meaningless - the
    /// reader refuses that selection rather than recording something it cannot honour.
    /// </para>
    /// </remarks>
    public sealed class ZOrderChange
    {
        public ZOrderChange(IReadOnlyList<ShapeKey> oldOrder, IReadOnlyList<ShapeKey> newOrder)
        {
            OldOrder = oldOrder ?? throw new ArgumentNullException(nameof(oldOrder));
            NewOrder = newOrder ?? throw new ArgumentNullException(nameof(newOrder));
        }

        /// <summary>The slide's stacking order as it was, back to front.</summary>
        public IReadOnlyList<ShapeKey> OldOrder { get; }

        /// <summary>The stacking order to apply, back to front.</summary>
        public IReadOnlyList<ShapeKey> NewOrder { get; }

        /// <summary>True when nothing would actually move, so the apply step can skip it.</summary>
        public bool IsNoOp
        {
            get
            {
                if (OldOrder.Count != NewOrder.Count) return false;
                for (var i = 0; i < OldOrder.Count; i++)
                {
                    if (OldOrder[i] != NewOrder[i]) return false;
                }

                return true;
            }
        }

        /// <summary>The same change running backwards, ready to apply as an undo.</summary>
        public ZOrderChange Inverted() => new ZOrderChange(NewOrder, OldOrder);

        public override string ToString() =>
            string.Format(CultureInfo.InvariantCulture, "z-order of {0} shapes", NewOrder.Count);
    }

    /// <summary>Slide-level geometry the solver needs to resolve non-selection references.</summary>
    public sealed class SlideMetrics
    {
        public SlideMetrics(double width, double height, double margin = 0, RectD? placeholderBounds = null)
        {
            if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width), width, "Slide width must be positive.");
            if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height), height, "Slide height must be positive.");
            if (margin < 0) throw new ArgumentOutOfRangeException(nameof(margin), margin, "Margin cannot be negative.");

            Width = width;
            Height = height;
            Margin = margin;
            PlaceholderBounds = placeholderBounds;
        }

        public double Width { get; }

        public double Height { get; }

        /// <summary>Inset used by <see cref="ReferenceTarget.SlideMargins"/>, in points.</summary>
        public double Margin { get; }

        /// <summary>The layout's body placeholder rectangle, when one could be read.</summary>
        public RectD? PlaceholderBounds { get; }

        public RectD Bounds => new RectD(0, 0, Width, Height);

        /// <summary>A 16:9 widescreen slide, the modern PowerPoint default.</summary>
        public static SlideMetrics Widescreen(double margin = 0) => new SlideMetrics(960, 540, margin);
    }

    /// <summary>Everything the solver needs beyond the shapes themselves.</summary>
    public sealed class AlignRequest
    {
        public AlignRequest(
            AlignVerb verb,
            ReferenceTarget reference = ReferenceTarget.SelectionBounds,
            BoundsModel boundsModel = BoundsModel.ShapeFrame,
            ShapeKey? anchor = null,
            double? exactSpacing = null,
            DistributeMode distributeMode = DistributeMode.Gap,
            ResizeOrigin resizeOrigin = ResizeOrigin.TopLeft,
            bool allowGroupResize = false,
            double sizeMargin = 0,
            SizeMarginMode sizeMarginMode = SizeMarginMode.None,
            int? gridColumns = null,
            double gridGapH = 0,
            double gridGapV = 0,
            GridFillOrder gridFillOrder = GridFillOrder.RowMajor)
        {
            Verb = verb;
            Reference = reference;
            BoundsModel = boundsModel;
            Anchor = anchor;
            ExactSpacing = exactSpacing;
            DistributeMode = distributeMode;
            ResizeOrigin = resizeOrigin;
            AllowGroupResize = allowGroupResize;
            SizeMargin = sizeMargin;
            SizeMarginMode = sizeMarginMode;
            GridColumns = gridColumns;
            GridGapH = gridGapH;
            GridGapV = gridGapV;
            GridFillOrder = gridFillOrder;
        }

        public AlignVerb Verb { get; }

        public ReferenceTarget Reference { get; }

        public BoundsModel BoundsModel { get; }

        /// <summary>
        /// The anchor shape. Required when <see cref="Reference"/> is
        /// <see cref="ReferenceTarget.Anchor"/> and by every match-size verb. Selection order is
        /// preserved by PowerPoint (measured, probe 2), so the caller can default this to the last
        /// shape selected.
        /// </summary>
        public ShapeKey? Anchor { get; }

        /// <summary>
        /// An exact gap or pitch in points. When null the distribute verbs equalise whatever space
        /// the selection already spans, holding the outermost shapes still.
        /// </summary>
        public double? ExactSpacing { get; }

        public DistributeMode DistributeMode { get; }

        public ResizeOrigin ResizeOrigin { get; }

        /// <summary>
        /// Allows match-size to act on a group. Off by default: scaling a group scales the gaps
        /// between its children (measured, probe 3), which silently distorts diagrams.
        /// </summary>
        public bool AllowGroupResize { get; }

        /// <summary>
        /// How much smaller than the anchor each match-size step is, measured <em>per side</em> in
        /// points. So one step takes twice this off the width and off the height, and a 10pt margin
        /// leaves a 10pt border showing all round. Negative values grow the shapes instead.
        /// </summary>
        /// <remarks>Ignored unless <see cref="SizeMarginMode"/> says otherwise.</remarks>
        public double SizeMargin { get; }

        /// <summary>
        /// Whether <see cref="SizeMargin"/> applies at all, once to every shape, or cumulatively
        /// along the selection.
        /// </summary>
        public SizeMarginMode SizeMarginMode { get; }

        /// <summary>Columns for <see cref="AlignVerb.GridArrange"/>. Defaults to a near-square grid.</summary>
        public int? GridColumns { get; }

        public double GridGapH { get; }

        public double GridGapV { get; }

        public GridFillOrder GridFillOrder { get; }
    }

    /// <summary>
    /// The outcome of a solve. A refusal is not an exception: it comes back as
    /// <see cref="Succeeded"/> false with a reason the ribbon can show in a status line.
    /// </summary>
    public sealed class SolveResult
    {
        private SolveResult(
            IReadOnlyList<GeometryChange> changes,
            IReadOnlyList<string> diagnostics,
            bool succeeded,
            bool notable)
        {
            Changes = changes;
            Diagnostics = diagnostics;
            Succeeded = succeeded;
            Notable = notable;
        }

        /// <summary>
        /// True when the user needs telling, because part of what they asked for did not happen -
        /// shapes were skipped, or nothing needed changing. Without this the caller cannot distinguish
        /// a quiet success from a deliberate refusal to touch something, and a button that declines to
        /// act while saying nothing just looks broken.
        /// </summary>
        public bool Notable { get; }

        public IReadOnlyList<GeometryChange> Changes { get; }

        /// <summary>Human-readable notes: why nothing happened, or what was skipped.</summary>
        public IReadOnlyList<string> Diagnostics { get; }

        public bool Succeeded { get; }

        /// <summary>Changes that would actually move something.</summary>
        public IEnumerable<GeometryChange> EffectiveChanges
        {
            get
            {
                foreach (var change in Changes)
                {
                    if (!change.IsNoOp) yield return change;
                }
            }
        }

        public static SolveResult Ok(IReadOnlyList<GeometryChange> changes, params string[] diagnostics) =>
            new SolveResult(changes, diagnostics, true, notable: false);

        /// <summary>Succeeded, but something the user asked for was skipped and they should be told.</summary>
        public static SolveResult OkWithNotice(IReadOnlyList<GeometryChange> changes, params string[] diagnostics) =>
            new SolveResult(changes, diagnostics, true, notable: true);

        public static SolveResult Refused(string reason) =>
            new SolveResult(Array.Empty<GeometryChange>(), new[] { reason }, false, notable: true);
    }
}
