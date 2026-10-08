using System;
using System.Collections.Generic;
using System.Linq;

namespace AlignPro.Geometry
{
    /// <summary>
    /// Solves a request for shapes inside a rotated group along the group's own axes, so "left" is the
    /// group's left as it is seen, not the slide's.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A change of coordinates around <see cref="AlignSolver"/>, not a new solver. PowerPoint reports a
    /// shape inside a rotated group as it really sits on the slide, and takes writes the same way
    /// (<c>docs/object-model-findings.md</c>, probe 14). So every snapshot is turned about the group's
    /// centre by minus the group's angle - its frame's centre moves, its size does not, and its angle
    /// becomes its angle relative to the group - the solver runs unchanged in that space, and every
    /// change is turned back. Any fixed point would do as the centre of the turn; moving a shape never
    /// shifts its siblings, however the group's frame re-fits.
    /// </para>
    /// <para>
    /// Some requests mean nothing along a turned group's axes and are refused rather than solved:
    /// the slide, its margins and the content placeholder are rectangles along the slide's axes, and
    /// PowerPoint's text bounds are a slide-axis box that stops describing the text once turned.
    /// </para>
    /// </remarks>
    public static class GroupSpace
    {
        /// <summary>
        /// Solves <paramref name="request"/> for shapes inside a group turned by
        /// <paramref name="groupAngle"/> degrees, whose frame is <see cref="SlideMetrics.GroupBounds"/>.
        /// Changes come back in slide space, ready to write. An angle of zero is solved directly.
        /// </summary>
        public static SolveResult Solve(
            AlignRequest request, IReadOnlyList<ShapeSnapshot> shapes, SlideMetrics slide, double groupAngle)
        {
            if (request is null) throw new ArgumentNullException(nameof(request));
            if (shapes is null) throw new ArgumentNullException(nameof(shapes));
            if (slide is null) throw new ArgumentNullException(nameof(slide));

            if (GeometryChange.SameAngle(groupAngle, 0)) return AlignSolver.Solve(request, shapes, slide);

            if (!slide.GroupBounds.HasValue)
            {
                throw new ArgumentException("A rotated group needs its frame in SlideMetrics.GroupBounds.", nameof(slide));
            }

            var refusal = Refusal(request);
            if (refusal != null) return SolveResult.Refused(refusal);

            var group = slide.GroupBounds.Value;
            var pivot = new PointD(group.CentreX, group.CentreY);

            var turned = shapes.Select(s => ToGroup(s, pivot, groupAngle)).ToList();
            var groupSlide = new SlideMetrics(slide.Width, slide.Height, slide.Margin, groupBounds: group);

            var result = AlignSolver.Solve(request, turned, groupSlide);
            if (!result.Succeeded) return result;

            var originals = shapes.ToDictionary(s => s.Key);
            var changes = result.Changes
                .Select(c => FromGroup(c, originals[c.Key], pivot, groupAngle))
                .ToList();

            return result.Notable
                ? SolveResult.OkWithNotice(changes, result.Diagnostics.ToArray())
                : SolveResult.Ok(changes, result.Diagnostics.ToArray());
        }

        /// <summary>
        /// Solves a duplicate for shapes inside a group turned by <paramref name="groupAngle"/>
        /// degrees: X and Y step along the group's own axes, and every placement comes back in slide
        /// space. An angle of zero is solved directly.
        /// </summary>
        public static DuplicateResult SolveDuplicate(
            DuplicateRequest request, IReadOnlyList<ShapeSnapshot> shapes, SlideMetrics slide, double groupAngle)
        {
            if (request is null) throw new ArgumentNullException(nameof(request));
            if (shapes is null) throw new ArgumentNullException(nameof(shapes));
            if (slide is null) throw new ArgumentNullException(nameof(slide));

            if (GeometryChange.SameAngle(groupAngle, 0)) return DuplicateSolver.Solve(request, shapes, slide);

            if (!slide.GroupBounds.HasValue)
            {
                throw new ArgumentException("A rotated group needs its frame in SlideMetrics.GroupBounds.", nameof(slide));
            }

            if (request.Pivot == DuplicatePivot.SlideCentre)
            {
                return DuplicateResult.Refused(
                    "Inside a rotated group, copies step along the group's own edges, so the slide's centre is no pivot for them. Use Own centre, Selection centre or Anchor centre.");
            }

            var group = slide.GroupBounds.Value;
            var pivot = new PointD(group.CentreX, group.CentreY);
            var turned = shapes.Select(s => ToGroup(s, pivot, groupAngle)).ToList();

            var result = DuplicateSolver.Solve(request, turned, slide, checkSlide: false);
            if (!result.Succeeded) return result;

            var copies = result.Copies
                .Select(copy => new DuplicateCopy(
                    copy.Index,
                    copy.Placements.Select(p => FromGroup(p, pivot, groupAngle)).ToList()))
                .ToList();

            var diagnostics = result.Diagnostics.ToList();
            var notice = DuplicateSolver.OffSlideNotice(copies, slide);
            if (notice != null) diagnostics.Add(notice);

            return result.Notable || notice != null
                ? DuplicateResult.OkWithNotice(copies, diagnostics.ToArray())
                : DuplicateResult.Ok(copies, diagnostics.ToArray());
        }

        /// <summary>
        /// Tidies shapes inside a group turned by <paramref name="groupAngle"/> degrees, lining them
        /// up and spacing them along the group's own axes. An angle of zero is solved directly.
        /// </summary>
        public static SolveResult SolveTidy(
            TidyRequest request, IReadOnlyList<ShapeSnapshot> shapes, RectD? groupBounds, double groupAngle)
        {
            if (request is null) throw new ArgumentNullException(nameof(request));
            if (shapes is null) throw new ArgumentNullException(nameof(shapes));

            if (GeometryChange.SameAngle(groupAngle, 0)) return TidySolver.Solve(request, shapes);

            if (!groupBounds.HasValue)
            {
                throw new ArgumentException("A rotated group needs its frame.", nameof(groupBounds));
            }

            if (request.BoundsModel == BoundsModel.TextBounds)
            {
                return SolveResult.Refused(
                    "Inside a rotated group, PowerPoint does not report where the text sits along the group's edges. Use Measure = Shape frame or Visual bounds.");
            }

            var pivot = new PointD(groupBounds.Value.CentreX, groupBounds.Value.CentreY);
            var turned = shapes.Select(s => ToGroup(s, pivot, groupAngle)).ToList();

            var result = TidySolver.Solve(request, turned);
            if (!result.Succeeded) return result;

            var originals = shapes.ToDictionary(s => s.Key);
            var changes = result.Changes
                .Select(c => FromGroup(c, originals[c.Key], pivot, groupAngle))
                .ToList();

            return result.Notable
                ? SolveResult.OkWithNotice(changes, result.Diagnostics.ToArray())
                : SolveResult.Ok(changes, result.Diagnostics.ToArray());
        }

        private static string? Refusal(AlignRequest request)
        {
            switch (request.Reference)
            {
                case ReferenceTarget.Slide:
                case ReferenceTarget.SlideMargins:
                case ReferenceTarget.PlaceholderBounds:
                    return "Inside a rotated group, shapes line up along the group's own edges, and the slide's edges run a different way. Use Reference = Group, Anchor or Selection bounds.";
            }

            if (request.BoundsModel == BoundsModel.TextBounds)
            {
                return "Inside a rotated group, PowerPoint does not report where the text sits along the group's edges. Use Measure = Shape frame or Visual bounds.";
            }

            return null;
        }

        /// <summary>A shape as seen along the group's axes: centre turned back, angle made relative.</summary>
        internal static ShapeSnapshot ToGroup(ShapeSnapshot shape, PointD pivot, double groupAngle)
        {
            var centre = Turn(new PointD(shape.Frame.CentreX, shape.Frame.CentreY), pivot, -groupAngle);
            return new ShapeSnapshot(
                shape.Key,
                Around(centre, shape.Frame),
                GeometryChange.NormaliseAngle(shape.Rotation - groupAngle),
                shape.FlipH,
                shape.FlipV,
                textBounds: null,
                shape.IsGroup,
                shape.IsPlaceholder,
                shape.Name,
                shape.IsConnector);
        }

        /// <summary>A change solved along the group's axes, turned back into slide space.</summary>
        internal static GeometryChange FromGroup(GeometryChange change, ShapeSnapshot original, PointD pivot, double groupAngle)
        {
            // Nothing moved: hand back the original exactly, rather than a frame a rounding error away.
            if (change.IsNoOp) return new GeometryChange(original.Key, original.Frame, original.Frame);

            var centre = Turn(new PointD(change.NewFrame.CentreX, change.NewFrame.CentreY), pivot, groupAngle);
            var frame = Around(centre, change.NewFrame);

            return change.NewRotation.HasValue
                ? new GeometryChange(
                    original.Key, original.Frame, frame,
                    original.Rotation, GeometryChange.NormaliseAngle(change.NewRotation.Value + groupAngle))
                : new GeometryChange(original.Key, original.Frame, frame);
        }

        /// <summary>A copy's placement solved along the group's axes, turned back into slide space.</summary>
        private static ShapePlacement FromGroup(ShapePlacement placement, PointD pivot, double groupAngle)
        {
            var centre = Turn(new PointD(placement.Frame.CentreX, placement.Frame.CentreY), pivot, groupAngle);
            return new ShapePlacement(
                placement.Source,
                Around(centre, placement.Frame),
                GeometryChange.NormaliseAngle(placement.Rotation + groupAngle));
        }

        /// <summary>A point turned clockwise by <paramref name="degrees"/> about <paramref name="pivot"/>.</summary>
        /// <remarks>Clockwise on screen, as PowerPoint's angles are, because y grows downwards.</remarks>
        internal static PointD Turn(PointD point, PointD pivot, double degrees)
        {
            var radians = degrees * Math.PI / 180.0;
            var cos = Math.Cos(radians);
            var sin = Math.Sin(radians);
            var dx = point.X - pivot.X;
            var dy = point.Y - pivot.Y;
            return new PointD(pivot.X + dx * cos - dy * sin, pivot.Y + dx * sin + dy * cos);
        }

        private static RectD Around(PointD centre, RectD size) =>
            new RectD(centre.X - size.Width / 2, centre.Y - size.Height / 2, size.Width, size.Height);
    }
}
