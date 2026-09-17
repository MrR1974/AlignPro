using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace AlignPro.Geometry
{
    /// <summary>
    /// Turns a request plus a set of shape snapshots into a list of new frames. Pure: no Office
    /// interop, no mutation of its inputs, no I/O.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The central invariant is that the solver reasons in whichever bounds space the request asked
    /// for, but always emits <em>frame</em> coordinates, because the frame is the only thing
    /// PowerPoint lets us write. Translation is rigid, so a delta in visual or text space is the same
    /// delta in frame space, and the align and distribute verbs get rotation- and text-awareness for
    /// free.
    /// </para>
    /// <para>
    /// The match-size verbs deliberately work in frame space whatever the request says. Matching a
    /// rotated shape's visual width is ill-posed - at 90 degrees the visual width is driven entirely
    /// by the frame's height, and the inverse has no solution for some targets - and "make these the
    /// same size as that one" almost always means the same frame size anyway.
    /// </para>
    /// </remarks>
    public static class AlignSolver
    {
        public static SolveResult Solve(AlignRequest request, IReadOnlyList<ShapeSnapshot> shapes, SlideMetrics slide)
        {
            if (request is null) throw new ArgumentNullException(nameof(request));
            if (shapes is null) throw new ArgumentNullException(nameof(shapes));
            if (slide is null) throw new ArgumentNullException(nameof(slide));

            if (shapes.Count == 0) return SolveResult.Refused("Nothing is selected.");

            var diagnostics = new List<string>();
            var bounds = BuildBoundsLookup(request.BoundsModel, shapes, diagnostics);

            switch (request.Verb)
            {
                case AlignVerb.AlignLeft:
                case AlignVerb.AlignRight:
                case AlignVerb.AlignTop:
                case AlignVerb.AlignBottom:
                case AlignVerb.AlignCentreH:
                case AlignVerb.AlignCentreV:
                    return SolveAlign(request, shapes, slide, bounds, diagnostics);

                case AlignVerb.DistributeH:
                case AlignVerb.DistributeV:
                    return SolveDistribute(request, shapes, slide, bounds, diagnostics);

                case AlignVerb.MatchWidth:
                case AlignVerb.MatchHeight:
                case AlignVerb.MatchBoth:
                    return SolveMatchSize(request, shapes, diagnostics);

                case AlignVerb.GridArrange:
                    return SolveGrid(request, shapes, slide, bounds, diagnostics);

                default:
                    return SolveResult.Refused($"Unsupported verb '{request.Verb}'.");
            }
        }

        // -----------------------------------------------------------------------------------------
        // Bounds models
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Resolves each shape's rectangle in the requested bounds space, once, up front. Text bounds
        /// fall back to the frame for shapes that hold no text, which is noted rather than silent.
        /// </summary>
        private static Dictionary<ShapeKey, RectD> BuildBoundsLookup(
            BoundsModel model, IReadOnlyList<ShapeSnapshot> shapes, List<string> diagnostics)
        {
            var lookup = new Dictionary<ShapeKey, RectD>(shapes.Count);
            var textFallbacks = 0;

            foreach (var shape in shapes)
            {
                RectD rect;
                switch (model)
                {
                    case BoundsModel.VisualBounds:
                        rect = shape.VisualBounds;
                        break;

                    case BoundsModel.TextBounds:
                        if (shape.TextBounds.HasValue)
                        {
                            rect = shape.TextBounds.Value;
                        }
                        else
                        {
                            rect = shape.Frame;
                            textFallbacks++;
                        }
                        break;

                    default:
                        rect = shape.Frame;
                        break;
                }

                lookup[shape.Key] = rect;
            }

            if (textFallbacks > 0)
            {
                diagnostics.Add(textFallbacks == 1
                    ? "1 shape has no text, so its frame was used instead."
                    : $"{textFallbacks} shapes have no text, so their frames were used instead.");
            }

            return lookup;
        }

        /// <summary>
        /// Resolves what the verb measures against. Returns null when the request names a reference
        /// the caller could not supply, with the reason in <paramref name="refusal"/>.
        /// </summary>
        private static RectD? ResolveReference(
            AlignRequest request,
            IReadOnlyList<ShapeSnapshot> shapes,
            SlideMetrics slide,
            IReadOnlyDictionary<ShapeKey, RectD> bounds,
            out string? refusal)
        {
            refusal = null;

            switch (request.Reference)
            {
                case ReferenceTarget.Slide:
                    return slide.Bounds;

                case ReferenceTarget.SlideMargins:
                    return slide.Bounds.Deflate(slide.Margin);

                case ReferenceTarget.PlaceholderBounds:
                    if (!slide.PlaceholderBounds.HasValue)
                    {
                        refusal = "This slide's layout has no body placeholder to align to.";
                        return null;
                    }
                    return slide.PlaceholderBounds.Value;

                case ReferenceTarget.Anchor:
                    var anchor = FindAnchor(request, shapes, out refusal);
                    if (anchor is null) return null;
                    return bounds[anchor.Key];

                default:
                    return UnionOf(shapes, bounds);
            }
        }

        private static ShapeSnapshot? FindAnchor(
            AlignRequest request, IReadOnlyList<ShapeSnapshot> shapes, out string? refusal)
        {
            refusal = null;

            if (!request.Anchor.HasValue)
            {
                refusal = "No anchor shape was given. Select the shape to align to last, or pin one.";
                return null;
            }

            var anchor = shapes.FirstOrDefault(s => s.Key == request.Anchor.Value);
            if (anchor is null)
            {
                refusal = "The anchor shape is not part of the selection.";
                return null;
            }

            return anchor;
        }

        private static RectD UnionOf(IReadOnlyList<ShapeSnapshot> shapes, IReadOnlyDictionary<ShapeKey, RectD> bounds)
        {
            var union = bounds[shapes[0].Key];
            for (var i = 1; i < shapes.Count; i++) union = union.Union(bounds[shapes[i].Key]);
            return union;
        }

        // -----------------------------------------------------------------------------------------
        // Align
        // -----------------------------------------------------------------------------------------

        private static SolveResult SolveAlign(
            AlignRequest request,
            IReadOnlyList<ShapeSnapshot> shapes,
            SlideMetrics slide,
            IReadOnlyDictionary<ShapeKey, RectD> bounds,
            List<string> diagnostics)
        {
            var isSelfReferential = request.Reference == ReferenceTarget.SelectionBounds ||
                                    request.Reference == ReferenceTarget.Anchor;
            if (isSelfReferential && shapes.Count < 2)
            {
                return SolveResult.Refused("Select at least two shapes, or align to the slide instead.");
            }

            var reference = ResolveReference(request, shapes, slide, bounds, out var refusal);
            if (!reference.HasValue) return SolveResult.Refused(refusal!);
            var refRect = reference.Value;

            // The anchor defines the line; it must not be moved onto itself.
            ShapeKey? anchorKey = null;
            if (request.Reference == ReferenceTarget.Anchor)
            {
                anchorKey = FindAnchor(request, shapes, out _)!.Key;
            }

            var changes = new List<GeometryChange>(shapes.Count);

            foreach (var shape in shapes)
            {
                if (anchorKey.HasValue && shape.Key == anchorKey.Value) continue;

                var b = bounds[shape.Key];
                double dx = 0, dy = 0;

                switch (request.Verb)
                {
                    case AlignVerb.AlignLeft:
                        dx = refRect.Left - b.Left;
                        break;
                    case AlignVerb.AlignRight:
                        dx = refRect.Right - b.Right;
                        break;
                    case AlignVerb.AlignCentreH:
                        dx = refRect.CentreX - b.CentreX;
                        break;
                    case AlignVerb.AlignTop:
                        dy = refRect.Top - b.Top;
                        break;
                    case AlignVerb.AlignBottom:
                        dy = refRect.Bottom - b.Bottom;
                        break;
                    case AlignVerb.AlignCentreV:
                        dy = refRect.CentreY - b.CentreY;
                        break;
                }

                // Rigid translation: the delta measured in bounds space applies unchanged to the frame.
                changes.Add(new GeometryChange(shape.Key, shape.Frame, shape.Frame.Offset(dx, dy)));
            }

            return SolveResult.Ok(changes, diagnostics.ToArray());
        }

        // -----------------------------------------------------------------------------------------
        // Distribute
        // -----------------------------------------------------------------------------------------

        private static SolveResult SolveDistribute(
            AlignRequest request,
            IReadOnlyList<ShapeSnapshot> shapes,
            SlideMetrics slide,
            IReadOnlyDictionary<ShapeKey, RectD> bounds,
            List<string> diagnostics)
        {
            var horizontal = request.Verb == AlignVerb.DistributeH;
            var hasExact = request.ExactSpacing.HasValue;

            // With no exact spacing we can only redistribute space that already exists between the
            // outermost shapes, and that needs a middle shape to move.
            var minimum = hasExact ? 2 : 3;
            if (shapes.Count < minimum)
            {
                return SolveResult.Refused(hasExact
                    ? "Select at least two shapes to space."
                    : "Select at least three shapes to distribute, or set an exact spacing instead.");
            }

            // A span from the reference rect means "spread these across the slide / content area";
            // otherwise the selection's own extremes hold and the interior is evened out.
            var spanFromReference = request.Reference == ReferenceTarget.Slide ||
                                    request.Reference == ReferenceTarget.SlideMargins ||
                                    request.Reference == ReferenceTarget.PlaceholderBounds;

            if (request.Reference == ReferenceTarget.Anchor)
            {
                diagnostics.Add("An anchor does not define a span, so the selection's own extent was used.");
            }

            var ordered = shapes
                .OrderBy(s => horizontal ? bounds[s.Key].Left : bounds[s.Key].Top)
                .ThenBy(s => horizontal ? bounds[s.Key].CentreX : bounds[s.Key].CentreY)
                .ThenBy(s => s.Key.ShapeId)
                .ToList();

            double spanStart, spanEnd;
            if (spanFromReference)
            {
                var reference = ResolveReference(request, shapes, slide, bounds, out var refusal);
                if (!reference.HasValue) return SolveResult.Refused(refusal!);
                spanStart = horizontal ? reference.Value.Left : reference.Value.Top;
                spanEnd = horizontal ? reference.Value.Right : reference.Value.Bottom;
            }
            else
            {
                var first = bounds[ordered[0].Key];
                var last = bounds[ordered[ordered.Count - 1].Key];
                spanStart = horizontal ? first.Left : first.Top;
                spanEnd = horizontal ? last.Right : last.Bottom;
            }

            var changes = new List<GeometryChange>(ordered.Count);

            if (request.DistributeMode == DistributeMode.Centre)
            {
                // Equalise centre-to-centre pitch.
                var firstCentre = Centre(bounds[ordered[0].Key], horizontal);
                double pitch;
                if (hasExact)
                {
                    pitch = request.ExactSpacing!.Value;
                }
                else
                {
                    var lastCentre = spanFromReference
                        ? spanEnd - Extent(bounds[ordered[ordered.Count - 1].Key], horizontal) / 2
                        : Centre(bounds[ordered[ordered.Count - 1].Key], horizontal);
                    if (spanFromReference) firstCentre = spanStart + Extent(bounds[ordered[0].Key], horizontal) / 2;
                    pitch = (lastCentre - firstCentre) / (ordered.Count - 1);
                }

                for (var i = 0; i < ordered.Count; i++)
                {
                    var shape = ordered[i];
                    var b = bounds[shape.Key];
                    var targetCentre = firstCentre + pitch * i;
                    var delta = targetCentre - Centre(b, horizontal);
                    changes.Add(Translate(shape, horizontal, delta));
                }
            }
            else
            {
                // Equalise edge-to-edge gaps.
                double gap;
                if (hasExact)
                {
                    gap = request.ExactSpacing!.Value;
                }
                else
                {
                    var occupied = ordered.Sum(s => Extent(bounds[s.Key], horizontal));
                    gap = (spanEnd - spanStart - occupied) / (ordered.Count - 1);
                }

                var cursor = spanStart;
                foreach (var shape in ordered)
                {
                    var b = bounds[shape.Key];
                    var delta = cursor - (horizontal ? b.Left : b.Top);
                    changes.Add(Translate(shape, horizontal, delta));
                    cursor += Extent(b, horizontal) + gap;
                }

                if (!hasExact && gap < 0)
                {
                    diagnostics.Add(string.Format(
                        CultureInfo.CurrentCulture,
                        "The shapes do not fit in the available span, so they overlap by {0:F1}pt each.",
                        -gap));
                }
            }

            return SolveResult.Ok(changes, diagnostics.ToArray());
        }

        // -----------------------------------------------------------------------------------------
        // Match size
        // -----------------------------------------------------------------------------------------

        private static SolveResult SolveMatchSize(
            AlignRequest request, IReadOnlyList<ShapeSnapshot> shapes, List<string> diagnostics)
        {
            if (shapes.Count < 2) return SolveResult.Refused("Select at least two shapes to match sizes.");

            var anchor = FindAnchor(request, shapes, out var refusal);
            if (anchor is null) return SolveResult.Refused(refusal!);

            if (anchor.IsGroup && !request.AllowGroupResize)
            {
                return SolveResult.Refused(
                    "The anchor is a group, and matching its size would rescale the spacing inside it.");
            }

            var matchWidth = request.Verb == AlignVerb.MatchWidth || request.Verb == AlignVerb.MatchBoth;
            var matchHeight = request.Verb == AlignVerb.MatchHeight || request.Verb == AlignVerb.MatchBoth;

            var changes = new List<GeometryChange>(shapes.Count);
            var skippedGroups = 0;

            foreach (var shape in shapes)
            {
                if (shape.Key == anchor.Key) continue;

                // Scaling a group scales the gaps between its children (measured, probe 3), which
                // silently distorts a diagram. Refuse unless explicitly allowed.
                if (shape.IsGroup && !request.AllowGroupResize)
                {
                    skippedGroups++;
                    continue;
                }

                var width = matchWidth ? anchor.Frame.Width : shape.Frame.Width;
                var height = matchHeight ? anchor.Frame.Height : shape.Frame.Height;

                var newFrame = request.ResizeOrigin == ResizeOrigin.Centre
                    ? RectD.FromCentre(shape.Frame.CentreX, shape.Frame.CentreY, width, height)
                    : new RectD(shape.Frame.X, shape.Frame.Y, width, height);

                changes.Add(new GeometryChange(shape.Key, shape.Frame, newFrame));
            }

            if (skippedGroups > 0)
            {
                diagnostics.Add(skippedGroups == 1
                    ? "Skipped 1 group: resizing it would rescale the spacing inside it."
                    : $"Skipped {skippedGroups} groups: resizing them would rescale the spacing inside them.");
            }

            if (changes.Count == 0)
            {
                return SolveResult.Refused(skippedGroups > 0
                    ? "Every shape to resize is a group, and resizing a group rescales its internal spacing."
                    : "Nothing to resize.");
            }

            return SolveResult.Ok(changes, diagnostics.ToArray());
        }

        // -----------------------------------------------------------------------------------------
        // Grid
        // -----------------------------------------------------------------------------------------

        private static SolveResult SolveGrid(
            AlignRequest request,
            IReadOnlyList<ShapeSnapshot> shapes,
            SlideMetrics slide,
            IReadOnlyDictionary<ShapeKey, RectD> bounds,
            List<string> diagnostics)
        {
            if (shapes.Count < 2) return SolveResult.Refused("Select at least two shapes to arrange in a grid.");

            var reference = ResolveReference(request, shapes, slide, bounds, out var refusal);
            if (!reference.HasValue) return SolveResult.Refused(refusal!);
            var refRect = reference.Value;

            var columns = request.GridColumns ?? (int)Math.Ceiling(Math.Sqrt(shapes.Count));
            if (columns < 1) return SolveResult.Refused("A grid needs at least one column.");
            var rows = (int)Math.Ceiling(shapes.Count / (double)columns);

            var cellWidth = (refRect.Width - request.GridGapH * (columns - 1)) / columns;
            var cellHeight = (refRect.Height - request.GridGapV * (rows - 1)) / rows;

            if (cellWidth <= 0 || cellHeight <= 0)
            {
                return SolveResult.Refused("The gaps leave no room for cells. Reduce the gap or the column count.");
            }

            // Sort into reading order by current position so a rough arrangement tidies into the grid
            // nearest to where it already was, rather than being reshuffled arbitrarily.
            var ordered = shapes
                .OrderBy(s => bounds[s.Key].CentreY)
                .ThenBy(s => bounds[s.Key].CentreX)
                .ThenBy(s => s.Key.ShapeId)
                .ToList();

            var changes = new List<GeometryChange>(ordered.Count);

            for (var i = 0; i < ordered.Count; i++)
            {
                int row, column;
                if (request.GridFillOrder == GridFillOrder.ColumnMajor)
                {
                    column = i / rows;
                    row = i % rows;
                }
                else
                {
                    row = i / columns;
                    column = i % columns;
                }

                var cellCentreX = refRect.Left + column * (cellWidth + request.GridGapH) + cellWidth / 2;
                var cellCentreY = refRect.Top + row * (cellHeight + request.GridGapV) + cellHeight / 2;

                // Grid only repositions - sizes are left alone, so a group's internal spacing is safe.
                var shape = ordered[i];
                var b = bounds[shape.Key];
                var dx = cellCentreX - b.CentreX;
                var dy = cellCentreY - b.CentreY;
                changes.Add(new GeometryChange(shape.Key, shape.Frame, shape.Frame.Offset(dx, dy)));
            }

            diagnostics.Add(string.Format(CultureInfo.CurrentCulture, "Arranged into {0} x {1}.", rows, columns));
            return SolveResult.Ok(changes, diagnostics.ToArray());
        }

        // -----------------------------------------------------------------------------------------
        // Helpers
        // -----------------------------------------------------------------------------------------

        private static double Extent(RectD rect, bool horizontal) => horizontal ? rect.Width : rect.Height;

        private static double Centre(RectD rect, bool horizontal) => horizontal ? rect.CentreX : rect.CentreY;

        private static GeometryChange Translate(ShapeSnapshot shape, bool horizontal, double delta) =>
            new GeometryChange(
                shape.Key,
                shape.Frame,
                horizontal ? shape.Frame.Offset(delta, 0) : shape.Frame.Offset(0, delta));
    }
}
