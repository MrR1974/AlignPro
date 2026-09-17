namespace AlignPro.Geometry.Tests;

/// <summary>
/// Two problems found by hand on the sample deck: match-size refusing a group it was never going to
/// resize, and grid packing everything into one shape's footprint.
/// </summary>
public class GroupAnchorAndReferenceTests
{
    private const double Tolerance = 1e-9;

    // -- a group is a fine anchor ------------------------------------------------------------------

    [Fact]
    public void MatchSize_AcceptsAGroupAsTheAnchor()
    {
        // The anchor is only measured, never resized, so the objection to resizing groups does not
        // apply to it. Refusing here was simply wrong.
        var result = AlignSolver.Solve(
            new AlignRequest(AlignVerb.MatchBoth, ReferenceTarget.Anchor, anchor: Make.Key(1)),
            new[]
            {
                Make.Shape(1, 0, 0, 200, 100, isGroup: true),
                Make.Shape(2, 300, 300, 50, 40)
            },
            Make.Slide());

        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        Assert.Equal(200, result.FrameOf(2).Width, Tolerance);
        Assert.Equal(100, result.FrameOf(2).Height, Tolerance);
    }

    [Fact]
    public void MatchSize_WithAGroupAnchor_StillLeavesTheGroupItselfAlone()
    {
        var result = AlignSolver.Solve(
            new AlignRequest(AlignVerb.MatchBoth, ReferenceTarget.Anchor, anchor: Make.Key(1)),
            new[]
            {
                Make.Shape(1, 0, 0, 200, 100, isGroup: true),
                Make.Shape(2, 300, 300, 50, 40)
            },
            Make.Slide());

        Assert.False(result.Touched(1));
    }

    [Fact]
    public void MatchSize_WithAGroupAnchor_StillSkipsGroupTargets()
    {
        // Anchoring on a group is fine; resizing one is not. Both rules apply at once.
        var result = AlignSolver.Solve(
            new AlignRequest(AlignVerb.MatchBoth, ReferenceTarget.Anchor, anchor: Make.Key(1)),
            new[]
            {
                Make.Shape(1, 0, 0, 200, 100, isGroup: true),
                Make.Shape(2, 300, 300, 50, 40, isGroup: true),
                Make.Shape(3, 500, 300, 50, 40)
            },
            Make.Slide());

        Assert.True(result.Succeeded);
        Assert.False(result.Touched(2));
        Assert.True(result.Touched(3));
        Assert.True(result.Notable);
    }

    // -- grid must not pack a selection into one shape ---------------------------------------------

    private static ShapeSnapshot[] NineCircles() => Enumerable.Range(0, 9)
        .Select(i => Make.Shape(i + 1, 100 + i * 90, 150 + (i % 3) * 120, 80, 80))
        .ToArray();

    [Fact]
    public void Grid_WithAnAnchorReference_UsesTheSelectionInstead()
    {
        // Reported symptom: with Reference left on Anchor from an earlier operation, nine 80pt circles
        // collapsed into an 80pt corner - the anchor shape's own footprint.
        var shapes = NineCircles();
        var result = AlignSolver.Solve(
            new AlignRequest(AlignVerb.GridArrange, ReferenceTarget.Anchor, anchor: Make.Key(9)),
            shapes,
            Make.Slide());

        Assert.True(result.Succeeded);
        Assert.True(result.Notable);
        Assert.Contains(result.Diagnostics, d => d.Contains("anchor", StringComparison.OrdinalIgnoreCase));

        // The grid should span the selection, not one 80pt shape.
        var xs = result.Changes.Select(c => c.NewFrame.CentreX).ToList();
        Assert.True(xs.Max() - xs.Min() > 400, $"grid spanned only {xs.Max() - xs.Min():F0}pt");
    }

    [Fact]
    public void Grid_WarnsWhenShapesAreBiggerThanTheirCells()
    {
        // Grid only repositions, so shapes too big for their cells overlap. The user should be told
        // rather than left looking at a pile.
        var slide = Make.Slide(placeholder: new RectD(0, 0, 120, 120));
        var result = AlignSolver.Solve(
            new AlignRequest(AlignVerb.GridArrange, ReferenceTarget.PlaceholderBounds),
            NineCircles(),
            slide);

        Assert.True(result.Succeeded);
        Assert.True(result.Notable);
        Assert.Contains(result.Diagnostics, d => d.Contains("overlap", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Grid_DoesNotWarnWhenTheCellsAreRoomy()
    {
        var result = AlignSolver.Solve(
            new AlignRequest(AlignVerb.GridArrange, ReferenceTarget.Slide),
            NineCircles(),
            Make.Slide());

        Assert.True(result.Succeeded);
        Assert.DoesNotContain(result.Diagnostics, d => d.Contains("overlap", StringComparison.OrdinalIgnoreCase));
        Assert.False(result.Notable);
    }
}
