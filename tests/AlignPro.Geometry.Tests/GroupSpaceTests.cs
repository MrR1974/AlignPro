namespace AlignPro.Geometry.Tests;

/// <summary>
/// Shapes inside a rotated group line up along the group's own axes. Each test lays shapes out in the
/// group's coordinates, turns them onto the slide as PowerPoint would report them (probe 14), solves,
/// and turns the answer back to check it - with its own turn, not GroupSpace's.
/// </summary>
public class GroupSpaceTests
{
    private const double Tolerance = 1e-6;
    private const double Angle = 30;

    // The group's frame, whose centre the group turns about.
    private static readonly RectD Group = new(100, 100, 400, 200);
    private static readonly (double X, double Y) Pivot = (300, 200);

    private static (double X, double Y) Turn((double X, double Y) p, double degrees)
    {
        var r = degrees * Math.PI / 180;
        var dx = p.X - Pivot.X;
        var dy = p.Y - Pivot.Y;
        return (Pivot.X + dx * Math.Cos(r) - dy * Math.Sin(r), Pivot.Y + dx * Math.Sin(r) + dy * Math.Cos(r));
    }

    /// <summary>A shape at (x, y, w, h) along the group's axes, as the slide reports it.</summary>
    private static ShapeSnapshot InGroup(int id, double x, double y, double w, double h, double relativeRotation = 0)
    {
        var (cx, cy) = Turn((x + w / 2, y + h / 2), Angle);
        return Make.Shape(id, cx - w / 2, cy - h / 2, w, h, rotation: relativeRotation + Angle);
    }

    /// <summary>Where a solved shape's frame sits along the group's axes.</summary>
    private static RectD AlongGroup(SolveResult result, int id)
    {
        var frame = result.FrameOf(id);
        var (cx, cy) = Turn((frame.CentreX, frame.CentreY), -Angle);
        return new RectD(cx - frame.Width / 2, cy - frame.Height / 2, frame.Width, frame.Height);
    }

    private static SolveResult Solve(AlignRequest request, params ShapeSnapshot[] shapes) =>
        GroupSpace.Solve(request, shapes, Make.Slide(group: Group), Angle);

    [Fact]
    public void AnUnrotatedGroup_SolvesExactlyAsTheSlideWould()
    {
        var shapes = new[] { Make.Shape(1, 120, 150, 60, 40), Make.Shape(2, 250, 220, 60, 40) };
        var request = new AlignRequest(AlignVerb.AlignLeft, ReferenceTarget.SelectionBounds);

        var direct = AlignSolver.Solve(request, shapes, Make.Slide(group: Group));
        var viaGroup = GroupSpace.Solve(request, shapes, Make.Slide(group: Group), 0);

        Assert.Equal(direct.FrameOf(2), viaGroup.FrameOf(2));
    }

    [Fact]
    public void AlignLeft_LinesShapesUpAlongTheGroupsAxis_NotTheSlides()
    {
        var result = Solve(
            new AlignRequest(AlignVerb.AlignLeft, ReferenceTarget.SelectionBounds),
            InGroup(1, 120, 150, 60, 40), InGroup(2, 250, 220, 60, 40));

        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        var moved = AlongGroup(result, 2);
        Assert.Equal(120, moved.Left, Tolerance);
        Assert.Equal(220, moved.Top, Tolerance);

        // On the slide that is a move down-left as well as left: both coordinates change.
        var before = InGroup(2, 250, 220, 60, 40).Frame;
        Assert.NotEqual(before.Top, result.FrameOf(2).Top, 3);
    }

    [Fact]
    public void AlignLeft_LeavesTheAngleAlone()
    {
        var result = Solve(
            new AlignRequest(AlignVerb.AlignLeft, ReferenceTarget.SelectionBounds),
            InGroup(1, 120, 150, 60, 40), InGroup(2, 250, 220, 60, 40));

        var change = result.Changes.Single(c => c.Key == Make.Key(2));
        Assert.Null(change.NewRotation);
    }

    [Fact]
    public void ShapesAlreadyFlushAlongTheGroup_DoNotMove()
    {
        var result = Solve(
            new AlignRequest(AlignVerb.AlignLeft, ReferenceTarget.SelectionBounds),
            InGroup(1, 120, 150, 60, 40), InGroup(2, 120, 220, 80, 40));

        Assert.Empty(result.EffectiveChanges);
    }

    [Fact]
    public void Distribute_SpacesEvenlyAlongTheGroupsAxis()
    {
        var result = Solve(
            new AlignRequest(AlignVerb.DistributeH, ReferenceTarget.SelectionBounds),
            InGroup(1, 120, 150, 40, 40), InGroup(2, 180, 200, 40, 40), InGroup(3, 400, 160, 40, 40));

        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        var a = AlongGroup(result, 1);
        var b = AlongGroup(result, 2);
        var c = AlongGroup(result, 3);
        Assert.Equal(b.Left - a.Right, c.Left - b.Right, Tolerance);
        Assert.Equal(200, b.Top, Tolerance);   // distributing across does not move it down the group
    }

    [Fact]
    public void AlignToGroup_UsesTheGroupsOwnEdge()
    {
        var result = Solve(
            new AlignRequest(AlignVerb.AlignTop, ReferenceTarget.Group),
            InGroup(1, 150, 180, 60, 40));

        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        Assert.Equal(100, AlongGroup(result, 1).Top, Tolerance);
    }

    [Fact]
    public void VisualBounds_AreMeasuredAlongTheGroup()
    {
        // Both shapes sit square to the group, so along its axes their visual bounds are their frames,
        // even though on the slide each is turned by 30 degrees.
        var result = Solve(
            new AlignRequest(AlignVerb.AlignLeft, ReferenceTarget.SelectionBounds, BoundsModel.VisualBounds),
            InGroup(1, 120, 150, 60, 40), InGroup(2, 250, 220, 100, 40));

        Assert.Equal(120, AlongGroup(result, 2).Left, Tolerance);
    }

    [Fact]
    public void MatchRotation_StillTakesTheAnchorsAngle()
    {
        var result = Solve(
            new AlignRequest(AlignVerb.MatchRotation, ReferenceTarget.Anchor, anchor: Make.Key(2)),
            InGroup(1, 120, 150, 60, 40, relativeRotation: 10), InGroup(2, 250, 220, 60, 40, relativeRotation: 40));

        var change = result.Changes.Single(c => c.Key == Make.Key(1));
        Assert.Equal(70, change.NewRotation!.Value, Tolerance);
        Assert.Equal(40, change.OldRotation!.Value, Tolerance);
    }

    [Fact]
    public void MatchSize_KeepsSizesAndCentresAlongTheGroup()
    {
        var result = Solve(
            new AlignRequest(AlignVerb.MatchWidth, ReferenceTarget.Anchor, anchor: Make.Key(2), resizeOrigin: ResizeOrigin.Centre),
            InGroup(1, 120, 150, 60, 40), InGroup(2, 250, 220, 100, 40));

        var resized = AlongGroup(result, 1);
        Assert.Equal(100, resized.Width, Tolerance);
        Assert.Equal(150, resized.CentreX, Tolerance);
        Assert.Equal(170, resized.CentreY, Tolerance);
    }

    [Theory]
    [InlineData(ReferenceTarget.Slide)]
    [InlineData(ReferenceTarget.SlideMargins)]
    [InlineData(ReferenceTarget.PlaceholderBounds)]
    public void SlideAxisReferences_AreRefused(ReferenceTarget reference)
    {
        var result = Solve(new AlignRequest(AlignVerb.AlignLeft, reference), InGroup(1, 120, 150, 60, 40));

        Assert.False(result.Succeeded);
        Assert.Contains(result.Diagnostics, d => d.Contains("rotated group", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void TextBounds_AreRefused()
    {
        var result = Solve(
            new AlignRequest(AlignVerb.AlignLeft, ReferenceTarget.SelectionBounds, BoundsModel.TextBounds),
            InGroup(1, 120, 150, 60, 40), InGroup(2, 250, 220, 60, 40));

        Assert.False(result.Succeeded);
    }

    // -- Duplicate -------------------------------------------------------------------------------

    private static DuplicateResult Duplicate(DuplicateRequest request, params ShapeSnapshot[] shapes) =>
        GroupSpace.SolveDuplicate(request, shapes, Make.Slide(group: Group), Angle);

    private static RectD AlongGroup(ShapePlacement placement)
    {
        var (cx, cy) = Turn((placement.Frame.CentreX, placement.Frame.CentreY), -Angle);
        return new RectD(cx - placement.Frame.Width / 2, cy - placement.Frame.Height / 2, placement.Frame.Width, placement.Frame.Height);
    }

    [Fact]
    public void Duplicate_InAnUnrotatedGroup_SolvesExactlyAsTheSlideWould()
    {
        var shapes = new[] { Make.Shape(1, 120, 150, 60, 40) };
        var request = new DuplicateRequest(50, 0, 0, 2, DuplicatePivot.OwnCentre, true, null);

        var direct = DuplicateSolver.Solve(request, shapes, Make.Slide(group: Group));
        var viaGroup = GroupSpace.SolveDuplicate(request, shapes, Make.Slide(group: Group), 0);

        Assert.Equal(direct.Copies[1].Placements[0].Frame, viaGroup.Copies[1].Placements[0].Frame);
    }

    [Fact]
    public void Duplicate_StepsAlongTheGroupsAxis()
    {
        var result = Duplicate(
            new DuplicateRequest(100, 0, 0, 2, DuplicatePivot.OwnCentre, true, null),
            InGroup(1, 120, 150, 60, 40, relativeRotation: 10));

        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        var first = AlongGroup(result.Copies[0].Placements[0]);
        var second = AlongGroup(result.Copies[1].Placements[0]);
        Assert.Equal(220, first.Left, Tolerance);
        Assert.Equal(150, first.Top, Tolerance);
        Assert.Equal(320, second.Left, Tolerance);

        // The copy keeps the original's angle on the slide: 10 within the group, 40 in all.
        Assert.Equal(40, result.Copies[0].Placements[0].Rotation, Tolerance);
    }

    [Fact]
    public void Duplicate_AboutTheSlideCentre_IsRefused()
    {
        var result = Duplicate(
            new DuplicateRequest(0, 0, 30, 3, DuplicatePivot.SlideCentre, true, null),
            InGroup(1, 120, 150, 60, 40));

        Assert.False(result.Succeeded);
        Assert.Contains(result.Diagnostics, d => d.Contains("rotated group", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Duplicate_ChecksTheSlideInSlideCoordinates()
    {
        // Far along the group's axis, which runs down-right on the slide: off the slide entirely.
        var result = Duplicate(
            new DuplicateRequest(2000, 0, 0, 1, DuplicatePivot.OwnCentre, true, null),
            InGroup(1, 120, 150, 60, 40));

        Assert.True(result.Succeeded);
        Assert.True(result.Notable);
        Assert.Contains(result.Diagnostics, d => d.Contains("off the slide", StringComparison.OrdinalIgnoreCase));
    }
}
