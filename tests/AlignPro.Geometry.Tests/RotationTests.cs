namespace AlignPro.Geometry.Tests;

/// <summary>Rotation carried in a change, and the Match rotation verb built on it.</summary>
public class RotationTests
{
    private const double Tolerance = 1e-6;

    private static SolveResult MatchRotation(IReadOnlyList<ShapeSnapshot> shapes, int anchorId) =>
        AlignSolver.Solve(
            new AlignRequest(AlignVerb.MatchRotation, ReferenceTarget.Anchor, anchor: Make.Key(anchorId)),
            shapes,
            Make.Slide());

    private static GeometryChange ChangeOf(SolveResult result, int id) =>
        result.Changes.Single(c => c.Key == Make.Key(id));

    // -- GeometryChange -------------------------------------------------------------------------

    [Fact]
    public void AChangeWithoutRotation_LeavesTheAngleAlone()
    {
        var frame = new RectD(0, 0, 10, 10);
        var change = new GeometryChange(Make.Key(1), frame, frame.Offset(5, 0));

        Assert.Null(change.NewRotation);
        Assert.False(change.ChangesRotation);
        Assert.Null(change.Inverted().NewRotation);
    }

    [Fact]
    public void ARotationOnlyChange_IsNotANoOp()
    {
        var frame = new RectD(0, 0, 10, 10);
        var change = new GeometryChange(Make.Key(1), frame, frame, 0, 30);

        Assert.True(change.ChangesRotation);
        Assert.False(change.IsNoOp);
    }

    [Fact]
    public void Inverting_SwapsTheAnglesAsWellAsTheFrames()
    {
        var change = new GeometryChange(
            Make.Key(1), new RectD(0, 0, 10, 10), new RectD(5, 5, 10, 10), 15, 45);

        var inverse = change.Inverted();

        Assert.Equal(45, inverse.OldRotation!.Value, Tolerance);
        Assert.Equal(15, inverse.NewRotation!.Value, Tolerance);
        Assert.Equal(new RectD(0, 0, 10, 10), inverse.NewFrame);
    }

    [Theory]
    [InlineData(0, 360)]
    [InlineData(-90, 270)]
    [InlineData(720.5, 0.5)]
    [InlineData(359.99999, 0)]
    public void SameAngle_IgnoresWholeTurns(double a, double b) =>
        Assert.True(GeometryChange.SameAngle(a, b));

    [Theory]
    [InlineData(-90, 270)]
    [InlineData(400, 40)]
    [InlineData(360, 0)]
    [InlineData(-0.0, 0)]
    public void NormaliseAngle_FoldsIntoPowerPointsRange(double input, double expected) =>
        Assert.Equal(expected, GeometryChange.NormaliseAngle(input), Tolerance);

    [Fact]
    public void AngleWithoutItsPartner_IsRefused()
    {
        var frame = new RectD(0, 0, 10, 10);
        Assert.Throws<ArgumentException>(() => new GeometryChange(Make.Key(1), frame, frame, 0, null));
    }

    // -- Match rotation -------------------------------------------------------------------------

    [Fact]
    public void EveryShapeTakesTheAnchorsAngle()
    {
        var result = MatchRotation(new[]
        {
            Make.Shape(1, 0, 0, 50, 50),
            Make.Shape(2, 100, 0, 80, 40, rotation: 200),
            Make.Shape(9, 300, 0, 100, 60, rotation: 30)
        }, anchorId: 9);

        Assert.True(result.Succeeded);
        Assert.Equal(30, ChangeOf(result, 1).NewRotation!.Value, Tolerance);
        Assert.Equal(30, ChangeOf(result, 2).NewRotation!.Value, Tolerance);
    }

    [Fact]
    public void FramesStayExactlyWhereTheyWere()
    {
        var result = MatchRotation(new[]
        {
            Make.Shape(1, 10, 20, 50, 70),
            Make.Shape(9, 300, 0, 100, 60, rotation: 45)
        }, anchorId: 9);

        // PowerPoint turns a shape about its frame's centre, so the frame itself must not move.
        Assert.Equal(new RectD(10, 20, 50, 70), result.FrameOf(1));
    }

    [Fact]
    public void TheAnchorIsNotTouched()
    {
        var result = MatchRotation(new[]
        {
            Make.Shape(1, 0, 0, 50, 50),
            Make.Shape(9, 300, 0, 100, 60, rotation: 45)
        }, anchorId: 9);

        Assert.False(result.Touched(9));
    }

    [Fact]
    public void GroupsAreRotatedToo()
    {
        // Unlike resizing, turning a group is rigid - its internal spacing survives.
        var result = MatchRotation(new[]
        {
            Make.Shape(1, 0, 0, 50, 50, isGroup: true),
            Make.Shape(9, 300, 0, 100, 60, rotation: 45)
        }, anchorId: 9);

        Assert.False(result.Notable);
        Assert.Equal(45, ChangeOf(result, 1).NewRotation!.Value, Tolerance);
    }

    [Fact]
    public void FlipsAreNotPartOfTheMatch()
    {
        var result = MatchRotation(new[]
        {
            Make.Shape(1, 0, 0, 50, 50, flipH: true),
            Make.Shape(9, 300, 0, 100, 60, rotation: 45, flipV: true)
        }, anchorId: 9);

        // A change carries a frame and an angle and nothing else, so a flip cannot leak across.
        Assert.Equal(45, ChangeOf(result, 1).NewRotation!.Value, Tolerance);
    }

    [Fact]
    public void AlreadyMatchingShapesAreNoOps()
    {
        var result = MatchRotation(new[]
        {
            Make.Shape(1, 0, 0, 50, 50, rotation: 405),
            Make.Shape(9, 300, 0, 100, 60, rotation: 45)
        }, anchorId: 9);

        Assert.Empty(result.EffectiveChanges);
    }

    [Fact]
    public void UndoRestoresTheOriginalAngle()
    {
        var result = MatchRotation(new[]
        {
            Make.Shape(1, 0, 0, 50, 50, rotation: 10),
            Make.Shape(9, 300, 0, 100, 60, rotation: 45)
        }, anchorId: 9);

        var undo = new UndoManager();
        undo.Push(AlignTransaction.FromResult("Match rotation", result));
        Assert.True(undo.TryUndo(out var inverse));

        Assert.Equal(10, inverse.Changes.Single().NewRotation!.Value, Tolerance);
    }

    [Fact]
    public void RefusesASingleShape()
    {
        var result = MatchRotation(new[] { Make.Shape(9, 0, 0, 50, 50, rotation: 45) }, anchorId: 9);

        Assert.False(result.Succeeded);
    }

    [Fact]
    public void TheAlignVerbsNeverCarryRotation()
    {
        // The guarantee that makes rotation optional safe: nothing that only moves shapes can write
        // an angle, so a rotated shape aligned by PowerPoint keeps its rotation.
        var result = AlignSolver.Solve(
            new AlignRequest(AlignVerb.AlignLeft),
            new[] { Make.Shape(1, 0, 0, 50, 50, rotation: 30), Make.Shape(2, 100, 0, 50, 50) },
            Make.Slide());

        Assert.All(result.Changes, c => Assert.Null(c.NewRotation));
    }
}
