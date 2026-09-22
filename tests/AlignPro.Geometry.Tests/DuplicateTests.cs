namespace AlignPro.Geometry.Tests;

public class DuplicateTests
{
    private const double Tolerance = 1e-6;

    private static DuplicateResult Duplicate(
        IReadOnlyList<ShapeSnapshot> shapes,
        double dx = 0,
        double dy = 0,
        double angle = 0,
        int copies = 1,
        DuplicatePivot pivot = DuplicatePivot.OwnCentre,
        bool rotateShapes = true,
        int? anchorId = null) =>
        DuplicateSolver.Solve(
            new DuplicateRequest(dx, dy, angle, copies, pivot, rotateShapes,
                anchorId.HasValue ? Make.Key(anchorId.Value) : null),
            shapes,
            Make.Slide());

    private static ShapePlacement Placement(DuplicateResult result, int copy, int sourceId) =>
        result.Copies.Single(c => c.Index == copy).Placements.Single(p => p.Source == Make.Key(sourceId));

    // -- translation ----------------------------------------------------------------------------

    [Fact]
    public void OneCopy_IsOffsetByOneStep()
    {
        var result = Duplicate(new[] { Make.Shape(1, 100, 100, 50, 40) }, dx: 60, dy: 10);

        Assert.True(result.Succeeded);
        Assert.Equal(new RectD(160, 110, 50, 40), Placement(result, 1, 1).Frame);
    }

    [Fact]
    public void CopyK_ReceivesTheStepKTimes()
    {
        var result = Duplicate(new[] { Make.Shape(1, 0, 0, 50, 40) }, dx: 60, copies: 4);

        Assert.Equal(4, result.Copies.Count);
        for (var k = 1; k <= 4; k++)
        {
            Assert.Equal(60 * k, Placement(result, k, 1).Frame.X, Tolerance);
        }
    }

    [Fact]
    public void AMultiShapeSelection_KeepsItsInternalLayout()
    {
        var result = Duplicate(new[]
        {
            Make.Shape(1, 0, 0, 50, 40),
            Make.Shape(2, 80, 30, 20, 20)
        }, dy: 100, copies: 2);

        var first = Placement(result, 2, 1).Frame;
        var second = Placement(result, 2, 2).Frame;
        Assert.Equal(80, second.X - first.X, Tolerance);
        Assert.Equal(30, second.Y - first.Y, Tolerance);
    }

    [Fact]
    public void PlacementsFollowSelectionOrder()
    {
        var result = Duplicate(new[]
        {
            Make.Shape(7, 0, 0, 50, 40),
            Make.Shape(3, 80, 30, 20, 20)
        }, dx: 100);

        Assert.Equal(new[] { 7, 3 }, result.Copies[0].Placements.Select(p => p.Source.ShapeId));
    }

    [Fact]
    public void CopiesKeepTheOriginalSize()
    {
        var result = Duplicate(new[] { Make.Shape(1, 0, 0, 57, 31) }, dx: 10, angle: 45, pivot: DuplicatePivot.SlideCentre);

        Assert.Equal(57, Placement(result, 1, 1).Frame.Width, Tolerance);
        Assert.Equal(31, Placement(result, 1, 1).Frame.Height, Tolerance);
    }

    // -- rotation -------------------------------------------------------------------------------

    [Fact]
    public void OwnCentre_SpinsInPlace()
    {
        var result = Duplicate(new[] { Make.Shape(1, 100, 100, 50, 40, rotation: 10) }, angle: 30, copies: 2);

        Assert.Equal(new RectD(100, 100, 50, 40), Placement(result, 2, 1).Frame);
        Assert.Equal(40, Placement(result, 1, 1).Rotation, Tolerance);
        Assert.Equal(70, Placement(result, 2, 1).Rotation, Tolerance);
    }

    [Fact]
    public void SlideCentre_RingsTheSlideClockwise()
    {
        // A 20x20 shape centred 100pt above the slide centre (480, 270).
        var result = Duplicate(
            new[] { Make.Shape(1, 470, 160, 20, 20) }, angle: 90, copies: 3, pivot: DuplicatePivot.SlideCentre);

        // Clockwise on screen: a quarter turn takes twelve o'clock to three o'clock.
        Assert.Equal(580, Placement(result, 1, 1).Frame.CentreX, Tolerance);
        Assert.Equal(270, Placement(result, 1, 1).Frame.CentreY, Tolerance);

        Assert.Equal(480, Placement(result, 2, 1).Frame.CentreX, Tolerance);
        Assert.Equal(370, Placement(result, 2, 1).Frame.CentreY, Tolerance);

        Assert.Equal(380, Placement(result, 3, 1).Frame.CentreX, Tolerance);
        Assert.Equal(270, Placement(result, 3, 1).Frame.CentreY, Tolerance);
    }

    [Fact]
    public void RotateShapesOn_TurnsCopiesWithTheStep()
    {
        var result = Duplicate(
            new[] { Make.Shape(1, 470, 160, 20, 20) }, angle: 90, copies: 3, pivot: DuplicatePivot.SlideCentre);

        Assert.Equal(90, Placement(result, 1, 1).Rotation, Tolerance);
        Assert.Equal(180, Placement(result, 2, 1).Rotation, Tolerance);
        Assert.Equal(270, Placement(result, 3, 1).Rotation, Tolerance);
    }

    [Fact]
    public void RotateShapesOff_KeepsCopiesUprightWhileTheyStillGoRound()
    {
        var result = Duplicate(
            new[] { Make.Shape(1, 470, 160, 20, 20, rotation: 15) },
            angle: 90, copies: 2, pivot: DuplicatePivot.SlideCentre, rotateShapes: false);

        Assert.Equal(15, Placement(result, 2, 1).Rotation, Tolerance);
        Assert.Equal(370, Placement(result, 2, 1).Frame.CentreY, Tolerance);
    }

    [Fact]
    public void RotationsWrapIntoPowerPointsRange()
    {
        var result = Duplicate(new[] { Make.Shape(1, 0, 0, 20, 20, rotation: 350) }, angle: 30);

        Assert.Equal(20, Placement(result, 1, 1).Rotation, Tolerance);
    }

    [Fact]
    public void RotateThenTranslate_InThatOrder()
    {
        // Centre (100, 0) relative to a pivot at the anchor's centre (0, 0): a quarter turn takes it
        // to (0, 100), then the step moves it right by 10. Translating first would give (0, 110).
        var result = Duplicate(new[]
        {
            Make.Shape(1, 90, -10, 20, 20),
            Make.Shape(9, -10, -10, 20, 20)
        }, dx: 10, angle: 90, pivot: DuplicatePivot.AnchorCentre, anchorId: 9);

        Assert.Equal(10, Placement(result, 1, 1).Frame.CentreX, Tolerance);
        Assert.Equal(100, Placement(result, 1, 1).Frame.CentreY, Tolerance);
    }

    [Fact]
    public void SelectionCentre_TurnsTheSelectionAsOneRigidPiece()
    {
        var result = Duplicate(new[]
        {
            Make.Shape(1, 0, 0, 20, 20),
            Make.Shape(2, 180, 0, 20, 20)
        }, angle: 180, pivot: DuplicatePivot.SelectionCentre);

        // Half a turn about the middle swaps the two ends.
        Assert.Equal(190, Placement(result, 1, 1).Frame.CentreX, Tolerance);
        Assert.Equal(10, Placement(result, 1, 2).Frame.CentreX, Tolerance);
    }

    [Fact]
    public void AnchorPivot_NeedsAnAnchor()
    {
        var result = Duplicate(new[] { Make.Shape(1, 0, 0, 20, 20) }, angle: 45, pivot: DuplicatePivot.AnchorCentre);

        Assert.False(result.Succeeded);
    }

    // -- refusals and notices -------------------------------------------------------------------

    [Fact]
    public void AllZero_IsRefused()
    {
        var result = Duplicate(new[] { Make.Shape(1, 0, 0, 20, 20) });

        Assert.False(result.Succeeded);
        Assert.Contains("zero", result.Diagnostics.Single());
    }

    [Fact]
    public void AnAngleThatMovesNothing_IsRefusedWithTheReason()
    {
        var result = Duplicate(new[] { Make.Shape(1, 0, 0, 20, 20) }, angle: 45, rotateShapes: false);

        Assert.False(result.Succeeded);
        Assert.Contains("Rotate shapes", result.Diagnostics.Single());
    }

    [Fact]
    public void AFullTurn_IsRefusedToo()
    {
        var result = Duplicate(new[] { Make.Shape(1, 0, 0, 20, 20) }, angle: 360, pivot: DuplicatePivot.SlideCentre);

        Assert.False(result.Succeeded);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(DuplicateRequest.MaxCopies + 1)]
    public void CopyCount_IsBounded(int copies)
    {
        var result = Duplicate(new[] { Make.Shape(1, 0, 0, 20, 20) }, dx: 10, copies: copies);

        Assert.False(result.Succeeded);
    }

    [Fact]
    public void CopiesOffTheSlide_AreReported()
    {
        var result = Duplicate(new[] { Make.Shape(1, 800, 100, 50, 50) }, dx: 100, copies: 3);

        Assert.True(result.Succeeded);
        Assert.True(result.Notable);
        Assert.Equal(3, result.Copies.Count);
    }

    // -- the journal ----------------------------------------------------------------------------

    private static ShapeCreation Creation(params int[] createdIds)
    {
        var result = Duplicate(new[] { Make.Shape(1, 0, 0, 20, 20) }, dx: 30, copies: createdIds.Length);
        return new ShapeCreation(result.Copies, createdIds.Select(Make.Key).ToList());
    }

    [Fact]
    public void ACreationTransaction_IsWorthRecording()
    {
        var transaction = AlignTransaction.FromCreation("Duplicate", Creation(50, 51));

        Assert.False(transaction.IsEmpty);
        Assert.Equal(CreationDirection.Create, transaction.Direction);
    }

    [Fact]
    public void Undo_RemovesTheCreatedShapes()
    {
        var undo = new UndoManager();
        undo.Push(AlignTransaction.FromCreation("Duplicate", Creation(50, 51)));

        Assert.True(undo.TryUndo(out var inverse));
        Assert.Equal(CreationDirection.Remove, inverse.Direction);
        Assert.Equal(new[] { 50, 51 }, inverse.Creation!.CreatedKeys.Select(k => k.ShapeId));
    }

    [Fact]
    public void Redo_CreatesAgain_AndItsNewKeysAreWhatTheNextUndoDeletes()
    {
        var undo = new UndoManager();
        undo.Push(AlignTransaction.FromCreation("Duplicate", Creation(50, 51)));

        undo.TryUndo(out _);
        Assert.True(undo.TryRedo(out var redo));
        Assert.Equal(CreationDirection.Create, redo.Direction);

        // PowerPoint hands recreated shapes fresh ids; the add-in records them after the redo.
        redo.Creation!.Rekey(new[] { Make.Key(60), Make.Key(61) });

        Assert.True(undo.TryUndo(out var secondUndo));
        Assert.Equal(new[] { 60, 61 }, secondUndo.Creation!.CreatedKeys.Select(k => k.ShapeId));
    }

    [Fact]
    public void TheRecipeSurvivesTheRoundTrip()
    {
        var undo = new UndoManager();
        var creation = Creation(50, 51);
        undo.Push(AlignTransaction.FromCreation("Duplicate", creation));

        undo.TryUndo(out _);
        undo.TryRedo(out var redo);

        Assert.Same(creation.Copies, redo.Creation!.Copies);
    }

    [Fact]
    public void AnEditOnlyTransaction_HasNoCreation()
    {
        var change = new GeometryChange(Make.Key(1), new RectD(0, 0, 1, 1), new RectD(5, 0, 1, 1));
        var transaction = new AlignTransaction("Align left", new[] { change });

        Assert.Null(transaction.Creation);
        Assert.Null(transaction.Inverted().Creation);
    }
}
