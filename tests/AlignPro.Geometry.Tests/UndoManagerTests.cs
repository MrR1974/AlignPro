namespace AlignPro.Geometry.Tests;

public class UndoManagerTests
{
    private static AlignTransaction Move(string label, int id, double fromX, double toX) =>
        new(label, new[]
        {
            new GeometryChange(Make.Key(id), new RectD(fromX, 0, 50, 50), new RectD(toX, 0, 50, 50))
        });

    [Fact]
    public void NewManager_HasNothingToUndoOrRedo()
    {
        var manager = new UndoManager();

        Assert.False(manager.CanUndo);
        Assert.False(manager.CanRedo);
        Assert.Null(manager.NextUndoLabel);
        Assert.Null(manager.NextRedoLabel);
    }

    [Fact]
    public void Push_ThenUndo_HandsBackTheInverse()
    {
        var manager = new UndoManager();
        manager.Push(Move("Align left", 1, 100, 300));

        Assert.True(manager.TryUndo(out var inverse));

        var change = inverse.Changes.Single();
        Assert.Equal(new RectD(300, 0, 50, 50), change.OldFrame);
        Assert.Equal(new RectD(100, 0, 50, 50), change.NewFrame);
    }

    [Fact]
    public void Undo_MakesTheOperationRedoable()
    {
        var manager = new UndoManager();
        manager.Push(Move("Align left", 1, 100, 300));
        manager.TryUndo(out _);

        Assert.False(manager.CanUndo);
        Assert.True(manager.CanRedo);
        Assert.Equal("Align left", manager.NextRedoLabel);
    }

    [Fact]
    public void Redo_HandsBackTheOriginalDirection()
    {
        var manager = new UndoManager();
        manager.Push(Move("Align left", 1, 100, 300));
        manager.TryUndo(out _);

        Assert.True(manager.TryRedo(out var redo));

        var change = redo.Changes.Single();
        Assert.Equal(new RectD(100, 0, 50, 50), change.OldFrame);
        Assert.Equal(new RectD(300, 0, 50, 50), change.NewFrame);
        Assert.True(manager.CanUndo);
        Assert.False(manager.CanRedo);
    }

    [Fact]
    public void TryUndo_OnEmptyStack_ReturnsFalse()
    {
        var manager = new UndoManager();

        Assert.False(manager.TryUndo(out _));
    }

    [Fact]
    public void TryRedo_OnEmptyStack_ReturnsFalse()
    {
        var manager = new UndoManager();

        Assert.False(manager.TryRedo(out _));
    }

    [Fact]
    public void Push_DiscardsTheRedoStack()
    {
        var manager = new UndoManager();
        manager.Push(Move("Align left", 1, 100, 300));
        manager.TryUndo(out _);
        Assert.True(manager.CanRedo);

        manager.Push(Move("Distribute horizontally", 2, 0, 50));

        Assert.False(manager.CanRedo);
    }

    [Fact]
    public void Push_IgnoresATransactionThatWouldNotMoveAnything()
    {
        var manager = new UndoManager();
        var noOp = new AlignTransaction("Align left", new[]
        {
            new GeometryChange(Make.Key(1), new RectD(100, 0, 50, 50), new RectD(100, 0, 50, 50))
        });

        Assert.False(manager.Push(noOp));
        Assert.False(manager.CanUndo);
    }

    [Fact]
    public void Push_IgnoresAnEmptyTransaction()
    {
        var manager = new UndoManager();

        Assert.False(manager.Push(new AlignTransaction("Align left", Array.Empty<GeometryChange>())));
        Assert.False(manager.CanUndo);
    }

    [Fact]
    public void UndoStack_UnwindsMostRecentFirst()
    {
        var manager = new UndoManager();
        manager.Push(Move("First", 1, 0, 100));
        manager.Push(Move("Second", 2, 0, 200));
        manager.Push(Move("Third", 3, 0, 300));

        Assert.Equal("Third", manager.NextUndoLabel);

        manager.TryUndo(out var third);
        Assert.Equal(Make.Key(3), third.Changes.Single().Key);

        manager.TryUndo(out var second);
        Assert.Equal(Make.Key(2), second.Changes.Single().Key);

        manager.TryUndo(out var first);
        Assert.Equal(Make.Key(1), first.Changes.Single().Key);

        Assert.False(manager.CanUndo);
    }

    [Fact]
    public void Capacity_DropsTheOldestOperation()
    {
        var manager = new UndoManager(capacity: 3);
        manager.Push(Move("First", 1, 0, 100));
        manager.Push(Move("Second", 2, 0, 200));
        manager.Push(Move("Third", 3, 0, 300));
        manager.Push(Move("Fourth", 4, 0, 400));

        Assert.Equal(3, manager.UndoDepth);

        // "First" should have fallen off the end, leaving Fourth, Third, Second.
        manager.TryUndo(out _);
        manager.TryUndo(out _);
        manager.TryUndo(out var oldest);

        Assert.Equal(Make.Key(2), oldest.Changes.Single().Key);
        Assert.False(manager.CanUndo);
    }

    [Fact]
    public void DefaultCapacity_IsTwenty()
    {
        Assert.Equal(20, new UndoManager().Capacity);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Capacity_MustBePositive(int capacity)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new UndoManager(capacity));
    }

    [Fact]
    public void Clear_ForgetsBothStacks()
    {
        var manager = new UndoManager();
        manager.Push(Move("First", 1, 0, 100));
        manager.Push(Move("Second", 2, 0, 200));
        manager.TryUndo(out _);

        manager.Clear();

        Assert.False(manager.CanUndo);
        Assert.False(manager.CanRedo);
        Assert.Equal(0, manager.UndoDepth);
        Assert.Equal(0, manager.RedoDepth);
    }

    [Fact]
    public void UndoThenRedo_RoundTripsRepeatedly()
    {
        var manager = new UndoManager();
        manager.Push(Move("Align left", 1, 100, 300));

        for (var i = 0; i < 5; i++)
        {
            Assert.True(manager.TryUndo(out var inverse));
            Assert.Equal(300, inverse.Changes.Single().OldFrame.Left);

            Assert.True(manager.TryRedo(out var redo));
            Assert.Equal(100, redo.Changes.Single().OldFrame.Left);
        }
    }

    [Fact]
    public void FromResult_DropsChangesThatWouldNotMoveAnything()
    {
        // Aligning three shapes when one is already on the line should record only the two that move.
        var result = AlignSolver.Solve(
            new AlignRequest(AlignVerb.AlignLeft),
            new[]
            {
                Make.Shape(1, 100, 10, 50, 20),
                Make.Shape(2, 300, 50, 50, 20),
                Make.Shape(3, 500, 90, 50, 20)
            },
            Make.Slide());

        var transaction = AlignTransaction.FromResult("Align left", result);

        Assert.Equal(2, transaction.Changes.Count);
        Assert.DoesNotContain(transaction.Changes, c => c.Key == Make.Key(1));
    }

    [Fact]
    public void FromResult_ThenUndo_RestoresTheOriginalFrames()
    {
        var shapes = new[]
        {
            Make.Shape(1, 100, 10, 50, 20),
            Make.Shape(2, 300, 50, 80, 20),
            Make.Shape(3, 500, 90, 60, 20)
        };
        var result = AlignSolver.Solve(new AlignRequest(AlignVerb.AlignLeft), shapes, Make.Slide());

        var manager = new UndoManager();
        Assert.True(manager.Push(AlignTransaction.FromResult("Align left", result)));
        Assert.True(manager.TryUndo(out var inverse));

        // Every inverse change must land each shape back on the frame it started with.
        foreach (var change in inverse.Changes)
        {
            var original = shapes.Single(s => s.Key == change.Key);
            Assert.Equal(original.Frame, change.NewFrame);
        }
    }

    [Fact]
    public void NextUndoLabel_TracksTheTopOfTheStack()
    {
        var manager = new UndoManager();
        manager.Push(Move("Align left", 1, 0, 100));
        Assert.Equal("Align left", manager.NextUndoLabel);

        manager.Push(Move("Match width", 2, 0, 200));
        Assert.Equal("Match width", manager.NextUndoLabel);

        manager.TryUndo(out _);
        Assert.Equal("Align left", manager.NextUndoLabel);
    }

    [Fact]
    public void Transaction_RequiresALabel()
    {
        Assert.Throws<ArgumentException>(() =>
            new AlignTransaction("  ", Array.Empty<GeometryChange>()));
    }
}
