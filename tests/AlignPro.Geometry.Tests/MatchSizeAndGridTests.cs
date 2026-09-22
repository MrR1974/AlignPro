namespace AlignPro.Geometry.Tests;

public class MatchSizeTests
{
    private const double Tolerance = 1e-9;

    private static SolveResult Match(
        AlignVerb verb,
        ShapeSnapshot[] shapes,
        int anchorId = 1,
        ResizeOrigin origin = ResizeOrigin.TopLeft,
        bool allowGroupResize = false,
        BoundsModel model = BoundsModel.ShapeFrame) =>
        AlignSolver.Solve(
            new AlignRequest(
                verb,
                ReferenceTarget.Anchor,
                model,
                Make.Key(anchorId),
                resizeOrigin: origin,
                allowGroupResize: allowGroupResize),
            shapes,
            Make.Slide());

    [Fact]
    public void MatchWidth_ChangesWidthOnly()
    {
        var result = Match(AlignVerb.MatchWidth, new[]
        {
            Make.Shape(1, 0, 0, 200, 100),
            Make.Shape(2, 300, 300, 50, 40)
        });

        Assert.Equal(200, result.FrameOf(2).Width, Tolerance);
        Assert.Equal(40, result.FrameOf(2).Height, Tolerance);
    }

    [Fact]
    public void MatchHeight_ChangesHeightOnly()
    {
        var result = Match(AlignVerb.MatchHeight, new[]
        {
            Make.Shape(1, 0, 0, 200, 100),
            Make.Shape(2, 300, 300, 50, 40)
        });

        Assert.Equal(50, result.FrameOf(2).Width, Tolerance);
        Assert.Equal(100, result.FrameOf(2).Height, Tolerance);
    }

    [Fact]
    public void MatchBoth_ChangesBothAndHoldsTheTopLeftByDefault()
    {
        var result = Match(AlignVerb.MatchBoth, new[]
        {
            Make.Shape(1, 0, 0, 200, 100),
            Make.Shape(2, 300, 300, 50, 40)
        });

        Assert.Equal(new RectD(300, 300, 200, 100), result.FrameOf(2));
    }

    [Fact]
    public void MatchBoth_WithCentreOrigin_HoldsTheCentre()
    {
        var result = Match(AlignVerb.MatchBoth, new[]
        {
            Make.Shape(1, 0, 0, 200, 100),
            Make.Shape(2, 300, 300, 50, 40)
        }, origin: ResizeOrigin.Centre);

        var frame = result.FrameOf(2);
        Assert.Equal(325, frame.CentreX, Tolerance);   // unchanged from 300 + 50/2
        Assert.Equal(320, frame.CentreY, Tolerance);   // unchanged from 300 + 40/2
        Assert.Equal(200, frame.Width, Tolerance);
        Assert.Equal(100, frame.Height, Tolerance);
    }

    [Fact]
    public void MatchSize_LeavesTheAnchorAlone()
    {
        var result = Match(AlignVerb.MatchBoth, new[]
        {
            Make.Shape(1, 0, 0, 200, 100),
            Make.Shape(2, 300, 300, 50, 40)
        });

        Assert.False(result.Touched(1));
    }

    [Fact]
    public void MatchSize_UsesFrameDimensionsEvenWhenVisualBoundsAreRequested()
    {
        // Documented decision: matching a rotated shape's visual width is ill-posed, so the match
        // verbs always work in frame space. A 45-degree anchor must still hand over 200x100.
        var result = Match(AlignVerb.MatchBoth, new[]
        {
            Make.Shape(1, 0, 0, 200, 100, rotation: 45),
            Make.Shape(2, 300, 300, 50, 40)
        }, model: BoundsModel.VisualBounds);

        Assert.Equal(200, result.FrameOf(2).Width, Tolerance);
        Assert.Equal(100, result.FrameOf(2).Height, Tolerance);
    }

    [Fact]
    public void MatchSize_SkipsGroupsAndSaysSo()
    {
        // Probe 3 measured that scaling a group scales the gaps between its children.
        var result = Match(AlignVerb.MatchBoth, new[]
        {
            Make.Shape(1, 0, 0, 200, 100),
            Make.Shape(2, 300, 300, 50, 40, isGroup: true),
            Make.Shape(3, 500, 300, 50, 40)
        });

        Assert.True(result.Succeeded);
        Assert.False(result.Touched(2));
        Assert.True(result.Touched(3));
        Assert.Contains(result.Diagnostics, d => d.Contains("group", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void MatchSize_WithAllowGroupResize_IncludesGroups()
    {
        var result = Match(AlignVerb.MatchBoth, new[]
        {
            Make.Shape(1, 0, 0, 200, 100),
            Make.Shape(2, 300, 300, 50, 40, isGroup: true)
        }, allowGroupResize: true);

        Assert.True(result.Succeeded);
        Assert.Equal(200, result.FrameOf(2).Width, Tolerance);
    }

    [Fact]
    public void MatchSize_AllowsAGroupAsTheAnchor()
    {
        // This once refused, which was a mistake: the anchor is only measured, never resized, so
        // "resizing a group rescales its internal spacing" is no objection to anchoring on one.
        // The rule belongs to group *targets*, which are still skipped.
        var result = Match(AlignVerb.MatchBoth, new[]
        {
            Make.Shape(1, 0, 0, 200, 100, isGroup: true),
            Make.Shape(2, 300, 300, 50, 40)
        });

        Assert.True(result.Succeeded);
        Assert.Equal(200, result.FrameOf(2).Width, Tolerance);
    }

    [Fact]
    public void MatchSize_RefusesWhenEverySubjectIsAGroup()
    {
        var result = Match(AlignVerb.MatchBoth, new[]
        {
            Make.Shape(1, 0, 0, 200, 100),
            Make.Shape(2, 300, 300, 50, 40, isGroup: true)
        });

        Assert.False(result.Succeeded);
        Assert.Contains("group", result.Diagnostics.Single(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MatchSize_RefusesWithoutAnAnchor()
    {
        var result = AlignSolver.Solve(
            new AlignRequest(AlignVerb.MatchBoth),
            new[] { Make.Shape(1, 0, 0, 200, 100), Make.Shape(2, 300, 300, 50, 40) },
            Make.Slide());

        Assert.False(result.Succeeded);
    }

    [Fact]
    public void MatchSize_RefusesASingleShape()
    {
        var result = Match(AlignVerb.MatchBoth, new[] { Make.Shape(1, 0, 0, 200, 100) });

        Assert.False(result.Succeeded);
    }
}

/// <summary>
/// The margin is measured per side, so one step takes twice it off each dimension. The anchor is last in
/// the selection in every test here, matching what SelectionReader hands the solver.
/// </summary>
public class MatchSizeMarginTests
{
    private const double Tolerance = 1e-9;

    private static SolveResult Match(
        AlignVerb verb,
        ShapeSnapshot[] shapes,
        double margin,
        SizeMarginMode mode,
        int anchorId,
        ResizeOrigin origin = ResizeOrigin.TopLeft) =>
        AlignSolver.Solve(
            new AlignRequest(
                verb,
                ReferenceTarget.Anchor,
                BoundsModel.ShapeFrame,
                Make.Key(anchorId),
                resizeOrigin: origin,
                sizeMargin: margin,
                sizeMarginMode: mode),
            shapes,
            Make.Slide());

    /// <summary>Three shapes to resize, then the anchor last - the shape of a real selection.</summary>
    private static ShapeSnapshot[] ThreeThenAnchor() => new[]
    {
        Make.Shape(1, 0, 0, 50, 50),
        Make.Shape(2, 100, 0, 50, 50),
        Make.Shape(3, 200, 0, 50, 50),
        Make.Shape(9, 300, 0, 200, 100)
    };

    [Fact]
    public void Uniform_TakesTheMarginOffEverySide()
    {
        var result = Match(AlignVerb.MatchBoth, ThreeThenAnchor(), 10, SizeMarginMode.Uniform, anchorId: 9);

        // 10pt each side, so 20 off each dimension - the same for all three.
        foreach (var id in new[] { 1, 2, 3 })
        {
            Assert.Equal(180, result.FrameOf(id).Width, Tolerance);
            Assert.Equal(80, result.FrameOf(id).Height, Tolerance);
        }
    }

    [Fact]
    public void Uniform_LeavesTheAnchorAlone()
    {
        var result = Match(AlignVerb.MatchBoth, ThreeThenAnchor(), 10, SizeMarginMode.Uniform, anchorId: 9);

        Assert.False(result.Touched(9));
    }

    [Fact]
    public void Cascade_MakesTheFirstSelectedSmallest()
    {
        var result = Match(AlignVerb.MatchBoth, ThreeThenAnchor(), 10, SizeMarginMode.Cascade, anchorId: 9);

        // Anchor 200 wide; three steps back from it, then two, then one.
        Assert.Equal(140, result.FrameOf(1).Width, Tolerance);
        Assert.Equal(160, result.FrameOf(2).Width, Tolerance);
        Assert.Equal(180, result.FrameOf(3).Width, Tolerance);

        Assert.True(result.FrameOf(1).Width < result.FrameOf(2).Width);
        Assert.True(result.FrameOf(2).Width < result.FrameOf(3).Width);
    }

    [Fact]
    public void Cascade_TiersHeightByTheSamePoints()
    {
        var result = Match(AlignVerb.MatchBoth, ThreeThenAnchor(), 10, SizeMarginMode.Cascade, anchorId: 9);

        // Same points off each dimension, so a 200x100 anchor tiers to 140x40, not proportionally.
        Assert.Equal(40, result.FrameOf(1).Height, Tolerance);
        Assert.Equal(60, result.FrameOf(2).Height, Tolerance);
        Assert.Equal(80, result.FrameOf(3).Height, Tolerance);
    }

    [Fact]
    public void Cascade_WithCentreOrigin_NestsConcentrically()
    {
        var result = Match(
            AlignVerb.MatchBoth, ThreeThenAnchor(), 10, SizeMarginMode.Cascade,
            anchorId: 9, origin: ResizeOrigin.Centre);

        // Every shape keeps its own centre, which is what makes the rings concentric once they are
        // stacked on the anchor.
        Assert.Equal(25, result.FrameOf(1).CentreX, Tolerance);
        Assert.Equal(125, result.FrameOf(2).CentreX, Tolerance);
        Assert.Equal(225, result.FrameOf(3).CentreX, Tolerance);
    }

    [Fact]
    public void NegativeMargin_GrowsTheShapesInstead()
    {
        var result = Match(AlignVerb.MatchBoth, ThreeThenAnchor(), -10, SizeMarginMode.Uniform, anchorId: 9);

        Assert.Equal(220, result.FrameOf(1).Width, Tolerance);
        Assert.Equal(120, result.FrameOf(1).Height, Tolerance);
    }

    [Fact]
    public void Cascade_WithGrow_MakesTheFirstSelectedLargest()
    {
        // What the ribbon's Direction = Grow sends: the same positive margin, negated. With the
        // anchor last, the cascade that made the first shape smallest now makes it largest - the
        // screentips promise exactly this, so it is pinned here.
        var result = Match(AlignVerb.MatchBoth, ThreeThenAnchor(), -10, SizeMarginMode.Cascade, anchorId: 9);

        Assert.Equal(260, result.FrameOf(1).Width, Tolerance);
        Assert.Equal(240, result.FrameOf(2).Width, Tolerance);
        Assert.Equal(220, result.FrameOf(3).Width, Tolerance);
        Assert.Equal(160, result.FrameOf(1).Height, Tolerance);
    }

    [Fact]
    public void MatchWidth_LeavesHeightAloneEvenWithAMargin()
    {
        var result = Match(AlignVerb.MatchWidth, ThreeThenAnchor(), 10, SizeMarginMode.Uniform, anchorId: 9);

        Assert.Equal(180, result.FrameOf(1).Width, Tolerance);
        Assert.Equal(50, result.FrameOf(1).Height, Tolerance);
    }

    [Fact]
    public void NoneMode_IgnoresTheMarginEntirely()
    {
        var result = Match(AlignVerb.MatchBoth, ThreeThenAnchor(), 25, SizeMarginMode.None, anchorId: 9);

        Assert.Equal(200, result.FrameOf(1).Width, Tolerance);
        Assert.Equal(100, result.FrameOf(1).Height, Tolerance);
    }

    [Fact]
    public void Cascade_SkipsOnlyTheShapesTheMarginWouldCollapse()
    {
        // 60pt a side against a 100pt-tall anchor: step 1 leaves -20 height, steps 2 and 3 worse.
        var result = Match(AlignVerb.MatchBoth, ThreeThenAnchor(), 60, SizeMarginMode.Cascade, anchorId: 9);

        Assert.False(result.Succeeded);
        Assert.Contains("too large", string.Join(" ", result.Diagnostics), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Cascade_ResizesWhatItCanAndReportsWhatItSkipped()
    {
        // 30pt a side: step 1 leaves 140x40, step 2 leaves 80x-20, step 3 worse still.
        var result = Match(AlignVerb.MatchBoth, ThreeThenAnchor(), 30, SizeMarginMode.Cascade, anchorId: 9);

        Assert.True(result.Succeeded);
        Assert.True(result.Notable);
        Assert.True(result.Touched(3));
        Assert.False(result.Touched(2));
        Assert.False(result.Touched(1));
        Assert.Contains("no size", string.Join(" ", result.Diagnostics), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Cascade_GrowsShapesSelectedAfterTheAnchor()
    {
        // Anchor first, so the others sit after it and the same rule tiers them upwards.
        var result = Match(AlignVerb.MatchBoth, new[]
        {
            Make.Shape(9, 300, 0, 200, 100),
            Make.Shape(1, 0, 0, 50, 50),
            Make.Shape(2, 100, 0, 50, 50)
        }, 10, SizeMarginMode.Cascade, anchorId: 9);

        Assert.Equal(220, result.FrameOf(1).Width, Tolerance);
        Assert.Equal(240, result.FrameOf(2).Width, Tolerance);
    }

    [Fact]
    public void Margin_StillSkipsGroups()
    {
        var result = Match(AlignVerb.MatchBoth, new[]
        {
            Make.Shape(1, 0, 0, 50, 50, isGroup: true),
            Make.Shape(2, 100, 0, 50, 50),
            Make.Shape(9, 300, 0, 200, 100)
        }, 10, SizeMarginMode.Uniform, anchorId: 9);

        Assert.True(result.Notable);
        Assert.False(result.Touched(1));
        Assert.Equal(180, result.FrameOf(2).Width, Tolerance);
    }
}

public class GridTests
{
    private const double Tolerance = 1e-9;

    private static SolveResult Grid(
        ShapeSnapshot[] shapes,
        int? columns = null,
        double gapH = 0,
        double gapV = 0,
        GridFillOrder order = GridFillOrder.RowMajor,
        ReferenceTarget reference = ReferenceTarget.SelectionBounds,
        SlideMetrics? slide = null) =>
        AlignSolver.Solve(
            new AlignRequest(
                AlignVerb.GridArrange,
                reference,
                gridColumns: columns,
                gridGapH: gapH,
                gridGapV: gapV,
                gridFillOrder: order),
            shapes,
            slide ?? Make.Slide());

    private static ShapeSnapshot[] FourSquares() => new[]
    {
        Make.Shape(1, 10, 10, 40, 40),
        Make.Shape(2, 300, 30, 40, 40),
        Make.Shape(3, 20, 300, 40, 40),
        Make.Shape(4, 310, 320, 40, 40)
    };

    [Fact]
    public void Grid_DerivesANearSquareGridWhenColumnsAreNotGiven()
    {
        var result = Grid(FourSquares());

        Assert.True(result.Succeeded);
        Assert.Contains(result.Diagnostics, d => d.Contains("2 x 2", StringComparison.Ordinal));
    }

    [Fact]
    public void Grid_RespectsAnExplicitColumnCount()
    {
        var result = Grid(FourSquares(), columns: 4);

        Assert.Contains(result.Diagnostics, d => d.Contains("1 x 4", StringComparison.Ordinal));
    }

    [Fact]
    public void Grid_PlacesShapesAcrossTheSlideWhenAskedTo()
    {
        var result = Grid(FourSquares(), columns: 2, reference: ReferenceTarget.Slide);

        // Cells are 480 x 270, so cell centres sit at 240/720 across and 135/405 down.
        Assert.Equal(240, result.FrameOf(1).CentreX, Tolerance);
        Assert.Equal(135, result.FrameOf(1).CentreY, Tolerance);
        Assert.Equal(720, result.FrameOf(2).CentreX, Tolerance);
        Assert.Equal(405, result.FrameOf(4).CentreY, Tolerance);
    }

    [Fact]
    public void Grid_NeverChangesSizes()
    {
        var result = Grid(FourSquares(), columns: 2, reference: ReferenceTarget.Slide);

        foreach (var change in result.Changes)
        {
            Assert.Equal(change.OldFrame.Width, change.NewFrame.Width, Tolerance);
            Assert.Equal(change.OldFrame.Height, change.NewFrame.Height, Tolerance);
        }
    }

    [Fact]
    public void Grid_IsSafeForGroupsBecauseItOnlyRepositions()
    {
        var shapes = new[]
        {
            Make.Shape(1, 10, 10, 40, 40, isGroup: true),
            Make.Shape(2, 300, 30, 40, 40),
            Make.Shape(3, 20, 300, 40, 40)
        };

        var result = Grid(shapes, columns: 3, reference: ReferenceTarget.Slide);

        Assert.True(result.Succeeded);
        Assert.True(result.Touched(1));
        Assert.Equal(40, result.FrameOf(1).Width, Tolerance);
    }

    [Fact]
    public void Grid_RespectsGaps()
    {
        var result = Grid(FourSquares(), columns: 2, gapH: 100, gapV: 40, reference: ReferenceTarget.Slide);

        // Cell width = (960 - 100) / 2 = 430, so centres at 215 and 215 + 430 + 100 = 745.
        Assert.Equal(215, result.FrameOf(1).CentreX, Tolerance);
        Assert.Equal(745, result.FrameOf(2).CentreX, Tolerance);
    }

    [Fact]
    public void Grid_ColumnMajorPutsTheLeftmostShapesInTheLeftColumn()
    {
        var result = Grid(FourSquares(), columns: 2, order: GridFillOrder.ColumnMajor,
            reference: ReferenceTarget.Slide);

        // Column-major bands by horizontal position, so the two leftmost shapes - 1 at x=30 and 3 at
        // x=40 - form the left column, and each column is then ordered top to bottom. Tidying should
        // keep shapes near where they already were, not reshuffle them across the slide.
        Assert.Equal(240, result.FrameOf(1).CentreX, Tolerance);
        Assert.Equal(240, result.FrameOf(3).CentreX, Tolerance);
        Assert.Equal(720, result.FrameOf(2).CentreX, Tolerance);
        Assert.Equal(720, result.FrameOf(4).CentreX, Tolerance);

        Assert.Equal(135, result.FrameOf(1).CentreY, Tolerance);
        Assert.Equal(405, result.FrameOf(3).CentreY, Tolerance);
    }

    [Fact]
    public void Grid_TidiesIntoReadingOrderRatherThanReshuffling()
    {
        // Shapes are handed over out of order; the one already top-left must land top-left.
        var shapes = new[]
        {
            Make.Shape(4, 310, 320, 40, 40),
            Make.Shape(1, 10, 10, 40, 40),
            Make.Shape(3, 20, 300, 40, 40),
            Make.Shape(2, 300, 30, 40, 40)
        };

        var result = Grid(shapes, columns: 2, reference: ReferenceTarget.Slide);

        Assert.Equal(240, result.FrameOf(1).CentreX, Tolerance);
        Assert.Equal(135, result.FrameOf(1).CentreY, Tolerance);
        Assert.Equal(720, result.FrameOf(4).CentreX, Tolerance);
        Assert.Equal(405, result.FrameOf(4).CentreY, Tolerance);
    }

    [Fact]
    public void Grid_RefusesWhenGapsLeaveNoRoomForCells()
    {
        var result = Grid(FourSquares(), columns: 2, gapH: 2000, reference: ReferenceTarget.Slide);

        Assert.False(result.Succeeded);
        Assert.Contains("gap", result.Diagnostics.Single(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Grid_RefusesASingleShape()
    {
        var result = Grid(new[] { Make.Shape(1, 0, 0, 40, 40) });

        Assert.False(result.Succeeded);
    }

    [Fact]
    public void Grid_HandlesARaggedLastRow()
    {
        var shapes = new[]
        {
            Make.Shape(1, 10, 10, 40, 40),
            Make.Shape(2, 300, 30, 40, 40),
            Make.Shape(3, 20, 300, 40, 40)
        };

        var result = Grid(shapes, columns: 2, reference: ReferenceTarget.Slide);

        Assert.True(result.Succeeded);
        Assert.Equal(3, result.Changes.Count);
        Assert.Contains(result.Diagnostics, d => d.Contains("2 x 2", StringComparison.Ordinal));
    }
}
