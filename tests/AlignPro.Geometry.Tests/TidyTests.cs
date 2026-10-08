namespace AlignPro.Geometry.Tests;

/// <summary>
/// Tidy: shapes that are nearly lined up are made exact (step 11), rows and columns that are nearly
/// evenly spaced are made even (step 12), and nothing else moves. The cases are the plan's list in
/// docs/development.md.
/// </summary>
public class TidyTests
{
    private const double Tolerance = 1e-9;

    private static SolveResult Tidy(params ShapeSnapshot[] shapes) =>
        TidySolver.Solve(new TidyRequest(), shapes);

    private static SolveResult Tidy(BoundsModel model, params ShapeSnapshot[] shapes) =>
        TidySolver.Solve(new TidyRequest(boundsModel: model), shapes);

    [Fact]
    public void AShapeTwoPointsOff_SnapsToTheOtherTwo_WhichDoNotMove()
    {
        var result = Tidy(
            Make.Shape(1, 100, 0, 50, 30),
            Make.Shape(2, 100, 100, 50, 30),
            Make.Shape(3, 102, 200, 50, 30));

        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        Assert.Equal(100, result.FrameOf(3).Left, Tolerance);
        Assert.False(result.Touched(1));
        Assert.False(result.Touched(2));
    }

    [Fact]
    public void AChainOfSmallSteps_DoesNotCollapseIntoOneCluster()
    {
        // 0, 2, 4, 6 at 3pt: each neighbour is close, but the run spans 6pt. It must split, not merge.
        var result = Tidy(
            Make.Shape(1, 100, 0, 40, 30),
            Make.Shape(2, 102, 100, 55, 30),
            Make.Shape(3, 104, 200, 70, 30),
            Make.Shape(4, 106, 300, 85, 30));

        var lefts = Enumerable.Range(1, 4)
            .Select(id => result.Touched(id) ? result.FrameOf(id).Left : 100 + (id - 1) * 2.0)
            .Distinct()
            .Count();
        Assert.Equal(2, lefts);
    }

    [Fact]
    public void ADeliberateStagger_IsLeftAlone()
    {
        var result = Tidy(
            Make.Shape(1, 100, 0, 50, 30),
            Make.Shape(2, 106, 100, 50, 30));

        Assert.True(result.Succeeded);
        Assert.Empty(result.EffectiveChanges);
        Assert.Contains(result.Diagnostics, d => d.StartsWith("Nothing to tidy", StringComparison.Ordinal));
    }

    [Fact]
    public void AnExactAlignment_IsNeverBrokenForANearOne()
    {
        // 1 and 2 are exactly centred on each other. 1, 3 and 4 are nearly left-aligned, and snapping
        // them to the median (3's left) would move 1 off that centre line - so that cluster is skipped.
        var result = Tidy(
            Make.Shape(1, 150, 0, 100, 30),
            Make.Shape(2, 170, 100, 60, 30),
            Make.Shape(3, 152, 200, 30, 30),
            Make.Shape(4, 152.5, 300, 45, 30));

        Assert.False(result.Touched(1));
        Assert.False(result.Touched(2));
        Assert.Contains(result.Diagnostics, d => d.Contains("skipped", StringComparison.Ordinal));
    }

    [Fact]
    public void CompetingClusters_TheLargerWins_AndTheOtherIsSkippedWhole()
    {
        // 1, 2 and 5 are nearly left-aligned (three shapes). 5 and 4 are nearly centre-aligned (two),
        // and that would move 5 again. The three-shape cluster wins; 4 stays exactly where it was.
        var result = Tidy(
            Make.Shape(1, 100, 0, 50, 30),
            Make.Shape(2, 101, 100, 50, 30),
            Make.Shape(5, 102, 200, 60, 30),
            Make.Shape(4, 105, 300, 50, 30));

        Assert.Equal(101, result.FrameOf(5).Left, Tolerance);
        Assert.False(result.Touched(4));
        Assert.Contains(result.Diagnostics, d => d.Contains("skipped", StringComparison.Ordinal));
    }

    [Fact]
    public void APlaceholder_HoldsStill_AndOthersSnapToIt()
    {
        var result = Tidy(
            Make.Shape(1, 102, 0, 50, 30),
            Make.Shape(9, 100, 100, 300, 200, isPlaceholder: true));

        Assert.Equal(100, result.FrameOf(1).Left, Tolerance);
        Assert.False(result.Touched(9));
    }

    [Fact]
    public void TwoPlaceholdersThatDisagree_SkipTheCluster()
    {
        var result = Tidy(
            Make.Shape(1, 102, 0, 50, 30),
            Make.Shape(8, 100, 100, 300, 50, isPlaceholder: true),
            Make.Shape(9, 101, 200, 200, 50, isPlaceholder: true));

        Assert.Empty(result.EffectiveChanges);
    }

    [Fact]
    public void Connectors_AreNeitherMeasuredNorMoved()
    {
        // Measured, the connector at 101 would be the median everyone snaps to. Left out, 1 and 2
        // settle on 1's edge, and the connector stays put.
        var result = Tidy(
            Make.Shape(1, 100, 0, 50, 30),
            Make.Shape(2, 102, 100, 50, 30),
            Make.Shape(3, 101, 200, 50, 30, isConnector: true));

        Assert.Equal(100, result.FrameOf(2).Left, Tolerance);
        Assert.False(result.Touched(3));
        Assert.Contains(result.Diagnostics, d => d.Contains("connector", StringComparison.Ordinal));
    }

    [Fact]
    public void AGroup_IsMovedWhole_NeverResized()
    {
        var result = Tidy(
            Make.Shape(1, 100, 0, 50, 30),
            Make.Shape(2, 102, 100, 200, 120, isGroup: true));

        var group = result.FrameOf(2);
        Assert.Equal(100, group.Left, Tolerance);
        Assert.Equal(200, group.Width, Tolerance);
        Assert.Equal(120, group.Height, Tolerance);
    }

    [Fact]
    public void NoChange_CarriesARotation()
    {
        var result = Tidy(Grid(jitter: true));

        Assert.NotEmpty(result.EffectiveChanges);
        Assert.All(result.Changes, c => Assert.Null(c.NewRotation));
    }

    [Fact]
    public void ARotatedShape_IsJudgedByItsVisualBounds_WhenAsked()
    {
        // Turned 90 degrees, a 100x40 frame at x=100 shows as 40 wide from x=130. Shape 2 is wide
        // enough that nothing of it is near shape 1's frame: only the visual left edges are close.
        var shapes = new[]
        {
            Make.Shape(1, 100, 100, 100, 40, rotation: 90),
            Make.Shape(2, 131, 300, 60, 40)
        };

        Assert.Empty(Tidy(BoundsModel.ShapeFrame, shapes).EffectiveChanges);

        var visual = Tidy(BoundsModel.VisualBounds, shapes);
        Assert.Equal(130, visual.FrameOf(2).Left, Tolerance);
    }

    [Fact]
    public void AJitteredGrid_ComesBackExactlyAlignedAndEvenlySpaced()
    {
        var shapes = Grid(jitter: true);
        var result = Tidy(shapes);

        var after = Apply(shapes, result);
        var columnLefts = new double[3];
        var rowTops = new double[3];
        for (var column = 0; column < 3; column++)
        {
            columnLefts[column] = Assert.Single(after.Where((_, i) => i % 3 == column).Select(s => Math.Round(s.Frame.Left, 6)).Distinct());
        }
        for (var row = 0; row < 3; row++)
        {
            rowTops[row] = Assert.Single(after.Where((_, i) => i / 3 == row).Select(s => Math.Round(s.Frame.Top, 6)).Distinct());
        }

        Assert.Equal(columnLefts[1] - columnLefts[0], columnLefts[2] - columnLefts[1], 1e-6);
        Assert.Equal(rowTops[1] - rowTops[0], rowTops[2] - rowTops[1], 1e-6);
    }

    [Fact]
    public void ASecondRun_FindsNothingToDo()
    {
        var shapes = Grid(jitter: true);
        var once = Apply(shapes, Tidy(shapes));

        var twice = Tidy(once);

        Assert.Empty(twice.EffectiveChanges);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(0.4)]
    [InlineData(20.5)]
    [InlineData(-3)]
    public void ATolerance_OutsideHalfAPointToTwenty_IsRefused(double tolerance)
    {
        var result = TidySolver.Solve(
            new TidyRequest(tolerance), new[] { Make.Shape(1, 100, 0, 50, 30), Make.Shape(2, 102, 100, 50, 30) });

        Assert.False(result.Succeeded);
    }

    [Theory]
    [InlineData(0.5)]
    [InlineData(20)]
    public void TheToleranceLimits_AreThemselvesAccepted(double tolerance)
    {
        var result = TidySolver.Solve(
            new TidyRequest(tolerance), new[] { Make.Shape(1, 100, 0, 50, 30), Make.Shape(2, 102, 100, 50, 30) });

        Assert.True(result.Succeeded);
    }

    [Fact]
    public void ItAlwaysSaysWhatItDid()
    {
        var result = Tidy(Make.Shape(1, 100, 0, 50, 30), Make.Shape(2, 102, 100, 50, 30));

        Assert.True(result.Notable);
        Assert.StartsWith("Tidied 1 shape", result.Diagnostics[0], StringComparison.Ordinal);
    }

    /// <summary>A 3x3 grid of 40pt squares at a 100pt pitch, each nudged by up to a point when asked.</summary>
    private static ShapeSnapshot[] Grid(bool jitter)
    {
        // Fixed nudges rather than random ones, so a failure can be reproduced.
        double[] nudge = { 0.4, -0.9, 0.7, -0.3, 1.0, -0.6, 0.2, -1.0, 0.8, -0.5, 0.9, -0.2, 0.6, -0.8, 0.1, -0.4, 0.5, -0.7 };
        var shapes = new ShapeSnapshot[9];
        for (var i = 0; i < 9; i++)
        {
            var x = 100 + (i % 3) * 100 + (jitter ? nudge[i * 2] : 0);
            var y = 100 + (i / 3) * 100 + (jitter ? nudge[i * 2 + 1] : 0);
            shapes[i] = Make.Shape(i + 1, x, y, 40, 40);
        }

        return shapes;
    }

    /// <summary>The shapes as they would stand after the result was applied.</summary>
    private static ShapeSnapshot[] Apply(ShapeSnapshot[] shapes, SolveResult result) =>
        shapes.Select(s =>
        {
            var change = result.Changes.FirstOrDefault(c => c.Key == s.Key);
            return change == null ? s : Make.Shape(s.Key.ShapeId, change.NewFrame.Left, change.NewFrame.Top, s.Frame.Width, s.Frame.Height);
        }).ToArray();

    // -- spacing (step 12) --------------------------------------------------------------------------

    private static double Gap(SolveResult result, ShapeSnapshot[] shapes, int leftId, int rightId)
    {
        var after = Apply(shapes, result);
        return after.Single(s => s.Key.ShapeId == rightId).Frame.Left - after.Single(s => s.Key.ShapeId == leftId).Frame.Right;
    }

    [Fact]
    public void ARowWithNearlyEvenGaps_IsMadeEven_HoldingItsEnds()
    {
        // Gaps of 18, 20 and 21: within 3pt of each other, so meant as even.
        var shapes = new[]
        {
            Make.Shape(1, 100, 100, 40, 40),
            Make.Shape(2, 158, 100, 40, 40),
            Make.Shape(3, 218, 100, 40, 40),
            Make.Shape(4, 279, 100, 40, 40)
        };

        var result = Tidy(shapes);

        Assert.False(result.Touched(1));
        Assert.False(result.Touched(4));
        Assert.Equal(Gap(result, shapes, 1, 2), Gap(result, shapes, 2, 3), 1e-6);
        Assert.Equal(Gap(result, shapes, 2, 3), Gap(result, shapes, 3, 4), 1e-6);
        Assert.Contains("1 row", result.Diagnostics[0], StringComparison.Ordinal);
    }

    [Fact]
    public void ARowTooUnevenToBeMeant_IsLeftAlone()
    {
        var result = Tidy(
            Make.Shape(1, 100, 100, 40, 40),
            Make.Shape(2, 150, 100, 40, 40),
            Make.Shape(3, 210, 100, 40, 40));

        Assert.Empty(result.EffectiveChanges);
    }

    [Fact]
    public void OverlappingShapes_AreNotARow()
    {
        var result = Tidy(
            Make.Shape(1, 100, 100, 60, 40),
            Make.Shape(2, 150, 100, 60, 40),
            Make.Shape(3, 208, 100, 60, 40));

        Assert.Empty(result.EffectiveChanges);
    }

    [Fact]
    public void GapsAreEvened_WhenTheyVaryLessThanThePitch()
    {
        // Widths 20, 60, 30. Gaps 20 and 21 vary by 1; pitches 60 and 66 vary by 6.
        var shapes = new[]
        {
            Make.Shape(1, 100, 100, 20, 40),
            Make.Shape(2, 140, 100, 60, 40),
            Make.Shape(3, 221, 100, 30, 40)
        };

        var result = Tidy(shapes);

        Assert.Equal(Gap(result, shapes, 1, 2), Gap(result, shapes, 2, 3), 1e-6);
    }

    [Fact]
    public void CentresAreEvened_WhenTheyVaryLessThanTheGaps()
    {
        // Widths 20, 60, 30. Pitches 100 and 101 vary by 1; gaps 60 and 56 vary by 4.
        var shapes = new[]
        {
            Make.Shape(1, 100, 100, 20, 40),
            Make.Shape(2, 180, 100, 60, 40),
            Make.Shape(3, 296, 100, 30, 40)
        };

        var result = Tidy(shapes);

        Assert.Equal(210.5, result.FrameOf(2).CentreX, 1e-6);
    }

    [Fact]
    public void ARowThatWouldMoveAPlaceholder_IsSkipped()
    {
        var result = Tidy(
            Make.Shape(1, 100, 100, 40, 40),
            Make.Shape(9, 158, 100, 40, 40, isPlaceholder: true),
            Make.Shape(3, 218, 100, 40, 40));

        Assert.Empty(result.EffectiveChanges);
        Assert.Contains(result.Diagnostics, d => d.Contains("skipped", StringComparison.Ordinal));
    }

    [Fact]
    public void EveningARow_SlidesTheWholeColumnWithIt()
    {
        // 2 is exactly above 4. Evening the row moves 2, and 4 must come too, or the column breaks.
        var shapes = new[]
        {
            Make.Shape(1, 100, 100, 40, 40),
            Make.Shape(2, 158, 100, 40, 40),
            Make.Shape(3, 218, 100, 40, 40),
            Make.Shape(4, 158, 300, 40, 40)
        };

        var result = Tidy(shapes);

        Assert.True(result.Touched(2));
        Assert.Equal(result.FrameOf(2).Left, result.FrameOf(4).Left, 1e-9);
    }

    [Fact]
    public void RoundingNoise_CountsAsExact_SoATidyLayoutStaysTidy()
    {
        // An evenly spaced, aligned row as PowerPoint might read it back: off by thousandths.
        var result = Tidy(
            Make.Shape(1, 100, 100, 40, 40),
            Make.Shape(2, 160.004, 100.003, 40, 40),
            Make.Shape(3, 219.998, 99.996, 40, 40));

        Assert.Empty(result.EffectiveChanges);
        Assert.StartsWith("Nothing to tidy", result.Diagnostics[0], StringComparison.Ordinal);
    }
}
