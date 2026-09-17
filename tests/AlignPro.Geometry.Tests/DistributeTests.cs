namespace AlignPro.Geometry.Tests;

public class DistributeTests
{
    private const double Tolerance = 1e-9;

    private static SolveResult Distribute(
        AlignVerb verb,
        ShapeSnapshot[] shapes,
        double? exactSpacing = null,
        DistributeMode mode = DistributeMode.Gap,
        ReferenceTarget reference = ReferenceTarget.SelectionBounds,
        BoundsModel model = BoundsModel.ShapeFrame,
        SlideMetrics? slide = null) =>
        AlignSolver.Solve(
            new AlignRequest(verb, reference, model, anchor: null, exactSpacing: exactSpacing, distributeMode: mode),
            shapes,
            slide ?? Make.Slide());

    [Fact]
    public void DistributeH_EqualisesTheGaps()
    {
        // Three 50-wide shapes spanning 0..350: 350 - 150 = 200 of free space, so 100 per gap.
        var result = Distribute(AlignVerb.DistributeH, new[]
        {
            Make.Shape(1, 0, 10, 50, 20),
            Make.Shape(2, 120, 10, 50, 20),
            Make.Shape(3, 300, 10, 50, 20)
        });

        Assert.True(result.Succeeded);
        Assert.Equal(0, result.FrameOf(1).Left, Tolerance);
        Assert.Equal(150, result.FrameOf(2).Left, Tolerance);
        Assert.Equal(300, result.FrameOf(3).Left, Tolerance);
    }

    [Fact]
    public void DistributeH_HoldsTheOutermostShapesStill()
    {
        var shapes = new[]
        {
            Make.Shape(1, 17, 10, 33, 20),
            Make.Shape(2, 120, 10, 41, 20),
            Make.Shape(3, 200, 10, 55, 20),
            Make.Shape(4, 400, 10, 29, 20)
        };

        var result = Distribute(AlignVerb.DistributeH, shapes);

        Assert.Equal(17, result.FrameOf(1).Left, Tolerance);
        Assert.Equal(429, result.FrameOf(4).Right, Tolerance);
    }

    [Fact]
    public void DistributeH_EqualGaps_AreActuallyEqual()
    {
        var shapes = new[]
        {
            Make.Shape(1, 0, 10, 33, 20),
            Make.Shape(2, 120, 10, 41, 20),
            Make.Shape(3, 200, 10, 55, 20),
            Make.Shape(4, 400, 10, 29, 20)
        };

        var result = Distribute(AlignVerb.DistributeH, shapes);
        var frames = new[] { result.FrameOf(1), result.FrameOf(2), result.FrameOf(3), result.FrameOf(4) }
            .OrderBy(f => f.Left)
            .ToList();

        var firstGap = frames[1].Left - frames[0].Right;
        for (var i = 2; i < frames.Count; i++)
        {
            Assert.Equal(firstGap, frames[i].Left - frames[i - 1].Right, Tolerance);
        }
    }

    [Fact]
    public void DistributeH_WithExactSpacing_AnchorsTheFirstShape()
    {
        var result = Distribute(AlignVerb.DistributeH, new[]
        {
            Make.Shape(1, 100, 10, 50, 20),
            Make.Shape(2, 400, 10, 50, 20),
            Make.Shape(3, 800, 10, 50, 20)
        }, exactSpacing: 12);

        Assert.Equal(100, result.FrameOf(1).Left, Tolerance);
        Assert.Equal(162, result.FrameOf(2).Left, Tolerance);   // 100 + 50 + 12
        Assert.Equal(224, result.FrameOf(3).Left, Tolerance);   // 162 + 50 + 12
    }

    [Fact]
    public void DistributeH_WithExactSpacing_AcceptsTwoShapes()
    {
        var result = Distribute(AlignVerb.DistributeH, new[]
        {
            Make.Shape(1, 100, 10, 50, 20),
            Make.Shape(2, 400, 10, 50, 20)
        }, exactSpacing: 20);

        Assert.True(result.Succeeded);
        Assert.Equal(170, result.FrameOf(2).Left, Tolerance);
    }

    [Fact]
    public void DistributeH_WithoutExactSpacing_RefusesTwoShapes()
    {
        var result = Distribute(AlignVerb.DistributeH, new[]
        {
            Make.Shape(1, 100, 10, 50, 20),
            Make.Shape(2, 400, 10, 50, 20)
        });

        Assert.False(result.Succeeded);
        Assert.Contains("three", result.Diagnostics.Single(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DistributeH_CentreMode_EqualisesThePitch()
    {
        // Centres start at 25, ?, 325. Even pitch puts the middle centre at 175.
        var result = Distribute(AlignVerb.DistributeH, new[]
        {
            Make.Shape(1, 0, 10, 50, 20),
            Make.Shape(2, 120, 10, 80, 20),
            Make.Shape(3, 300, 10, 50, 20)
        }, mode: DistributeMode.Centre);

        Assert.Equal(25, result.FrameOf(1).CentreX, Tolerance);
        Assert.Equal(175, result.FrameOf(2).CentreX, Tolerance);
        Assert.Equal(325, result.FrameOf(3).CentreX, Tolerance);
    }

    [Fact]
    public void DistributeH_CentreMode_WithExactSpacing_UsesItAsThePitch()
    {
        var result = Distribute(AlignVerb.DistributeH, new[]
        {
            Make.Shape(1, 0, 10, 50, 20),
            Make.Shape(2, 120, 10, 80, 20),
            Make.Shape(3, 300, 10, 50, 20)
        }, exactSpacing: 200, mode: DistributeMode.Centre);

        Assert.Equal(25, result.FrameOf(1).CentreX, Tolerance);
        Assert.Equal(225, result.FrameOf(2).CentreX, Tolerance);
        Assert.Equal(425, result.FrameOf(3).CentreX, Tolerance);
    }

    [Fact]
    public void DistributeH_AcrossTheSlide_SpreadsToTheSlideEdges()
    {
        var result = Distribute(AlignVerb.DistributeH, new[]
        {
            Make.Shape(1, 400, 10, 50, 20),
            Make.Shape(2, 450, 10, 50, 20),
            Make.Shape(3, 500, 10, 50, 20)
        }, reference: ReferenceTarget.Slide);

        Assert.Equal(0, result.FrameOf(1).Left, Tolerance);
        Assert.Equal(960, result.FrameOf(3).Right, Tolerance);
        Assert.Equal(455, result.FrameOf(2).Left, Tolerance);   // (960 - 150) / 2 = 405 gap each
    }

    [Fact]
    public void DistributeH_AcrossTheMargins_RespectsTheInset()
    {
        var result = Distribute(AlignVerb.DistributeH, new[]
        {
            Make.Shape(1, 400, 10, 50, 20),
            Make.Shape(2, 450, 10, 50, 20),
            Make.Shape(3, 500, 10, 50, 20)
        }, reference: ReferenceTarget.SlideMargins, slide: Make.Slide(margin: 36));

        Assert.Equal(36, result.FrameOf(1).Left, Tolerance);
        Assert.Equal(924, result.FrameOf(3).Right, Tolerance);
    }

    [Fact]
    public void DistributeH_ReportsOverlapWhenTheShapesDoNotFit()
    {
        // Three 200-wide shapes crammed into a 300pt span cannot avoid overlapping.
        var result = Distribute(AlignVerb.DistributeH, new[]
        {
            Make.Shape(1, 0, 10, 200, 20),
            Make.Shape(2, 50, 10, 200, 20),
            Make.Shape(3, 100, 10, 200, 20)
        });

        Assert.True(result.Succeeded);
        Assert.Contains(result.Diagnostics, d => d.Contains("overlap", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void DistributeH_SortsByPositionRegardlessOfInputOrder()
    {
        var result = Distribute(AlignVerb.DistributeH, new[]
        {
            Make.Shape(3, 300, 10, 50, 20),
            Make.Shape(1, 0, 10, 50, 20),
            Make.Shape(2, 120, 10, 50, 20)
        });

        // Shape 1 is leftmost and must stay put; shape 3 is rightmost and must stay put.
        Assert.Equal(0, result.FrameOf(1).Left, Tolerance);
        Assert.Equal(150, result.FrameOf(2).Left, Tolerance);
        Assert.Equal(300, result.FrameOf(3).Left, Tolerance);
    }

    [Fact]
    public void DistributeV_EqualisesTheGapsVertically()
    {
        var result = Distribute(AlignVerb.DistributeV, new[]
        {
            Make.Shape(1, 10, 0, 20, 50),
            Make.Shape(2, 10, 120, 20, 50),
            Make.Shape(3, 10, 300, 20, 50)
        });

        Assert.Equal(0, result.FrameOf(1).Top, Tolerance);
        Assert.Equal(150, result.FrameOf(2).Top, Tolerance);
        Assert.Equal(300, result.FrameOf(3).Top, Tolerance);
    }

    [Fact]
    public void DistributeH_WithVisualBounds_SpacesWhatTheEyeSees()
    {
        // The middle shape is rotated, so its visual box is wider than its frame and the gaps must be
        // computed from the wider box.
        var result = Distribute(AlignVerb.DistributeH, new[]
        {
            Make.Shape(1, 0, 100, 50, 50),
            Make.Shape(2, 200, 100, 100, 50, rotation: 45),
            Make.Shape(3, 500, 100, 50, 50)
        }, model: BoundsModel.VisualBounds);

        var visualWidth = (100 + 50) * Math.Cos(Math.PI / 4);
        var expectedGap = (550 - 50 - visualWidth - 50) / 2;

        // Middle shape's visual left edge should sit one gap past shape 1's right edge.
        var middleVisualLeft = result.FrameOf(2).CentreX - visualWidth / 2;
        Assert.Equal(50 + expectedGap, middleVisualLeft, 1e-9);
    }

    // -- reference-point modes ---------------------------------------------------------------------

    /// <summary>
    /// Three shapes of different widths, so all four modes disagree. Spans 0..460 throughout:
    /// lefts 0 / 150 / 400, rights 100 / 190 / 460, centres 50 / 170 / 430.
    /// </summary>
    private static ShapeSnapshot[] UnequalWidths() => new[]
    {
        Make.Shape(1, 0, 10, 100, 20),
        Make.Shape(2, 150, 10, 40, 20),
        Make.Shape(3, 400, 10, 60, 20)
    };

    [Theory]
    [InlineData(DistributeMode.LeadingEdge, 200)]   // lefts 0, 200, 400 - pitch 200
    [InlineData(DistributeMode.Centre, 220)]        // centres 50, 240, 430 - pitch 190
    [InlineData(DistributeMode.TrailingEdge, 240)]  // rights 100, 280, 460 - pitch 180
    [InlineData(DistributeMode.Gap, 230)]           // 260pt of free space, 130 per gap
    public void DistributeH_TheFourModesGiveFourDifferentAnswers(DistributeMode mode, double expectedMiddleLeft)
    {
        var result = Distribute(AlignVerb.DistributeH, UnequalWidths(), mode: mode);

        Assert.True(result.Succeeded);
        Assert.Equal(expectedMiddleLeft, result.FrameOf(2).Left, Tolerance);
    }

    [Theory]
    [InlineData(DistributeMode.LeadingEdge)]
    [InlineData(DistributeMode.Centre)]
    [InlineData(DistributeMode.TrailingEdge)]
    [InlineData(DistributeMode.Gap)]
    public void DistributeH_EveryModeHoldsTheOutermostShapesStill(DistributeMode mode)
    {
        var result = Distribute(AlignVerb.DistributeH, UnequalWidths(), mode: mode);

        Assert.Equal(0, result.FrameOf(1).Left, Tolerance);
        Assert.Equal(460, result.FrameOf(3).Right, Tolerance);
    }

    [Fact]
    public void DistributeH_EveryModeAgreesWhenTheShapesAreTheSameSize()
    {
        var shapes = new[]
        {
            Make.Shape(1, 0, 10, 50, 20),
            Make.Shape(2, 130, 10, 50, 20),
            Make.Shape(3, 300, 10, 50, 20)
        };

        var lefts = new[]
        {
            DistributeMode.LeadingEdge, DistributeMode.Centre,
            DistributeMode.TrailingEdge, DistributeMode.Gap
        }.Select(m => Distribute(AlignVerb.DistributeH, shapes, mode: m).FrameOf(2).Left).ToList();

        Assert.All(lefts, left => Assert.Equal(150, left, Tolerance));
    }

    [Fact]
    public void DistributeH_LeadingEdge_WithExactSpacing_StepsLeftEdges()
    {
        var result = Distribute(AlignVerb.DistributeH, UnequalWidths(),
            mode: DistributeMode.LeadingEdge, exactSpacing: 120);

        Assert.Equal(0, result.FrameOf(1).Left, Tolerance);
        Assert.Equal(120, result.FrameOf(2).Left, Tolerance);
        Assert.Equal(240, result.FrameOf(3).Left, Tolerance);
    }

    [Fact]
    public void DistributeH_TrailingEdge_WithExactSpacing_StepsRightEdges()
    {
        var result = Distribute(AlignVerb.DistributeH, UnequalWidths(),
            mode: DistributeMode.TrailingEdge, exactSpacing: 120);

        Assert.Equal(100, result.FrameOf(1).Right, Tolerance);
        Assert.Equal(220, result.FrameOf(2).Right, Tolerance);
        Assert.Equal(340, result.FrameOf(3).Right, Tolerance);
    }

    [Fact]
    public void DistributeH_TrailingEdge_OrdersByTrailingEdgeNotLeadingEdge()
    {
        // A starts left of B but its right edge is further right, so leading- and trailing-edge
        // order disagree. Distributing right edges must use the trailing-edge order: B, A, C.
        var shapes = new[]
        {
            Make.Shape(1, 0, 10, 300, 20),    // right = 300
            Make.Shape(2, 50, 40, 100, 20),   // right = 150
            Make.Shape(3, 500, 70, 50, 20)    // right = 550
        };

        var result = Distribute(AlignVerb.DistributeH, shapes, mode: DistributeMode.TrailingEdge);

        // Rights run 150, 350, 550 - pitch 200. Only A moves, from right 300 to 350.
        Assert.Equal(150, result.FrameOf(2).Right, Tolerance);
        Assert.Equal(350, result.FrameOf(1).Right, Tolerance);
        Assert.Equal(550, result.FrameOf(3).Right, Tolerance);
    }

    [Fact]
    public void DistributeH_LeadingEdge_AcrossTheSlide_RunsEdgeToEdge()
    {
        var result = Distribute(AlignVerb.DistributeH, UnequalWidths(),
            mode: DistributeMode.LeadingEdge, reference: ReferenceTarget.Slide);

        // First flush left at 0, last flush right at 960 so its left is 900. Pitch 450.
        Assert.Equal(0, result.FrameOf(1).Left, Tolerance);
        Assert.Equal(450, result.FrameOf(2).Left, Tolerance);
        Assert.Equal(900, result.FrameOf(3).Left, Tolerance);
    }

    [Fact]
    public void DistributeH_TrailingEdge_AcrossTheSlide_RunsEdgeToEdge()
    {
        var result = Distribute(AlignVerb.DistributeH, UnequalWidths(),
            mode: DistributeMode.TrailingEdge, reference: ReferenceTarget.Slide);

        // First flush left at 0 so its right is 100; last flush right at 960. Pitch 430.
        Assert.Equal(100, result.FrameOf(1).Right, Tolerance);
        Assert.Equal(530, result.FrameOf(2).Right, Tolerance);
        Assert.Equal(960, result.FrameOf(3).Right, Tolerance);
    }

    [Theory]
    [InlineData(DistributeMode.LeadingEdge, 200)]
    [InlineData(DistributeMode.Centre, 220)]
    [InlineData(DistributeMode.TrailingEdge, 240)]
    [InlineData(DistributeMode.Gap, 230)]
    public void DistributeV_MirrorsTheHorizontalModesOnTheVerticalAxis(DistributeMode mode, double expectedMiddleTop)
    {
        var shapes = new[]
        {
            Make.Shape(1, 10, 0, 20, 100),
            Make.Shape(2, 10, 150, 20, 40),
            Make.Shape(3, 10, 400, 20, 60)
        };

        var result = Distribute(AlignVerb.DistributeV, shapes, mode: mode);

        Assert.Equal(expectedMiddleTop, result.FrameOf(2).Top, Tolerance);
    }

    [Fact]
    public void DistributeV_TrailingEdge_StepsBottomEdges()
    {
        var shapes = new[]
        {
            Make.Shape(1, 10, 0, 20, 100),
            Make.Shape(2, 10, 150, 20, 40),
            Make.Shape(3, 10, 400, 20, 60)
        };

        var result = Distribute(AlignVerb.DistributeV, shapes, mode: DistributeMode.TrailingEdge);

        Assert.Equal(100, result.FrameOf(1).Bottom, Tolerance);
        Assert.Equal(280, result.FrameOf(2).Bottom, Tolerance);
        Assert.Equal(460, result.FrameOf(3).Bottom, Tolerance);
    }

    [Fact]
    public void DistributeH_NotesThatAnAnchorDoesNotDefineASpan()
    {
        var result = AlignSolver.Solve(
            new AlignRequest(AlignVerb.DistributeH, ReferenceTarget.Anchor, anchor: Make.Key(1)),
            new[]
            {
                Make.Shape(1, 0, 10, 50, 20),
                Make.Shape(2, 120, 10, 50, 20),
                Make.Shape(3, 300, 10, 50, 20)
            },
            Make.Slide());

        Assert.True(result.Succeeded);
        Assert.Contains(result.Diagnostics, d => d.Contains("span", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(150, result.FrameOf(2).Left, Tolerance);
    }
}
