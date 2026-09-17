namespace AlignPro.Geometry.Tests;

public class AlignTests
{
    private const double Tolerance = 1e-9;

    private static SolveResult Align(
        AlignVerb verb,
        ReferenceTarget reference = ReferenceTarget.SelectionBounds,
        BoundsModel model = BoundsModel.ShapeFrame,
        ShapeKey? anchor = null,
        SlideMetrics? slide = null,
        params ShapeSnapshot[] shapes) =>
        AlignSolver.Solve(
            new AlignRequest(verb, reference, model, anchor),
            shapes,
            slide ?? Make.Slide());

    // -- against the selection's own bounds --------------------------------------------------------

    [Fact]
    public void AlignLeft_MovesEverythingToTheLeftmostEdge()
    {
        var result = Align(AlignVerb.AlignLeft, shapes: new[]
        {
            Make.Shape(1, 100, 10, 50, 20),
            Make.Shape(2, 300, 50, 80, 20),
            Make.Shape(3, 200, 90, 60, 20)
        });

        Assert.True(result.Succeeded);
        Assert.Equal(100, result.FrameOf(1).Left);
        Assert.Equal(100, result.FrameOf(2).Left);
        Assert.Equal(100, result.FrameOf(3).Left);
    }

    [Fact]
    public void AlignRight_MovesEverythingToTheRightmostEdge()
    {
        var result = Align(AlignVerb.AlignRight, shapes: new[]
        {
            Make.Shape(1, 100, 10, 50, 20),   // right = 150
            Make.Shape(2, 300, 50, 80, 20)    // right = 380
        });

        Assert.Equal(380, result.FrameOf(1).Right, Tolerance);
        Assert.Equal(380, result.FrameOf(2).Right, Tolerance);
    }

    [Fact]
    public void AlignCentreH_CentresOnTheSelectionCentre()
    {
        var result = Align(AlignVerb.AlignCentreH, shapes: new[]
        {
            Make.Shape(1, 100, 10, 50, 20),   // spans 100..150
            Make.Shape(2, 300, 50, 80, 20)    // spans 300..380
        });

        // Union spans 100..380, centre 240.
        Assert.Equal(240, result.FrameOf(1).CentreX, Tolerance);
        Assert.Equal(240, result.FrameOf(2).CentreX, Tolerance);
    }

    [Fact]
    public void AlignTop_And_AlignBottom_WorkOnTheVerticalAxis()
    {
        var shapes = new[] { Make.Shape(1, 10, 100, 20, 50), Make.Shape(2, 50, 300, 20, 80) };

        var top = Align(AlignVerb.AlignTop, shapes: shapes);
        Assert.Equal(100, top.FrameOf(1).Top);
        Assert.Equal(100, top.FrameOf(2).Top);

        var bottom = Align(AlignVerb.AlignBottom, shapes: shapes);
        Assert.Equal(380, bottom.FrameOf(1).Bottom, Tolerance);
        Assert.Equal(380, bottom.FrameOf(2).Bottom, Tolerance);
    }

    [Fact]
    public void AlignCentreV_CentresOnTheSelectionMiddle()
    {
        var result = Align(AlignVerb.AlignCentreV, shapes: new[]
        {
            Make.Shape(1, 10, 100, 20, 50),
            Make.Shape(2, 50, 300, 20, 80)
        });

        Assert.Equal(240, result.FrameOf(1).CentreY, Tolerance);
        Assert.Equal(240, result.FrameOf(2).CentreY, Tolerance);
    }

    [Fact]
    public void Align_LeavesSizeAndRotationAlone()
    {
        var result = Align(AlignVerb.AlignLeft, shapes: new[]
        {
            Make.Shape(1, 100, 10, 50, 20),
            Make.Shape(2, 300, 50, 80, 30, rotation: 37)
        });

        Assert.Equal(80, result.FrameOf(2).Width);
        Assert.Equal(30, result.FrameOf(2).Height);
    }

    // -- against the slide, its margins, and a placeholder -----------------------------------------

    [Fact]
    public void AlignLeft_ToSlide_UsesTheSlideEdge()
    {
        var result = Align(AlignVerb.AlignLeft, ReferenceTarget.Slide, shapes: new[]
        {
            Make.Shape(1, 100, 10, 50, 20)
        });

        Assert.Equal(0, result.FrameOf(1).Left);
    }

    [Fact]
    public void AlignLeft_ToSlideMargins_UsesTheInsetEdge()
    {
        var result = Align(
            AlignVerb.AlignLeft, ReferenceTarget.SlideMargins, slide: Make.Slide(margin: 36),
            shapes: new[] { Make.Shape(1, 500, 10, 50, 20) });

        Assert.Equal(36, result.FrameOf(1).Left);
    }

    [Fact]
    public void AlignCentreH_ToSlide_CentresOnTheSlide()
    {
        var result = Align(AlignVerb.AlignCentreH, ReferenceTarget.Slide, shapes: new[]
        {
            Make.Shape(1, 0, 10, 100, 20)
        });

        Assert.Equal(480, result.FrameOf(1).CentreX);
    }

    [Fact]
    public void AlignLeft_ToPlaceholder_UsesTheLayoutRect()
    {
        var slide = Make.Slide(placeholder: new RectD(120, 88, 720, 188));

        var result = Align(
            AlignVerb.AlignLeft, ReferenceTarget.PlaceholderBounds, slide: slide,
            shapes: new[] { Make.Shape(1, 500, 10, 50, 20) });

        Assert.Equal(120, result.FrameOf(1).Left);
    }

    [Fact]
    public void AlignLeft_ToPlaceholder_RefusesWhenTheLayoutHasNone()
    {
        var result = Align(
            AlignVerb.AlignLeft, ReferenceTarget.PlaceholderBounds,
            shapes: new[] { Make.Shape(1, 500, 10, 50, 20) });

        Assert.False(result.Succeeded);
        Assert.Contains("placeholder", result.Diagnostics.Single(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AlignLeft_ToSlide_AcceptsASingleShape()
    {
        var result = Align(AlignVerb.AlignLeft, ReferenceTarget.Slide, shapes: new[]
        {
            Make.Shape(1, 100, 10, 50, 20)
        });

        Assert.True(result.Succeeded);
    }

    [Fact]
    public void AlignLeft_ToSelectionBounds_RefusesASingleShape()
    {
        var result = Align(AlignVerb.AlignLeft, shapes: new[] { Make.Shape(1, 100, 10, 50, 20) });

        Assert.False(result.Succeeded);
    }

    [Fact]
    public void Align_RefusesAnEmptySelection()
    {
        var result = AlignSolver.Solve(new AlignRequest(AlignVerb.AlignLeft), Array.Empty<ShapeSnapshot>(), Make.Slide());

        Assert.False(result.Succeeded);
    }

    // -- against an anchor -------------------------------------------------------------------------

    [Fact]
    public void AlignLeft_ToAnchor_MovesTheOthersAndLeavesTheAnchorAlone()
    {
        var result = Align(
            AlignVerb.AlignLeft, ReferenceTarget.Anchor, anchor: Make.Key(2),
            shapes: new[]
            {
                Make.Shape(1, 100, 10, 50, 20),
                Make.Shape(2, 300, 50, 80, 20),
                Make.Shape(3, 500, 90, 60, 20)
            });

        Assert.True(result.Succeeded);
        Assert.False(result.Touched(2));
        Assert.Equal(300, result.FrameOf(1).Left);
        Assert.Equal(300, result.FrameOf(3).Left);
    }

    [Fact]
    public void AlignLeft_ToAnchor_RefusesWithoutAnAnchor()
    {
        var result = Align(AlignVerb.AlignLeft, ReferenceTarget.Anchor, shapes: new[]
        {
            Make.Shape(1, 100, 10, 50, 20),
            Make.Shape(2, 300, 50, 80, 20)
        });

        Assert.False(result.Succeeded);
        Assert.Contains("anchor", result.Diagnostics.Single(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AlignLeft_ToAnchor_RefusesWhenTheAnchorIsNotSelected()
    {
        var result = Align(
            AlignVerb.AlignLeft, ReferenceTarget.Anchor, anchor: Make.Key(99),
            shapes: new[] { Make.Shape(1, 100, 10, 50, 20), Make.Shape(2, 300, 50, 80, 20) });

        Assert.False(result.Succeeded);
    }

    // -- bounds models -----------------------------------------------------------------------------

    [Fact]
    public void AlignLeft_WithVisualBounds_MakesARotatedShapeVisuallyFlush()
    {
        // A 100x50 frame at 45 degrees has a visual box of 106.066 wide centred on (150, 125),
        // so its visual left edge is 96.967 - further left than its frame claims.
        var rotated = Make.Shape(1, 100, 100, 100, 50, rotation: 45);
        var plain = Make.Shape(2, 300, 200, 60, 40);
        var expectedVisualLeft = 150 - (100 + 50) * Math.Cos(Math.PI / 4) / 2;

        var visual = Align(AlignVerb.AlignLeft, model: BoundsModel.VisualBounds, shapes: new[] { rotated, plain });
        Assert.Equal(expectedVisualLeft, visual.FrameOf(2).Left, 1e-9);

        // The frame model is what native PowerPoint does, and it lands 3.03pt off.
        var frame = Align(AlignVerb.AlignLeft, shapes: new[] { rotated, plain });
        Assert.Equal(100, frame.FrameOf(2).Left);
    }

    [Fact]
    public void AlignLeft_WithVisualBounds_LeavesTheRotatedShapeWhereItIs()
    {
        var rotated = Make.Shape(1, 100, 100, 100, 50, rotation: 45);
        var plain = Make.Shape(2, 300, 200, 60, 40);

        var result = Align(AlignVerb.AlignLeft, model: BoundsModel.VisualBounds, shapes: new[] { rotated, plain });

        Assert.Equal(rotated.Frame, result.FrameOf(1));
    }

    [Fact]
    public void AlignLeft_WithTextBounds_AlignsTheTextNotTheFrame()
    {
        // Frames start level at 300, but the text inside sits at different insets.
        var a = Make.Shape(1, 300, 100, 200, 40, textBounds: new RectD(307.2, 103.6, 70, 21.6));
        var b = Make.Shape(2, 300, 200, 200, 40, textBounds: new RectD(320.0, 203.6, 70, 21.6));

        var result = Align(AlignVerb.AlignLeft, model: BoundsModel.TextBounds, shapes: new[] { a, b });

        // b's text must come back to 307.2, so its frame shifts left by 12.8.
        Assert.Equal(287.2, result.FrameOf(2).Left, 1e-9);
        Assert.Equal(300, result.FrameOf(1).Left);
    }

    [Fact]
    public void AlignLeft_WithTextBounds_FallsBackToTheFrameAndSaysSo()
    {
        var withText = Make.Shape(1, 300, 100, 200, 40, textBounds: new RectD(307.2, 103.6, 70, 21.6));
        var without = Make.Shape(2, 400, 200, 60, 40);

        var result = Align(AlignVerb.AlignLeft, model: BoundsModel.TextBounds, shapes: new[] { withText, without });

        Assert.True(result.Succeeded);
        Assert.Contains(result.Diagnostics, d => d.Contains("no text", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(307.2, result.FrameOf(2).Left, 1e-9);
    }
}
