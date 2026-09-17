namespace AlignPro.Geometry.Tests;

/// <summary>
/// Reproduces the nine scattered circles from the sample deck's grid slide. They collapsed onto one
/// another in PowerPoint, so these tests pin down what the solver actually produces for that input.
/// </summary>
public class GridScatterTests
{
    /// <summary>The sample deck's slide 9: nine 80x80 circles, deliberately untidy.</summary>
    private static ShapeSnapshot[] NineCircles() => new[]
    {
        Make.Shape(1, 120, 150, 80, 80),
        Make.Shape(2, 430, 190, 80, 80),
        Make.Shape(3, 760, 140, 80, 80),
        Make.Shape(4, 190, 300, 80, 80),
        Make.Shape(5, 520, 260, 80, 80),
        Make.Shape(6, 690, 330, 80, 80),
        Make.Shape(7, 110, 420, 80, 80),
        Make.Shape(8, 400, 400, 80, 80),
        Make.Shape(9, 800, 430, 80, 80)
    };

    private static SolveResult Grid(ShapeSnapshot[] shapes, int? columns = null) =>
        AlignSolver.Solve(
            new AlignRequest(AlignVerb.GridArrange, gridColumns: columns), shapes, Make.Slide());

    [Fact]
    public void NineShapes_LandOnNineDistinctPositions()
    {
        var result = Grid(NineCircles());

        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));

        var centres = result.Changes
            .Select(c => $"{c.NewFrame.CentreX:F1},{c.NewFrame.CentreY:F1}")
            .ToList();

        Assert.Equal(9, centres.Count);
        Assert.Equal(9, centres.Distinct().Count());
    }

    [Fact]
    public void NineShapes_FormThreeRowsAndThreeColumns()
    {
        var result = Grid(NineCircles());

        var xs = result.Changes.Select(c => Math.Round(c.NewFrame.CentreX, 1)).Distinct().OrderBy(x => x).ToList();
        var ys = result.Changes.Select(c => Math.Round(c.NewFrame.CentreY, 1)).Distinct().OrderBy(y => y).ToList();

        Assert.Equal(3, xs.Count);
        Assert.Equal(3, ys.Count);
    }

    [Fact]
    public void ShapesStayNearTheirOriginalRowAndColumn()
    {
        // The point of a tidy is that things end up near where they were. The shape that started
        // top-left should finish top-left, and the one that started bottom-right should finish
        // bottom-right - not swapped across the slide.
        var shapes = NineCircles();
        var result = Grid(shapes);

        var byKey = result.Changes.ToDictionary(c => c.Key, c => c.NewFrame);

        var topLeftStart = shapes.OrderBy(s => s.Frame.CentreX + s.Frame.CentreY).First();
        var bottomRightStart = shapes.OrderByDescending(s => s.Frame.CentreX + s.Frame.CentreY).First();

        var allX = byKey.Values.Select(f => f.CentreX).ToList();
        var allY = byKey.Values.Select(f => f.CentreY).ToList();

        Assert.Equal(allX.Min(), byKey[topLeftStart.Key].CentreX, 1);
        Assert.Equal(allY.Min(), byKey[topLeftStart.Key].CentreY, 1);
        Assert.Equal(allX.Max(), byKey[bottomRightStart.Key].CentreX, 1);
        Assert.Equal(allY.Max(), byKey[bottomRightStart.Key].CentreY, 1);
    }

    [Fact]
    public void ExplicitColumnCountIsHonoured()
    {
        var result = Grid(NineCircles(), columns: 9);

        var xs = result.Changes.Select(c => Math.Round(c.NewFrame.CentreX, 1)).Distinct().ToList();
        var ys = result.Changes.Select(c => Math.Round(c.NewFrame.CentreY, 1)).Distinct().ToList();

        Assert.Equal(9, xs.Count);
        Assert.Single(ys);
    }
}
