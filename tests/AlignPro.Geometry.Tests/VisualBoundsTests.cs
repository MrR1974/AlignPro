namespace AlignPro.Geometry.Tests;

/// <summary>
/// Probe 1 measured that the object model's rectangle ignores rotation and that a shape's centre is
/// invariant under it. Everything here follows from those two facts, so these tests are the contract
/// between the measurement and the rest of the engine.
/// </summary>
public class VisualBoundsTests
{
    private const double Tolerance = 1e-9;

    [Fact]
    public void Unrotated_MatchesTheFrame()
    {
        var shape = Make.Shape(1, 100, 100, 100, 50);

        Assert.Equal(shape.Frame, shape.VisualBounds);
    }

    [Fact]
    public void Rotated45_ExpandsToTheDiagonalExtent()
    {
        var shape = Make.Shape(1, 100, 100, 100, 50, rotation: 45);

        // (100 + 50) * cos(45) in both axes.
        var expected = (100 + 50) * Math.Cos(Math.PI / 4);
        Assert.Equal(expected, shape.VisualBounds.Width, Tolerance);
        Assert.Equal(expected, shape.VisualBounds.Height, Tolerance);
    }

    [Fact]
    public void Rotated90_SwapsTheExtents()
    {
        var shape = Make.Shape(1, 100, 100, 100, 50, rotation: 90);

        Assert.Equal(50, shape.VisualBounds.Width, Tolerance);
        Assert.Equal(100, shape.VisualBounds.Height, Tolerance);
    }

    [Theory]
    [InlineData(180)]
    [InlineData(360)]
    [InlineData(-180)]
    public void HalfTurns_MatchTheFrame(double rotation)
    {
        var shape = Make.Shape(1, 100, 100, 100, 50, rotation);

        Assert.Equal(shape.Frame.Width, shape.VisualBounds.Width, Tolerance);
        Assert.Equal(shape.Frame.Height, shape.VisualBounds.Height, Tolerance);
    }

    [Theory]
    [InlineData(45, -45)]
    [InlineData(90, 270)]
    [InlineData(30, 150)]
    public void MirroredAngles_GiveTheSameBounds(double a, double b)
    {
        var first = Make.Shape(1, 100, 100, 100, 50, a).VisualBounds;
        var second = Make.Shape(1, 100, 100, 100, 50, b).VisualBounds;

        Assert.Equal(first.Width, second.Width, Tolerance);
        Assert.Equal(first.Height, second.Height, Tolerance);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(17)]
    [InlineData(45)]
    [InlineData(90)]
    [InlineData(123.456)]
    [InlineData(-33)]
    public void CentreIsInvariantUnderRotation(double rotation)
    {
        var shape = Make.Shape(1, 100, 100, 100, 50, rotation);

        Assert.Equal(shape.Frame.CentreX, shape.VisualBounds.CentreX, Tolerance);
        Assert.Equal(shape.Frame.CentreY, shape.VisualBounds.CentreY, Tolerance);
    }

    [Fact]
    public void VisualBoundsNeverShrinkBelowTheFrame()
    {
        for (var angle = 0; angle < 360; angle += 7)
        {
            var shape = Make.Shape(1, 0, 0, 100, 50, angle);
            var area = shape.VisualBounds.Width * shape.VisualBounds.Height;

            Assert.True(area >= 100 * 50 - Tolerance, $"rotation {angle} shrank the bounding box");
        }
    }

    [Fact]
    public void Flips_DoNotAffectBounds()
    {
        var plain = Make.Shape(1, 100, 100, 100, 50, rotation: 30);
        var flipped = Make.Shape(1, 100, 100, 100, 50, rotation: 30, flipH: true, flipV: true);

        Assert.Equal(plain.VisualBounds, flipped.VisualBounds);
    }
}
