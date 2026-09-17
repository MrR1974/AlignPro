namespace AlignPro.Geometry.Tests;

public class RectDTests
{
    [Fact]
    public void FromEdges_ProducesOriginAndSize()
    {
        var rect = RectD.FromEdges(10, 20, 110, 70);

        Assert.Equal(10, rect.Left);
        Assert.Equal(20, rect.Top);
        Assert.Equal(100, rect.Width);
        Assert.Equal(50, rect.Height);
        Assert.Equal(110, rect.Right);
        Assert.Equal(70, rect.Bottom);
    }

    [Fact]
    public void FromCentre_CentresTheRect()
    {
        var rect = RectD.FromCentre(150, 125, 100, 50);

        Assert.Equal(100, rect.Left);
        Assert.Equal(100, rect.Top);
        Assert.Equal(150, rect.CentreX);
        Assert.Equal(125, rect.CentreY);
    }

    [Fact]
    public void Union_CoversBothOperands()
    {
        var a = new RectD(0, 0, 10, 10);
        var b = new RectD(50, 60, 10, 10);

        Assert.Equal(RectD.FromEdges(0, 0, 60, 70), a.Union(b));
        Assert.Equal(a.Union(b), b.Union(a));
    }

    [Fact]
    public void Deflate_CollapsesRatherThanInverting()
    {
        // Six points of inset off each side of a ten-point square would invert a naive implementation.
        var rect = new RectD(0, 0, 10, 10).Deflate(6);

        Assert.Equal(0, rect.Width);
        Assert.Equal(0, rect.Height);
        Assert.Equal(6, rect.Left);
        Assert.Equal(6, rect.Top);
    }

    [Fact]
    public void Deflate_AppliesPerEdgeInsets()
    {
        var rect = new RectD(0, 0, 100, 100).Deflate(10, 20, 30, 40);

        Assert.Equal(new RectD(10, 20, 60, 40), rect);
    }

    [Theory]
    [InlineData(-1, 10)]
    [InlineData(10, -1)]
    public void NegativeSize_Throws(double width, double height)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new RectD(0, 0, width, height));
    }

    [Fact]
    public void IsDegenerate_TrueWhenAnAxisHasNoExtent()
    {
        Assert.True(new RectD(0, 0, 0, 10).IsDegenerate);
        Assert.True(new RectD(0, 0, 10, 0).IsDegenerate);
        Assert.False(new RectD(0, 0, 10, 10).IsDegenerate);
    }

    [Fact]
    public void Equality_ToleratesFloatingPointNoise()
    {
        var a = new RectD(100, 100, 50, 50);
        var b = new RectD(100 + 1e-9, 100, 50, 50);

        Assert.Equal(a, b);
        Assert.True(a == b);
        Assert.Equal(a.GetHashCode(), b.GetHashCode());
    }

    [Fact]
    public void Equality_DistinguishesRealDifferences()
    {
        Assert.NotEqual(new RectD(100, 100, 50, 50), new RectD(100.1, 100, 50, 50));
    }

    [Fact]
    public void Offset_MovesWithoutResizing()
    {
        var rect = new RectD(10, 10, 100, 50).Offset(5, -5);

        Assert.Equal(new RectD(15, 5, 100, 50), rect);
    }
}
