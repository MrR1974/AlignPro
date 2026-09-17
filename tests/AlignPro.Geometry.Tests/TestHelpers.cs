namespace AlignPro.Geometry.Tests;

/// <summary>Terse builders so the tests read as geometry rather than as constructor calls.</summary>
internal static class Make
{
    public const int SlideId = 1;

    public static ShapeSnapshot Shape(
        int id,
        double x,
        double y,
        double w,
        double h,
        double rotation = 0,
        RectD? textBounds = null,
        bool isGroup = false,
        bool flipH = false,
        bool flipV = false) =>
        new(
            new ShapeKey(SlideId, id),
            new RectD(x, y, w, h),
            rotation,
            flipH,
            flipV,
            textBounds,
            isGroup,
            isPlaceholder: false,
            name: $"S{id}");

    public static ShapeKey Key(int id) => new(SlideId, id);

    /// <summary>A 960x540 widescreen slide, matching what the probe measured.</summary>
    public static SlideMetrics Slide(double margin = 0, RectD? placeholder = null) =>
        new(960, 540, margin, placeholder);
}

internal static class SolveResultExtensions
{
    /// <summary>The new frame for one shape. Fails the test if the solver did not touch it.</summary>
    public static RectD FrameOf(this SolveResult result, int id)
    {
        var change = result.Changes.SingleOrDefault(c => c.Key == Make.Key(id));
        Assert.NotNull(change);
        return change!.NewFrame;
    }

    public static bool Touched(this SolveResult result, int id) =>
        result.Changes.Any(c => c.Key == Make.Key(id));
}
