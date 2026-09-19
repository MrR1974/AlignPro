namespace AlignPro.Geometry.Tests;

/// <summary>
/// Orderings are written back to front, matching PowerPoint's own ZOrderPosition, so the last id in
/// a list is the one on top.
/// </summary>
public class ZOrderTests
{
    private static ShapeKey[] Order(params int[] ids) =>
        ids.Select(Make.Key).ToArray();

    private static int[] Ids(IReadOnlyList<ShapeKey> order) =>
        order.Select(k => k.ShapeId).ToArray();

    private static OrderResult Solve(
        int[] slide, int[] selection, OrderVerb verb = OrderVerb.StackFirstOnTop) =>
        ZOrderSolver.Solve(Order(slide), Order(selection), verb);

    [Fact]
    public void FirstSelected_EndsOnTop()
    {
        var result = Solve(slide: new[] { 1, 2, 3 }, selection: new[] { 3, 2, 1 });

        Assert.True(result.Succeeded);

        // Selected 3, then 2, then 1 - so 3 must end on top, and top is last.
        Assert.Equal(new[] { 1, 2, 3 }, Ids(result.Change!.NewOrder));
    }

    [Fact]
    public void SelectionOrderDrivesTheResult_NotTheExistingStack()
    {
        var result = Solve(slide: new[] { 1, 2, 3 }, selection: new[] { 2, 3, 1 });

        // 2 first so 2 on top, then 3, then 1 at the back.
        Assert.Equal(new[] { 1, 3, 2 }, Ids(result.Change!.NewOrder));
    }

    [Fact]
    public void UnselectedShapes_KeepTheirExactLayer()
    {
        // 9 and 8 are not selected and sit in the middle of the stack.
        var result = Solve(slide: new[] { 1, 9, 2, 8, 3 }, selection: new[] { 1, 2, 3 });

        // The selected shapes occupy slots 0, 2 and 4; 9 and 8 must still be at slots 1 and 3.
        var ids = Ids(result.Change!.NewOrder);
        Assert.Equal(9, ids[1]);
        Assert.Equal(8, ids[3]);

        // First selected on top means 1 takes the frontmost of the three slots it had between them.
        Assert.Equal(new[] { 3, 9, 2, 8, 1 }, ids);
    }

    [Fact]
    public void Reverse_PutsTheFirstSelectedAtTheBottom()
    {
        var result = Solve(
            slide: new[] { 1, 2, 3 }, selection: new[] { 3, 2, 1 }, verb: OrderVerb.StackFirstOnBottom);

        Assert.Equal(new[] { 3, 2, 1 }, Ids(result.Change!.NewOrder));
    }

    [Fact]
    public void TheOldOrderIsRecordedForUndo()
    {
        var result = Solve(slide: new[] { 1, 2, 3 }, selection: new[] { 2, 3, 1 });

        Assert.Equal(new[] { 1, 2, 3 }, Ids(result.Change!.OldOrder));
    }

    [Fact]
    public void InvertingRestoresTheOriginalOrder()
    {
        var result = Solve(slide: new[] { 1, 9, 2, 8, 3 }, selection: new[] { 1, 2, 3 });
        var inverse = result.Change!.Inverted();

        Assert.Equal(new[] { 1, 9, 2, 8, 3 }, Ids(inverse.NewOrder));
    }

    [Fact]
    public void AnOrderingThatChangesNothing_IsANoOp()
    {
        // Already stacked with 3 on top, and 3 is selected first.
        var result = Solve(slide: new[] { 1, 2, 3 }, selection: new[] { 3, 2, 1 });

        Assert.True(result.Change!.IsNoOp);
    }

    [Fact]
    public void PartialSelection_OnlyTouchesItsOwnSlots()
    {
        var result = Solve(slide: new[] { 1, 2, 3, 4, 5 }, selection: new[] { 4, 2 });

        // Slots 1 and 3 hold the selection; 4 goes to the front one, 2 to the back one.
        Assert.Equal(new[] { 1, 2, 3, 4, 5 }, Ids(result.Change!.OldOrder));
        Assert.Equal(new[] { 1, 2, 3, 4, 5 }, Ids(result.Change!.NewOrder));
    }

    [Fact]
    public void PartialSelection_SwapsWhenTheOrderAsksForIt()
    {
        var result = Solve(slide: new[] { 1, 2, 3, 4, 5 }, selection: new[] { 2, 4 });

        // 2 selected first so it must end above 4, which means they swap slots.
        Assert.Equal(new[] { 1, 4, 3, 2, 5 }, Ids(result.Change!.NewOrder));
    }

    [Fact]
    public void RefusesASingleShape()
    {
        var result = Solve(slide: new[] { 1, 2, 3 }, selection: new[] { 2 });

        Assert.False(result.Succeeded);
        Assert.Null(result.Change);
    }

    [Fact]
    public void RefusesAShapeThatIsNotOnTheSlide()
    {
        // A shape selected inside a group never appears in the slide's own collection.
        var result = Solve(slide: new[] { 1, 2, 3 }, selection: new[] { 1, 77 });

        Assert.False(result.Succeeded);
        Assert.Contains("group", result.Diagnostics.Single(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ADuplicateInTheSelectionDoesNotScrambleTheResult()
    {
        var result = Solve(slide: new[] { 1, 2, 3 }, selection: new[] { 3, 3, 2 });

        // Two distinct shapes, so two slots: 3 on top of 2, and 1 untouched at the back.
        Assert.True(result.Succeeded);
        Assert.Equal(new[] { 1, 2, 3 }, Ids(result.Change!.NewOrder));
    }

    [Fact]
    public void EveryShapeSurvivesTheReordering()
    {
        var result = Solve(slide: new[] { 5, 1, 4, 2, 3 }, selection: new[] { 3, 1, 4 });

        var ids = Ids(result.Change!.NewOrder);
        Assert.Equal(5, ids.Length);
        Assert.Equal(new[] { 1, 2, 3, 4, 5 }, ids.OrderBy(i => i).ToArray());
    }
}
