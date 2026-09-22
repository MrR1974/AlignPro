namespace AlignPro.Geometry.Tests;

public class CurveTests
{
    // Flattening puts points within a small fraction of a point of the true curve.
    private const double Close = 0.05;

    private static SolveResult Place(
        IReadOnlyList<ShapeSnapshot> shapes,
        CurveDefinition curve,
        int anchorId = 9,
        double? spacing = null,
        bool rotateShapes = true) =>
        CurveSolver.Solve(new CurveRequest(Make.Key(anchorId), curve, spacing, rotateShapes), shapes);

    /// <summary>Four 10x10 shapes, then a 200x200 circle centred on (300, 300) as the anchor.</summary>
    private static ShapeSnapshot[] FourThenCircle(double rotation = 0) => new[]
    {
        Make.Shape(1, 0, 0, 10, 10),
        Make.Shape(2, 20, 0, 10, 10),
        Make.Shape(3, 40, 0, 10, 10),
        Make.Shape(4, 60, 0, 10, 10),
        Make.Shape(9, 200, 200, 200, 200, rotation: rotation)
    };

    private static (double X, double Y) Centre(SolveResult result, int id)
    {
        var frame = result.FrameOf(id);
        return (frame.CentreX, frame.CentreY);
    }

    private static double Angle(SolveResult result, int id) =>
        result.Changes.Single(c => c.Key == Make.Key(id)).NewRotation!.Value;

    // -- ellipse --------------------------------------------------------------------------------

    [Fact]
    public void Circle_EvenSlots_StartingAtTwelveAndGoingClockwise()
    {
        var result = Place(FourThenCircle(), CurveDefinition.Ellipse());

        Assert.True(result.Succeeded);
        var expected = new[] { (300.0, 200.0), (400.0, 300.0), (300.0, 400.0), (200.0, 300.0) };
        for (var i = 0; i < 4; i++)
        {
            var (x, y) = Centre(result, i + 1);
            Assert.Equal(expected[i].Item1, x, Close);
            Assert.Equal(expected[i].Item2, y, Close);
        }
    }

    [Fact]
    public void Circle_TheLastShapeDoesNotLandOnTheFirst()
    {
        var result = Place(FourThenCircle(), CurveDefinition.Ellipse());

        var first = Centre(result, 1);
        var last = Centre(result, 4);
        Assert.True(Math.Abs(first.X - last.X) + Math.Abs(first.Y - last.Y) > 50);
    }

    [Fact]
    public void Circle_RotateShapes_FollowsTheTangent()
    {
        var result = Place(FourThenCircle(), CurveDefinition.Ellipse());

        // At twelve o'clock, travelling clockwise, the tangent points right: upright.
        Assert.Equal(0, Angle(result, 1), 0.01);
        Assert.Equal(90, Angle(result, 2), 0.01);
        Assert.Equal(180, Angle(result, 3), 0.01);
        Assert.Equal(270, Angle(result, 4), 0.01);
    }

    [Fact]
    public void RotateShapesOff_LeavesTheAngleAlone()
    {
        var result = Place(FourThenCircle(), CurveDefinition.Ellipse(), rotateShapes: false);

        Assert.All(result.Changes, c => Assert.Null(c.NewRotation));
    }

    [Fact]
    public void Ellipse_IsSpacedByArcLengthNotAngle()
    {
        // A wide, flat ellipse. Even steps of angle would bunch shapes at the narrow ends; even steps
        // of distance keep the gaps between neighbours equal.
        var shapes = Enumerable.Range(1, 8).Select(i => Make.Shape(i, 0, 0, 4, 4)).ToList();
        shapes.Add(Make.Shape(9, 0, 0, 400, 80));

        var result = Place(shapes, CurveDefinition.Ellipse());
        var centres = Enumerable.Range(1, 8).Select(i => Centre(result, i)).ToList();

        var chords = new List<double>();
        for (var i = 0; i < 8; i++)
        {
            var a = centres[i];
            var b = centres[(i + 1) % 8];
            chords.Add(Math.Sqrt((a.X - b.X) * (a.X - b.X) + (a.Y - b.Y) * (a.Y - b.Y)));
        }

        // Chords under-measure curved arcs slightly, so allow a little - but angle stepping on this
        // ellipse gives neighbours that differ by a factor of three or more.
        Assert.True(chords.Max() / chords.Min() < 1.15, string.Join(", ", chords.Select(c => c.ToString("F1"))));
    }

    [Fact]
    public void ARotatedEllipse_IsFollowedAsDrawn()
    {
        // A 200x100 ellipse turned a quarter: it now stands tall, so its twelve-o'clock point - its
        // own local top - has swung round to the right.
        var shapes = new[]
        {
            Make.Shape(1, 0, 0, 10, 10),
            Make.Shape(9, 200, 250, 200, 100, rotation: 90)
        };

        var result = Place(shapes, CurveDefinition.Ellipse());

        var (x, y) = Centre(result, 1);
        Assert.Equal(350, x, Close);
        Assert.Equal(300, y, Close);
    }

    // -- arc ------------------------------------------------------------------------------------
    // PowerPoint's Arc frames the arc plus its centre, not the whole ellipse. Every case below is a
    // frame PowerPoint itself reported in probe 10.

    private static ShapeSnapshot[] ShapesThenArc(int count, RectD frame, bool flipH = false, double rotation = 0)
    {
        var shapes = Enumerable.Range(1, count).Select(i => Make.Shape(i, 0, 0, 10, 10)).ToList();
        shapes.Add(new ShapeSnapshot(Make.Key(9), frame, rotation, flipH: flipH));
        return shapes.ToArray();
    }

    [Fact]
    public void DefaultArc_HasItsCentreAtTheFramesBottomLeft()
    {
        // The default quarter, twelve o'clock to three, drawn 200x200: a circle of radius 200 about
        // (200, 400), running from the frame's top-left corner to its bottom-right.
        var result = Place(ShapesThenArc(3, new RectD(200, 200, 200, 200)), CurveDefinition.Arc(-90, 0));

        Assert.Equal(200, Centre(result, 1).X, Close);
        Assert.Equal(200, Centre(result, 1).Y, Close);
        Assert.Equal(200 + 200 * Math.Cos(Math.PI / 4), Centre(result, 2).X, Close);
        Assert.Equal(400 - 200 * Math.Sin(Math.PI / 4), Centre(result, 2).Y, Close);
        Assert.Equal(400, Centre(result, 3).X, Close);
        Assert.Equal(400, Centre(result, 3).Y, Close);
    }

    [Fact]
    public void Arc_AnglesAreDirectionsFromTheCentre()
    {
        // Measured: a 300x100 ellipse centred on (100, 300), cut to 0..45 degrees, reports this
        // frame. The 45 degree direction meets that ellipse at (194.87, 394.87); the 45 degree
        // parameter would have been (312.1, 370.7).
        var result = Place(
            ShapesThenArc(2, new RectD(100, 300, 300, 94.86827)), CurveDefinition.Arc(0, 45), rotateShapes: false);

        Assert.Equal(400, Centre(result, 1).X, Close);
        Assert.Equal(300, Centre(result, 1).Y, Close);
        Assert.Equal(194.868, Centre(result, 2).X, Close);
        Assert.Equal(394.868, Centre(result, 2).Y, Close);
    }

    [Fact]
    public void Arc_ThatCrossesAxes_IsRebuiltToo()
    {
        // Measured: the same ellipse cut to -30..200 degrees (PowerPoint stores 200 as -160).
        var result = Place(
            ShapesThenArc(2, new RectD(-200.0501, 213.3976, 600.1002, 186.6192)), CurveDefinition.Arc(-30, -160));

        // -30 degrees as a direction on a 300x100 ellipse: radius 173.2 from (100, 300).
        Assert.Equal(100 + 173.205 * Math.Cos(Math.PI / 6), Centre(result, 1).X, 0.2);
        Assert.Equal(300 - 173.205 * Math.Sin(Math.PI / 6), Centre(result, 1).Y, 0.2);
    }

    [Fact]
    public void Arc_FlippedHorizontally_IsMirroredWithinItsFrame()
    {
        var result = Place(ShapesThenArc(2, new RectD(200, 200, 200, 200), flipH: true), CurveDefinition.Arc(-90, 0));

        // The quarter now runs from the top-right corner round to the bottom-left.
        Assert.Equal(400, Centre(result, 1).X, Close);
        Assert.Equal(200, Centre(result, 1).Y, Close);
        Assert.Equal(200, Centre(result, 2).X, Close);
        Assert.Equal(400, Centre(result, 2).Y, Close);
    }

    [Fact]
    public void ARotatedArc_IsRefused()
    {
        // A rotated arc's frame changes size in ways that do not describe its ellipse (probe 10).
        var result = Place(ShapesThenArc(2, new RectD(200, 200, 200, 200), rotation: 30), CurveDefinition.Arc(-90, 0));

        Assert.False(result.Succeeded);
        Assert.Contains("rotated arc", result.Diagnostics.Single());
    }

    // -- line -----------------------------------------------------------------------------------

    [Fact]
    public void Line_FollowsItsFlips()
    {
        // Measured: AddLine(100,400 -> 300,350) reports this frame with VerticalFlip set.
        var shapes = new[]
        {
            Make.Shape(1, 0, 0, 10, 10),
            Make.Shape(2, 0, 0, 10, 10),
            new ShapeSnapshot(Make.Key(9), new RectD(100, 350, 200, 50), flipV: true)
        };

        var result = Place(shapes, CurveDefinition.Line());

        Assert.Equal((100.0, 400.0), Centre(result, 1));
        Assert.Equal((300.0, 350.0), Centre(result, 2));
    }

    // -- path -----------------------------------------------------------------------------------

    private static CurveDefinition LShape() => CurveDefinition.Path(new[]
    {
        new PathNode(100, 100),
        new PathNode(300, 100),
        new PathNode(300, 200)
    });

    private static ShapeSnapshot[] ThreeThenPath() => new[]
    {
        Make.Shape(1, 0, 0, 10, 10),
        Make.Shape(2, 0, 0, 10, 10),
        Make.Shape(3, 0, 0, 10, 10),
        Make.Shape(9, 100, 100, 200, 100)
    };

    [Fact]
    public void Path_PlacesByDistanceAlongIt()
    {
        var result = Place(ThreeThenPath(), LShape());

        // 300pt long in all; the middle shape sits 150pt along, which is on the first leg.
        Assert.Equal((100.0, 100.0), Centre(result, 1));
        Assert.Equal(250, Centre(result, 2).X, Close);
        Assert.Equal(100, Centre(result, 2).Y, Close);
        Assert.Equal(300, Centre(result, 3).X, Close);
        Assert.Equal(200, Centre(result, 3).Y, Close);
    }

    [Fact]
    public void Path_NodesAreTakenAsDrawn_NotTurnedAgain()
    {
        // PowerPoint reports a freeform's nodes with its rotation and flips already applied (probe 9),
        // so the anchor's own angle and flip must not be applied a second time.
        var shapes = ThreeThenPath();
        shapes[3] = new ShapeSnapshot(Make.Key(9), new RectD(100, 100, 200, 100), rotation: 90, flipH: true);

        var result = Place(shapes, LShape());

        Assert.Equal((100.0, 100.0), Centre(result, 1));
        Assert.Equal((300.0, 200.0), Centre(result, 3));
    }

    [Fact]
    public void Path_ShapesTurnAtTheCorner()
    {
        var result = Place(ThreeThenPath(), LShape());

        Assert.Equal(0, Angle(result, 2), 1e-6);
        Assert.Equal(90, Angle(result, 3), 1e-6);
    }

    [Fact]
    public void Path_ExactSpacingStepsFromTheStart()
    {
        var result = Place(ThreeThenPath(), LShape(), spacing: 120);

        Assert.Equal(220, Centre(result, 2).X, Close);
        Assert.Equal(300, Centre(result, 3).X, Close);
        Assert.Equal(140, Centre(result, 3).Y, Close);
    }

    [Fact]
    public void Path_ExactSpacingThatRunsOffTheEnd_IsRefused()
    {
        var result = Place(ThreeThenPath(), LShape(), spacing: 200);

        Assert.False(result.Succeeded);
        Assert.Contains("300", result.Diagnostics.Single());
    }

    [Fact]
    public void Path_BezierSegmentsAreFollowed()
    {
        // A cubic whose control points sit at its ends is a straight line - an exact check that the
        // two middle nodes are read as controls, not as corners.
        var curve = CurveDefinition.Path(new[]
        {
            new PathNode(0, 0, PathSegmentKind.Curve),
            new PathNode(0, 0),
            new PathNode(200, 0),
            new PathNode(200, 0)
        });
        var shapes = new[]
        {
            Make.Shape(1, 0, 0, 10, 10),
            Make.Shape(2, 0, 0, 10, 10),
            Make.Shape(9, 0, 0, 200, 1)
        };

        var result = Place(shapes, curve);

        Assert.Equal(200, Centre(result, 2).X, Close);
        Assert.Equal(0, Centre(result, 2).Y, Close);
    }

    [Fact]
    public void Path_AClosedPathGetsAFullCircuit()
    {
        // A square loop, 400pt round, whose last node returns to its first.
        var curve = CurveDefinition.Path(new[]
        {
            new PathNode(0, 0), new PathNode(100, 0), new PathNode(100, 100), new PathNode(0, 100), new PathNode(0, 0)
        });
        var shapes = new[]
        {
            Make.Shape(1, 0, 0, 10, 10),
            Make.Shape(2, 0, 0, 10, 10),
            Make.Shape(3, 0, 0, 10, 10),
            Make.Shape(4, 0, 0, 10, 10),
            Make.Shape(9, 0, 0, 100, 100)
        };

        var result = Place(shapes, curve);

        // Four corners, and the fourth is not back on the first.
        Assert.Equal((0.0, 100.0), Centre(result, 4));
    }

    [Fact]
    public void ASingleShapeOnAnOpenCurve_GoesInTheMiddle()
    {
        var shapes = new[] { Make.Shape(1, 0, 0, 10, 10), Make.Shape(9, 100, 100, 200, 100) };

        var result = Place(shapes, LShape());

        Assert.Equal((250.0, 100.0), Centre(result, 1));
    }

    // -- refusals -------------------------------------------------------------------------------

    [Fact]
    public void TheCurveAloneIsRefused()
    {
        var result = Place(new[] { Make.Shape(9, 0, 0, 100, 100) }, CurveDefinition.Ellipse());

        Assert.False(result.Succeeded);
    }

    [Fact]
    public void TheCurveIsNeverMoved()
    {
        var result = Place(FourThenCircle(), CurveDefinition.Ellipse());

        Assert.False(result.Touched(9));
    }

    [Fact]
    public void ExactSpacingOnACircleThatWouldOverlap_IsRefused()
    {
        // Circumference about 628pt; four shapes at 200pt apart need 800pt to get back round.
        var result = Place(FourThenCircle(), CurveDefinition.Ellipse(), spacing: 200);

        Assert.False(result.Succeeded);
    }

    [Fact]
    public void Undo_RestoresFramesAndAngles()
    {
        var shapes = FourThenCircle();
        var result = Place(shapes, CurveDefinition.Ellipse());

        var inverse = AlignTransaction.FromResult("Distribute along curve", result).Inverted();
        var shape2 = inverse.Changes.Single(c => c.Key == Make.Key(2));

        Assert.Equal(shapes[1].Frame, shape2.NewFrame);
        Assert.Equal(0, shape2.NewRotation!.Value, 1e-9);
    }
}
