namespace AlignPro.Geometry
{
    /// <summary>The operation to perform. One verb per ribbon button.</summary>
    public enum AlignVerb
    {
        AlignLeft,
        AlignRight,
        AlignTop,
        AlignBottom,
        AlignCentreH,
        AlignCentreV,
        DistributeH,
        DistributeV,
        MatchWidth,
        MatchHeight,
        MatchBoth,
        GridArrange
    }

    /// <summary>What the verb aligns or distributes against.</summary>
    public enum ReferenceTarget
    {
        /// <summary>The designated anchor shape - by default the last one selected.</summary>
        Anchor,

        /// <summary>The union of the selection's own bounds. PowerPoint's native behaviour.</summary>
        SelectionBounds,

        /// <summary>The whole slide.</summary>
        Slide,

        /// <summary>The slide inset by <see cref="SlideMetrics.Margin"/>.</summary>
        SlideMargins,

        /// <summary>The body placeholder rectangle from the slide's layout.</summary>
        PlaceholderBounds
    }

    /// <summary>Which rectangle of a shape the verb reasons about.</summary>
    public enum BoundsModel
    {
        /// <summary>The raw object-model rectangle, which ignores rotation. PowerPoint's native behaviour.</summary>
        ShapeFrame,

        /// <summary>The rotation-aware rectangle - what the eye actually sees.</summary>
        VisualBounds,

        /// <summary>The bounds of the shape's text rather than its frame.</summary>
        TextBounds
    }

    /// <summary>
    /// Which feature of each shape the distribute verbs space evenly. The axis comes from the verb,
    /// so "leading" means the left edge for <see cref="AlignVerb.DistributeH"/> and the top edge for
    /// <see cref="AlignVerb.DistributeV"/>.
    /// </summary>
    /// <remarks>
    /// The first three space a reference <em>point</em> on each shape at a constant pitch, and give
    /// identical results when every shape is the same size. They diverge as soon as sizes differ:
    /// only <see cref="Gap"/> equalises the visible space between shapes, and only the pitch modes
    /// give a regular rhythm regardless of what sits in each slot.
    /// </remarks>
    public enum DistributeMode
    {
        /// <summary>Left edge to left edge horizontally, top edge to top edge vertically.</summary>
        LeadingEdge,

        /// <summary>Centre to centre along the verb's axis.</summary>
        Centre,

        /// <summary>Right edge to right edge horizontally, bottom edge to bottom edge vertically.</summary>
        TrailingEdge,

        /// <summary>
        /// Equalise the space between neighbours - one shape's trailing edge to the next shape's
        /// leading edge. The only mode where shapes of differing sizes end up at an irregular pitch.
        /// </summary>
        Gap
    }

    /// <summary>
    /// How the match-size verbs apply <see cref="AlignRequest.SizeMargin"/>.
    /// </summary>
    /// <remarks>
    /// The margin is measured <em>per side</em>, so one step of it takes twice that off each dimension.
    /// That convention is what makes the result concentric when paired with
    /// <see cref="ResizeOrigin.Centre"/>: a 10pt margin leaves a 10pt border visible all the way
    /// round, which is what someone nesting one shape inside another is actually asking for.
    /// </remarks>
    public enum SizeMarginMode
    {
        /// <summary>Match the anchor exactly. The original behaviour.</summary>
        None,

        /// <summary>Every shape is one step from the anchor, so they all end the same size.</summary>
        Uniform,

        /// <summary>
        /// The step grows with distance from the anchor in selection order, so the shapes tier. With
        /// the anchor selected last - the default - the first shape selected ends smallest.
        /// </summary>
        Cascade
    }

    /// <summary>Which part of a shape stays put while it is resized.</summary>
    public enum ResizeOrigin
    {
        /// <summary>The top-left corner holds. PowerPoint's native behaviour.</summary>
        TopLeft,

        /// <summary>The centre holds, so the shape grows outwards in both directions.</summary>
        Centre
    }

    /// <summary>The order cells are filled in by <see cref="AlignVerb.GridArrange"/>.</summary>
    public enum GridFillOrder
    {
        RowMajor,
        ColumnMajor
    }
}
