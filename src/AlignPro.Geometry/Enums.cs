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

    /// <summary>What "evenly spaced" means for the distribute verbs.</summary>
    public enum DistributeMode
    {
        /// <summary>Equalise the edge-to-edge gaps between neighbours.</summary>
        Gap,

        /// <summary>Equalise the centre-to-centre pitch.</summary>
        Centre
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
