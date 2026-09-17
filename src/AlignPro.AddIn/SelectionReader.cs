using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using AlignPro.Geometry;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AlignPro.AddIn
{
    /// <summary>What the solver needs, read out of PowerPoint in one pass.</summary>
    internal sealed class SelectionSnapshot
    {
        public SelectionSnapshot(
            IReadOnlyList<ShapeSnapshot> shapes, SlideMetrics slide, int slideId, ShapeKey? anchor)
        {
            Shapes = shapes;
            Slide = slide;
            SlideId = slideId;
            Anchor = anchor;
        }

        public IReadOnlyList<ShapeSnapshot> Shapes { get; }

        public SlideMetrics Slide { get; }

        public int SlideId { get; }

        /// <summary>
        /// The last shape in the selection. PowerPoint preserves selection order (measured, probe 2),
        /// so this is the shape the user selected last.
        /// </summary>
        public ShapeKey? Anchor { get; }
    }

    /// <summary>
    /// Turns the live PowerPoint selection into immutable snapshots. This is the only place that reads
    /// shape geometry, so the solver never sees a COM object.
    /// </summary>
    internal static class SelectionReader
    {
        // MsoShapeType values used here, named to keep the intent readable.
        private const int MsoGroup = 6;
        private const int MsoPlaceholder = 14;

        /// <summary>
        /// Reads the current selection. Returns null and sets <paramref name="problem"/> when there is
        /// nothing usable selected - that is an ordinary outcome, not an error.
        /// </summary>
        public static SelectionSnapshot? TryRead(PowerPoint.Application app, double margin, out string? problem)
        {
            problem = null;

            PowerPoint.DocumentWindow? window = null;
            PowerPoint.Selection? selection = null;
            PowerPoint.ShapeRange? range = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Presentation? presentation = null;

            try
            {
                if (app.Windows.Count == 0)
                {
                    problem = "Open a presentation first.";
                    return null;
                }

                window = app.ActiveWindow;
                selection = window.Selection;

                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    problem = "Select one or more shapes on a slide.";
                    return null;
                }

                range = selection.ShapeRange;
                if (range.Count == 0)
                {
                    problem = "Select one or more shapes on a slide.";
                    return null;
                }

                slide = window.View.Slide as PowerPoint.Slide;
                if (slide == null)
                {
                    problem = "AlignPro works on slides, not on the master or notes pages.";
                    return null;
                }

                presentation = slide.Parent as PowerPoint.Presentation;
                if (presentation == null)
                {
                    problem = "Could not read the presentation's slide size.";
                    return null;
                }

                var slideId = slide.SlideID;
                var metrics = new SlideMetrics(
                    presentation.PageSetup.SlideWidth,
                    presentation.PageSetup.SlideHeight,
                    margin,
                    ReadPlaceholderBounds(slide));

                var shapes = new List<ShapeSnapshot>(range.Count);
                ShapeKey? anchor = null;

                for (var i = 1; i <= range.Count; i++)
                {
                    PowerPoint.Shape? shape = null;
                    try
                    {
                        shape = range[i];
                        var snapshot = ReadShape(shape, slideId);
                        shapes.Add(snapshot);
                        anchor = snapshot.Key;   // ends up holding the last one
                    }
                    finally
                    {
                        Release(shape);
                    }
                }

                return new SelectionSnapshot(shapes, metrics, slideId, anchor);
            }
            finally
            {
                Release(presentation);
                Release(slide);
                Release(range);
                Release(selection);
                Release(window);
            }
        }

        private static ShapeSnapshot ReadShape(PowerPoint.Shape shape, int slideId)
        {
            // Read each property exactly once. Every access is a COM call, and chained expressions
            // leave RCWs we cannot release.
            var id = shape.Id;
            var left = shape.Left;
            var top = shape.Top;
            var width = shape.Width;
            var height = shape.Height;
            var rotation = shape.Rotation;
            var type = (int)shape.Type;
            var name = shape.Name;

            var flipH = SafeFlip(() => shape.HorizontalFlip);
            var flipV = SafeFlip(() => shape.VerticalFlip);

            return new ShapeSnapshot(
                new ShapeKey(slideId, id),
                // Guard the size: RectD rejects negatives, and a degenerate shape should not throw.
                new RectD(left, top, Math.Max(0, width), Math.Max(0, height)),
                rotation,
                flipH,
                flipV,
                ReadTextBounds(shape),
                isGroup: type == MsoGroup,
                isPlaceholder: type == MsoPlaceholder,
                name: name);
        }

        /// <summary>
        /// Text bounds from <c>TextRange2.Bound*</c>, already in slide coordinates. Returns null for a
        /// shape with no text frame or no text, which the solver treats as "fall back to the frame".
        /// </summary>
        private static RectD? ReadTextBounds(PowerPoint.Shape shape)
        {
            // A mixed pair: PowerPoint defines its own TextFrame2, but its TextRange is the shared
            // Office TextRange2. Getting either namespace wrong here fails to compile.
            PowerPoint.TextFrame2? frame = null;
            Office.TextRange2? text = null;
            try
            {
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue) return null;

                frame = shape.TextFrame2;
                if (frame.HasText != Office.MsoTriState.msoTrue) return null;

                text = frame.TextRange;
                var width = text.BoundWidth;
                var height = text.BoundHeight;
                if (width <= 0 || height <= 0) return null;

                return new RectD(text.BoundLeft, text.BoundTop, width, height);
            }
            catch (COMException)
            {
                // Some shape types claim a text frame then refuse to measure it. Not worth failing over.
                return null;
            }
            finally
            {
                Release(text);
                Release(frame);
            }
        }

        /// <summary>
        /// The body placeholder rectangle from the slide's layout, for the content-placeholder
        /// reference. Null when the layout has no body placeholder.
        /// </summary>
        private static RectD? ReadPlaceholderBounds(PowerPoint.Slide slide)
        {
            const int ppPlaceholderBody = 2;
            const int ppPlaceholderObject = 7;

            PowerPoint.CustomLayout? layout = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Placeholders? placeholders = null;
            try
            {
                layout = slide.CustomLayout;
                shapes = layout.Shapes;
                placeholders = shapes.Placeholders;

                for (var i = 1; i <= placeholders.Count; i++)
                {
                    PowerPoint.Shape? placeholder = null;
                    PowerPoint.PlaceholderFormat? format = null;
                    try
                    {
                        placeholder = placeholders[i];
                        format = placeholder.PlaceholderFormat;
                        var type = (int)format.Type;
                        if (type != ppPlaceholderBody && type != ppPlaceholderObject) continue;

                        return new RectD(
                            placeholder.Left,
                            placeholder.Top,
                            Math.Max(0, placeholder.Width),
                            Math.Max(0, placeholder.Height));
                    }
                    finally
                    {
                        Release(format);
                        Release(placeholder);
                    }
                }

                return null;
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                Release(placeholders);
                Release(shapes);
                Release(layout);
            }
        }

        private static bool SafeFlip(Func<Office.MsoTriState> read)
        {
            try
            {
                return read() == Office.MsoTriState.msoTrue;
            }
            catch (COMException)
            {
                // Lines and a few other shape types have no flip state.
                return false;
            }
        }

        private static void Release(object? comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
    }
}
