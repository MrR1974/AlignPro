using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace AlignPro.AddIn
{
    /// <summary>
    /// Draws AlignPro's own ribbon icons.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Office's built-in icons are referenced by <c>imageMso</c> identifier, and an identifier that
    /// does not exist renders as blank rather than raising an error - so a wrong guess is silent, and
    /// there is no reliable way to validate one from script (<c>GetEnabledMso</c> answers a different
    /// question: whether the command is available, not whether the icon exists).
    /// </para>
    /// <para>
    /// Drawing our own removes the guesswork, and lets each pictogram show what the verb actually does
    /// rather than the nearest thing Microsoft happened to ship. They are drawn in code rather than
    /// embedded as PNGs, so there are no binaries in the repository.
    /// </para>
    /// <para>
    /// Everything is designed on a 16-unit grid and multiplied by a whole number, so every edge lands
    /// on a pixel boundary at both 16px and 32px. Anti-aliasing is off for the same reason: these are
    /// axis-aligned rectangles, and smoothing only turns crisp edges to mush at small sizes.
    /// </para>
    /// </remarks>
    internal static class Icons
    {
        /// <summary>Office's own icons are dark on the light ribbon; this sits alongside them.</summary>
        private static readonly Color Ink = Color.FromArgb(68, 68, 68);

        /// <summary>Picks out the part carrying the meaning, the way Office accents its own icons.</summary>
        private static readonly Color Accent = Color.FromArgb(31, 119, 180);

        private static readonly Dictionary<string, Bitmap> Cache = new Dictionary<string, Bitmap>();

        /// <summary>
        /// The icon for a ribbon control, drawn once and reused. Returns null for a control with no
        /// custom icon, which the caller treats as "leave it alone".
        /// </summary>
        public static Bitmap? For(string controlId)
        {
            lock (Cache)
            {
                if (Cache.TryGetValue(controlId, out var cached)) return cached;

                // Large buttons are 32px, normal ones 16px. Drawing at the target size keeps thin
                // strokes crisp instead of scaling a single image and going soft.
                var bitmap = controlId switch
                {
                    "btnDistributeH" => Distribute(32, horizontal: true),
                    "btnDistributeV" => Distribute(32, horizontal: false),
                    "btnMatchWidth" => MatchWidth(16),
                    "btnMatchHeight" => MatchHeight(16),
                    "btnMatchBoth" => MatchBoth(16),
                    "tbResizeFromCentre" => ResizeFromCentre(16),
                    "btnGrid" => Grid(16),
                    _ => null
                };

                if (bitmap != null) Cache[controlId] = bitmap;
                return bitmap;
            }
        }

        private static Bitmap Create(int size, out Graphics g, out int unit)
        {
            var bitmap = new Bitmap(size, size);
            g = Graphics.FromImage(bitmap);
            g.SmoothingMode = SmoothingMode.None;
            g.PixelOffsetMode = PixelOffsetMode.None;
            g.InterpolationMode = InterpolationMode.NearestNeighbor;
            unit = Math.Max(1, size / 16);
            return bitmap;
        }

        private static void Fill(Graphics g, Brush brush, int unit, int x, int y, int w, int h) =>
            g.FillRectangle(brush, x * unit, y * unit, w * unit, h * unit);

        /// <summary>
        /// A one-unit border, drawn as four filled rectangles. A <see cref="Pen"/> would straddle the
        /// path and land on half pixels, which is exactly the softness these icons avoid.
        /// </summary>
        private static void Outline(Graphics g, Brush brush, int unit, int x, int y, int w, int h)
        {
            Fill(g, brush, unit, x, y, w, 1);
            Fill(g, brush, unit, x, y + h - 1, w, 1);
            Fill(g, brush, unit, x, y, 1, h);
            Fill(g, brush, unit, x + w - 1, y, 1, h);
        }

        /// <summary>
        /// Three bars at an even pitch, with the two equal gaps marked along the far edge. The gaps are
        /// the point of the verb, so they are what the accent picks out.
        /// </summary>
        private static Bitmap Distribute(int size, bool horizontal)
        {
            var bitmap = Create(size, out var g, out var u);
            using (g)
            using (var ink = new SolidBrush(Ink))
            using (var accent = new SolidBrush(Accent))
            {
                // Bars occupy 2-4, 7-9 and 12-14, leaving two identical three-unit gaps.
                foreach (var p in new[] { 2, 7, 12 })
                {
                    if (horizontal) Fill(g, ink, u, p, 2, 2, 9);
                    else Fill(g, ink, u, 2, p, 9, 2);
                }

                foreach (var gap in new[] { 4, 9 })
                {
                    if (horizontal) Fill(g, accent, u, gap, 13, 3, 1);
                    else Fill(g, accent, u, 13, gap, 1, 3);
                }
            }

            return bitmap;
        }

        /// <summary>Two bars sharing a width, differing in height. The shared edges are accented.</summary>
        private static Bitmap MatchWidth(int size)
        {
            var bitmap = Create(size, out var g, out var u);
            using (g)
            using (var ink = new SolidBrush(Ink))
            using (var accent = new SolidBrush(Accent))
            {
                Fill(g, accent, u, 2, 2, 12, 3);
                Fill(g, ink, u, 2, 8, 12, 6);
            }

            return bitmap;
        }

        /// <summary>Two bars sharing a height, differing in width.</summary>
        private static Bitmap MatchHeight(int size)
        {
            var bitmap = Create(size, out var g, out var u);
            using (g)
            using (var ink = new SolidBrush(Ink))
            using (var accent = new SolidBrush(Accent))
            {
                Fill(g, accent, u, 2, 2, 3, 12);
                Fill(g, ink, u, 8, 2, 6, 12);
            }

            return bitmap;
        }

        /// <summary>
        /// Two identical squares set on a diagonal. Deliberately not overlapping: at 16px an overlap
        /// reads as one muddy blob rather than as two shapes of equal size.
        /// </summary>
        private static Bitmap MatchBoth(int size)
        {
            var bitmap = Create(size, out var g, out var u);
            using (g)
            using (var ink = new SolidBrush(Ink))
            using (var accent = new SolidBrush(Accent))
            {
                Fill(g, accent, u, 1, 1, 6, 6);
                Fill(g, ink, u, 9, 9, 6, 6);
            }

            return bitmap;
        }

        /// <summary>
        /// A small shape and the larger one it becomes, sharing a centre - the toggle's whole meaning
        /// is that the centre holds while the edges move outward equally.
        /// </summary>
        /// <remarks>
        /// Concentric rectangles rather than outward arrows: at 16px an arrowhead is three or four
        /// pixels and reads as noise, whereas two boxes on a common centre stay legible.
        /// </remarks>
        private static Bitmap ResizeFromCentre(int size)
        {
            var bitmap = Create(size, out var g, out var u);
            using (g)
            using (var ink = new SolidBrush(Ink))
            using (var accent = new SolidBrush(Accent))
            {
                // Both are centred on unit 8: the outer spans 1-15, the inner 5-11.
                Outline(g, ink, u, 1, 1, 14, 14);
                Fill(g, accent, u, 5, 5, 6, 6);
            }

            return bitmap;
        }

        /// <summary>A three by three of cells: the tidy the verb produces.</summary>
        private static Bitmap Grid(int size)
        {
            var bitmap = Create(size, out var g, out var u);
            using (g)
            using (var ink = new SolidBrush(Ink))
            using (var accent = new SolidBrush(Accent))
            {
                var positions = new[] { 1, 6, 11 };
                for (var row = 0; row < 3; row++)
                {
                    for (var column = 0; column < 3; column++)
                    {
                        // One accented cell stops it reading as a plain table icon.
                        var brush = (row == 0 && column == 0) ? accent : ink;
                        Fill(g, brush, u, positions[column], positions[row], 4, 4);
                    }
                }
            }

            return bitmap;
        }
    }
}
