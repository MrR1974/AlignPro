using System;
using System.Globalization;

namespace AlignPro.Geometry
{
    /// <summary>
    /// An axis-aligned rectangle in PowerPoint points (1/72"), measured from the top-left of the
    /// slide with Y increasing downwards - the same convention as the PowerPoint object model.
    /// </summary>
    public readonly struct RectD : IEquatable<RectD>
    {
        /// <summary>Tolerance for geometric comparison, in points. Well below one screen pixel.</summary>
        public const double Epsilon = 1e-6;

        public RectD(double x, double y, double width, double height)
        {
            if (width < 0) throw new ArgumentOutOfRangeException(nameof(width), width, "Width cannot be negative.");
            if (height < 0) throw new ArgumentOutOfRangeException(nameof(height), height, "Height cannot be negative.");

            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        public double X { get; }
        public double Y { get; }
        public double Width { get; }
        public double Height { get; }

        public double Left => X;
        public double Top => Y;
        public double Right => X + Width;
        public double Bottom => Y + Height;
        public double CentreX => X + Width / 2;
        public double CentreY => Y + Height / 2;

        /// <summary>True when the rectangle has no area, so edge alignment is still meaningful but
        /// distribution by gap is not.</summary>
        public bool IsDegenerate => Width <= Epsilon || Height <= Epsilon;

        /// <summary>Builds a rectangle from edges rather than an origin and size.</summary>
        public static RectD FromEdges(double left, double top, double right, double bottom) =>
            new RectD(left, top, right - left, bottom - top);

        /// <summary>Builds a rectangle of the given size around a centre point.</summary>
        public static RectD FromCentre(double centreX, double centreY, double width, double height) =>
            new RectD(centreX - width / 2, centreY - height / 2, width, height);

        /// <summary>Translates by a delta, leaving the size untouched.</summary>
        public RectD Offset(double dx, double dy) => new RectD(X + dx, Y + dy, Width, Height);

        /// <summary>Returns a copy with a new left edge, size unchanged.</summary>
        public RectD WithLeft(double left) => new RectD(left, Y, Width, Height);

        /// <summary>Returns a copy with a new top edge, size unchanged.</summary>
        public RectD WithTop(double top) => new RectD(X, top, Width, Height);

        /// <summary>Returns a copy with a new size, origin unchanged.</summary>
        public RectD WithSize(double width, double height) => new RectD(X, Y, width, height);

        /// <summary>Shrinks the rectangle inwards by the given inset on all four sides. A larger
        /// inset than the rectangle can absorb collapses that axis to zero rather than inverting.</summary>
        public RectD Deflate(double inset) => Deflate(inset, inset, inset, inset);

        /// <summary>Shrinks the rectangle inwards by a per-edge inset, collapsing rather than
        /// inverting when the insets exceed the available extent.</summary>
        public RectD Deflate(double left, double top, double right, double bottom)
        {
            var newWidth = Math.Max(0, Width - left - right);
            var newHeight = Math.Max(0, Height - top - bottom);
            return new RectD(X + left, Y + top, newWidth, newHeight);
        }

        /// <summary>The smallest rectangle containing both operands.</summary>
        public RectD Union(RectD other) => FromEdges(
            Math.Min(Left, other.Left),
            Math.Min(Top, other.Top),
            Math.Max(Right, other.Right),
            Math.Max(Bottom, other.Bottom));

        /// <summary>True when every edge matches within <see cref="Epsilon"/>.</summary>
        public bool ApproximatelyEquals(RectD other, double tolerance = Epsilon) =>
            Math.Abs(X - other.X) <= tolerance &&
            Math.Abs(Y - other.Y) <= tolerance &&
            Math.Abs(Width - other.Width) <= tolerance &&
            Math.Abs(Height - other.Height) <= tolerance;

        public bool Equals(RectD other) => ApproximatelyEquals(other);

        public override bool Equals(object? obj) => obj is RectD other && Equals(other);

        public override int GetHashCode()
        {
            // Quantised to the epsilon so that values comparing equal also hash equal.
            unchecked
            {
                var hash = 17;
                hash = hash * 31 + Math.Round(X, 6).GetHashCode();
                hash = hash * 31 + Math.Round(Y, 6).GetHashCode();
                hash = hash * 31 + Math.Round(Width, 6).GetHashCode();
                hash = hash * 31 + Math.Round(Height, 6).GetHashCode();
                return hash;
            }
        }

        public static bool operator ==(RectD left, RectD right) => left.Equals(right);

        public static bool operator !=(RectD left, RectD right) => !left.Equals(right);

        public override string ToString() => string.Format(
            CultureInfo.InvariantCulture,
            "L={0:F2} T={1:F2} W={2:F2} H={3:F2}",
            X, Y, Width, Height);
    }
}
