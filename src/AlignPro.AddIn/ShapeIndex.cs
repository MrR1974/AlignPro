using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AlignPro.AddIn
{
    /// <summary>
    /// The shapes one operation read, by id, held until it is done. Disposing releases them all.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The reader keeps every shape it reads here, and the operation writes through the same objects,
    /// so nothing is looked up twice. Measured on a 105-shape slide: fetching a shape cost about 4ms,
    /// by position or by walking the collection, so indexing the whole slide again before writing
    /// took most of a second even to move one shape.
    /// </para>
    /// <para>
    /// The objects stay valid while shapes are moved, resized and restacked - only positions in the
    /// collections go stale after a restack, and nothing here relies on positions.
    /// </para>
    /// </remarks>
    internal sealed class ShapeIndex : IDisposable
    {
        private readonly Dictionary<int, PowerPoint.Shape> _byId = new Dictionary<int, PowerPoint.Shape>();
        private readonly List<object> _owned = new List<object>();

        /// <summary>The shape with this id, or null when it was not read. Owned by the index.</summary>
        public PowerPoint.Shape? Find(int id) => _byId.TryGetValue(id, out var shape) ? shape : null;

        /// <summary>Holds a shape under its id until the index is disposed.</summary>
        public void Keep(PowerPoint.Shape shape, int id)
        {
            _owned.Add(shape);
            _byId[id] = shape;
        }

        public void Dispose()
        {
            for (var i = _owned.Count - 1; i >= 0; i--) Com.Release(_owned[i]);
            _owned.Clear();
            _byId.Clear();
        }
    }
}
