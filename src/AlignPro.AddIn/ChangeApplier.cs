using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using AlignPro.Geometry;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AlignPro.AddIn
{
    /// <summary>How much of an operation actually reached the document.</summary>
    internal sealed class ApplyOutcome
    {
        public ApplyOutcome(int applied, int missing)
        {
            Applied = applied;
            Missing = missing;
        }

        public int Applied { get; }

        /// <summary>Shapes that could not be written - deleted since the snapshot, or locked.</summary>
        public int Missing { get; }
    }

    /// <summary>
    /// Writes new frames and angles back to PowerPoint, and restacks. With <see cref="ShapeCreator"/>,
    /// the only place in the add-in that mutates the document.
    /// </summary>
    internal static class ChangeApplier
    {
        /// <summary>
        /// Applies changes to the shapes the operation read. A shape that has since gone, or cannot be
        /// written, is skipped rather than aborting the rest.
        /// </summary>
        public static ApplyOutcome Apply(
            PowerPoint.Application app, ShapeIndex live, IReadOnlyList<GeometryChange> changes)
        {
            if (changes.Count == 0) return new ApplyOutcome(0, 0);

            // Close PowerPoint's open undo entry first, so the writes below form one entry of their
            // own rather than joining whatever automation has already accumulated. See UndoBoundary.
            UndoBoundary.TryClose(app);

            // Everything that has to be read is read before the first write: once a shape on the
            // slide in view has changed, PowerPoint finishes redrawing it before answering any read,
            // so reading between writes made each write wait for the last one's redraw. Measured on a
            // 105-shape slide: 40 writes took 0.7 to 2.2s with a read before each, and 51 to 69ms
            // with none.
            var missing = 0;
            var work = new List<(PowerPoint.Shape Shape, GeometryChange Change, bool Locked)>(changes.Count);
            foreach (var change in changes)
            {
                var shape = live.Find(change.Key.ShapeId);
                if (shape == null)
                {
                    missing++;
                    continue;
                }

                work.Add((shape, change, Resizes(change.OldFrame, change.NewFrame) && IsAspectLocked(shape)));
            }

            var applied = 0;
            foreach (var (shape, change, locked) in work)
            {
                try
                {
                    ApplyChange(shape, change, locked);
                    applied++;
                }
                catch (COMException)
                {
                    // A deleted, locked or otherwise unwritable shape should not abort the rest.
                    missing++;
                }
            }

            return new ApplyOutcome(applied, missing);
        }

        /// <summary>
        /// Restacks the slide, or the shapes inside one group, from <paramref name="currentOrder"/>
        /// into <paramref name="targetOrder"/>, both back to front.
        /// </summary>
        /// <remarks>
        /// <para>
        /// <c>ZOrderPosition</c> is read-only, so the only lever PowerPoint offers is "bring this one
        /// to the front". Applied to each shape in turn from the back of the target order forwards,
        /// that lands on exactly the order asked for: every shape brought to the front lands above the
        /// one before it, so after the last call the stack reads as the list does.
        /// </para>
        /// <para>
        /// Shapes at the back that are already where the target wants them are left alone: bringing
        /// everything after them to the front leaves them at the back, in their order.
        /// </para>
        /// <para>
        /// The shapes are the ones the reader held while reading the stack, so a restack - which
        /// renumbers the collection it acts in, slide or group - cannot send a call to the wrong shape.
        /// </para>
        /// </remarks>
        public static ApplyOutcome ApplyOrder(
            PowerPoint.Application app,
            ShapeIndex live,
            IReadOnlyList<ShapeKey> currentOrder,
            IReadOnlyList<ShapeKey> targetOrder)
        {
            if (targetOrder.Count == 0) return new ApplyOutcome(0, 0);

            // As in Apply: one native undo entry per AlignPro operation, restacking included.
            UndoBoundary.TryClose(app);

            const int msoBringToFront = 0;

            var settled = 0;
            while (settled < targetOrder.Count && settled < currentOrder.Count &&
                   targetOrder[settled] == currentOrder[settled])
            {
                settled++;
            }

            var applied = 0;
            var missing = 0;
            for (var i = settled; i < targetOrder.Count; i++)
            {
                var shape = live.Find(targetOrder[i].ShapeId);
                if (shape == null)
                {
                    missing++;
                    continue;
                }

                try
                {
                    shape.ZOrder((Office.MsoZOrderCmd)msoBringToFront);
                    applied++;
                }
                catch (COMException)
                {
                    // A locked shape should not abort the rest of the stack.
                    missing++;
                }
            }

            return new ApplyOutcome(applied, missing);
        }

        /// <summary>
        /// Writes one change: the angle when the change carries a new one, then whichever parts of the
        /// frame differ. Nothing is read from the shape - the change already says what it was - so a
        /// run of writes never waits on PowerPoint redrawing the one before.
        /// </summary>
        /// <param name="locked">
        /// Whether the shape's aspect ratio is locked, read before any write. A locked shape - every
        /// picture, by default - rescales its height when its width is set and the other way round,
        /// so writing one would undo the other; the lock is lifted for the size writes and put back.
        /// </param>
        internal static void ApplyChange(PowerPoint.Shape shape, GeometryChange change, bool locked)
        {
            if (change.ChangesRotation)
            {
                shape.Rotation = (float)GeometryChange.NormaliseAngle(change.NewRotation!.Value);
            }

            var from = change.OldFrame;
            var to = change.NewFrame;

            var resizeWidth = Differs(from.Width, to.Width);
            var resizeHeight = Differs(from.Height, to.Height);
            if (resizeWidth || resizeHeight)
            {
                if (locked) shape.LockAspectRatio = Office.MsoTriState.msoFalse;
                try
                {
                    if (resizeWidth) shape.Width = (float)to.Width;
                    if (resizeHeight) shape.Height = (float)to.Height;
                }
                finally
                {
                    if (locked) shape.LockAspectRatio = Office.MsoTriState.msoTrue;
                }
            }

            // Position last: setting a size holds the top-left corner, so it must not come after.
            // Written whenever the size changed too, since a resize can nudge a locked shape's corner.
            if (resizeWidth || resizeHeight || Differs(from.X, to.X)) shape.Left = (float)to.X;
            if (resizeWidth || resizeHeight || Differs(from.Y, to.Y)) shape.Top = (float)to.Y;
        }

        /// <summary>Whether a change alters the size, so the aspect lock matters.</summary>
        private static bool Resizes(RectD from, RectD to) => Differs(from.Width, to.Width) || Differs(from.Height, to.Height);

        /// <summary>Whether the aspect ratio is locked; false for shapes that have no such setting.</summary>
        private static bool IsAspectLocked(PowerPoint.Shape shape)
        {
            try
            {
                return shape.LockAspectRatio == Office.MsoTriState.msoTrue;
            }
            catch (COMException)
            {
                return false;
            }
        }

        private static bool Differs(double a, double b) => Math.Abs(a - b) > 1e-4;
    }
}
