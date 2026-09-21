using System;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AlignPro.AddIn
{
    /// <summary>
    /// Makes PowerPoint's own undo behave correctly for AlignPro operations.
    /// </summary>
    /// <remarks>
    /// <para>
    /// PowerPoint's built-in commands each bracket their own undo entry, which is why native Align
    /// left followed by native Align middle undoes one step at a time. Object-model writes get no
    /// such bracket: they accumulate into a single open entry until something ends it. A click on
    /// <em>our</em> ribbon button is not a PowerPoint command - it is a callback into managed code
    /// that then writes through the object model - so an AlignPro operation joins whatever entry is
    /// already open, which can reach back over an unbounded amount of earlier work. One Ctrl+Z
    /// discards all of it. It removed four slides during testing.
    /// </para>
    /// <para>
    /// <see cref="PowerPoint._Application.StartNewUndoEntry"/> is PowerPoint's own answer to this: it
    /// closes the open entry so that changes made next form their own. Call it before writing and
    /// PowerPoint's entry contains exactly one AlignPro operation, so Ctrl+Z, the ribbon Undo button
    /// and the Quick Access Toolbar button all behave, with no keyboard hook anywhere.
    /// </para>
    /// <para>
    /// Two earlier attempts failed and are worth not repeating. Repurposing the built-in Undo through
    /// <c>customUI</c> is parsed and then ignored by PowerPoint - the callback is never invoked.
    /// Closing the group with a formatting toggle followed by <c>ExecuteMso("Undo")</c> works from
    /// outside PowerPoint but not from a ribbon callback, because PowerPoint defers a <em>command</em>
    /// while another is executing, so that undo landed after the geometry writes and reverted them.
    /// <c>StartNewUndoEntry</c> avoids both traps by being an object-model method rather than a
    /// command: like the <c>Left</c>/<c>Top</c> writes and <c>Shape.ZOrder</c>, it executes in place.
    /// See <c>docs/object-model-findings.md</c>.
    /// </para>
    /// </remarks>
    internal static class UndoBoundary
    {
        /// <summary>
        /// Closes PowerPoint's open undo entry so that changes made next form their own.
        /// </summary>
        /// <returns>True when a boundary was established.</returns>
        /// <remarks>
        /// Never throws. Failing to establish a boundary only costs undo granularity, so it must not
        /// stop the operation the user asked for.
        /// </remarks>
        public static bool TryClose(PowerPoint.Application app)
        {
            try
            {
                app.StartNewUndoEntry();
                return true;
            }
            catch (COMException ex)
            {
                Diagnostics.Log("UndoBoundary: " + ex.Message);
                return false;
            }
            catch (Exception ex)
            {
                Diagnostics.Log("UndoBoundary unexpected: " + ex.Message);
                return false;
            }
        }
    }
}
