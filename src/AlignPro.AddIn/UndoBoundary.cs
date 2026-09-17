using System;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AlignPro.AddIn
{
    /// <summary>
    /// Makes PowerPoint's own undo behave correctly for AlignPro operations.
    /// </summary>
    /// <remarks>
    /// <para>
    /// PowerPoint coalesces object-model changes into a single undo entry and keeps that entry open
    /// until a <em>modifying</em> command arrives from the UI. A ribbon click is not one. The measured
    /// consequence is that an AlignPro operation joins whatever entry is already open - which can
    /// reach back over an unbounded amount of earlier work - so one Ctrl+Z discards all of it. It
    /// removed four slides during testing.
    /// </para>
    /// <para>
    /// Repurposing the built-in Undo command would have been the tidy fix, but PowerPoint ignores
    /// <c>customUI</c> command repurposing: the callback is never invoked. Instead we close the group
    /// ourselves before writing anything, so PowerPoint's entry contains exactly one AlignPro
    /// operation. Ctrl+Z, the ribbon Undo button and the Quick Access Toolbar button then all behave,
    /// with no keyboard hook anywhere.
    /// </para>
    /// <para>
    /// The mechanism is a formatting toggle followed by a real undo of that toggle. Only a modifying
    /// command closes the group, and a real undo restores the previous formatting exactly - unlike a
    /// second toggle, which is lossy when the selection's formatting is mixed. See
    /// <c>docs/object-model-findings.md</c> and <c>tools/Probe-UndoGrouping.ps1</c>.
    /// </para>
    /// </remarks>
    internal static class UndoBoundary
    {
        /// <summary>
        /// Closes PowerPoint's undo coalescing group so that changes made next form their own entry.
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
                // Read the selection's italic state first. If it cannot be read, do not toggle at all:
                // without a before-and-after comparison there is no safe way to know whether the
                // toggle created an undo entry.
                if (!TryReadItalic(app, out var before)) return false;

                app.CommandBars.ExecuteMso("Italic");

                if (!TryReadItalic(app, out var after))
                {
                    // The toggle may have landed but we cannot prove it. Issuing Undo here would risk
                    // reverting the user's own work, so accept a stray italic instead - visible and
                    // recoverable, rather than silent and destructive.
                    Diagnostics.Log("UndoBoundary: state unreadable after toggle; leaving it alone.");
                    return false;
                }

                if (after == before)
                {
                    // Nothing changed, so no undo entry was created. Undoing now would reach past us.
                    Diagnostics.Log("UndoBoundary: toggle had no effect; no boundary established.");
                    return false;
                }

                // Safe: our own toggle is demonstrably the newest entry, so this undoes exactly it.
                app.CommandBars.ExecuteMso("Undo");
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

        /// <summary>
        /// Reads the italic state of the whole selection. A mixed selection reports
        /// <c>msoTriStateMixed</c>, which is what makes the before-and-after comparison work: toggling
        /// a mixed selection moves it to a definite value, and that is a detectable change.
        /// </summary>
        private static bool TryReadItalic(PowerPoint.Application app, out Office.MsoTriState state)
        {
            state = Office.MsoTriState.msoTriStateMixed;

            PowerPoint.DocumentWindow? window = null;
            PowerPoint.Selection? selection = null;
            PowerPoint.ShapeRange? range = null;
            PowerPoint.TextFrame2? frame = null;
            Office.TextRange2? text = null;
            Office.Font2? font = null;

            try
            {
                if (app.Windows.Count == 0) return false;

                window = app.ActiveWindow;
                selection = window.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes) return false;

                range = selection.ShapeRange;
                frame = range.TextFrame2;
                text = frame.TextRange;
                font = text.Font;
                state = font.Italic;
                return true;
            }
            catch (COMException)
            {
                // Pictures, lines and some placeholders have no usable text formatting.
                return false;
            }
            finally
            {
                Com.Release(font);
                Com.Release(text);
                Com.Release(frame);
                Com.Release(range);
                Com.Release(selection);
                Com.Release(window);
            }
        }
    }
}
