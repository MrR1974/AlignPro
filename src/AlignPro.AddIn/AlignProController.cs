using System;
using System.Runtime.InteropServices;
using System.Globalization;
using AlignPro.Geometry;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AlignPro.AddIn
{
    /// <summary>The result of a ribbon command, in a form the ribbon can report without knowing why.</summary>
    internal sealed class CommandResult
    {
        private CommandResult(bool succeeded, string? message)
        {
            Succeeded = succeeded;
            Message = message;
        }

        public bool Succeeded { get; }

        /// <summary>Why nothing happened, or what was skipped. Null when there is nothing to say.</summary>
        public string? Message { get; }

        public static CommandResult Ok(string? message = null) => new CommandResult(true, message);

        public static CommandResult Failed(string message) => new CommandResult(false, message);
    }

    /// <summary>
    /// Orchestrates one operation: read the selection, solve, apply, record for undo. Holds the
    /// settings the ribbon's dropdowns bind to.
    /// </summary>
    internal sealed class AlignProController
    {
        private readonly PowerPoint.Application _app;

        public AlignProController(PowerPoint.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public UndoManager Undo { get; } = new UndoManager();

        public ReferenceTarget Reference { get; set; } = ReferenceTarget.SelectionBounds;

        public BoundsModel Bounds { get; set; } = BoundsModel.ShapeFrame;

        public DistributeMode DistributeMode { get; set; } = DistributeMode.Gap;

        /// <summary>Null means "even out whatever space is already there".</summary>
        public double? ExactSpacing { get; set; }

        /// <summary>Inset for the slide-margins reference, in points.</summary>
        public double Margin { get; set; } = 36;

        public ResizeOrigin ResizeOrigin { get; set; } = ResizeOrigin.TopLeft;

        /// <summary>Null means a near-square grid.</summary>
        public int? GridColumns { get; set; }

        /// <summary>Tracks which slide the undo stack belongs to; keys are only valid within one.</summary>
        private int _undoSlideId;

        public CommandResult Run(AlignVerb verb, string label)
        {
            var selection = SelectionReader.TryRead(_app, Margin, out var problem);
            if (selection == null) return CommandResult.Failed(problem ?? "Nothing to align.");

            // Match-size and grid are only meaningful against an anchor and a rectangle respectively,
            // so supply a sensible reference rather than refusing on a technicality.
            var reference = Reference;
            if (IsMatchSize(verb)) reference = ReferenceTarget.Anchor;

            var request = new AlignRequest(
                verb,
                reference,
                Bounds,
                selection.Anchor,
                ExactSpacing,
                DistributeMode,
                ResizeOrigin,
                allowGroupResize: false,
                gridColumns: GridColumns);

            var solved = AlignSolver.Solve(request, selection.Shapes, selection.Slide);
            if (!solved.Succeeded)
            {
                return CommandResult.Failed(string.Join(" ", solved.Diagnostics));
            }

            var transaction = AlignTransaction.FromResult(label, solved);
            if (transaction.IsEmpty)
            {
                return CommandResult.Ok("Everything was already in place.");
            }

            var outcome = ChangeApplier.Apply(_app, selection.SlideId, transaction.Changes);
            if (outcome.Applied == 0)
            {
                return CommandResult.Failed("None of the selected shapes could be moved.");
            }

            // Keys only mean anything within one slide, so a slide change invalidates the stack.
            if (_undoSlideId != selection.SlideId)
            {
                Undo.Clear();
                _undoSlideId = selection.SlideId;
            }
            Undo.Push(transaction);
            Diagnostics.Log(string.Format(
                CultureInfo.InvariantCulture,
                "Ran '{0}': applied={1} missing={2} slideId={3} undoDepth={4}",
                label, outcome.Applied, outcome.Missing, selection.SlideId, Undo.UndoDepth));

            return CommandResult.Ok(BuildNote(solved, outcome));
        }

        /// <summary>
        /// Whether we should claim PowerPoint's Undo. True only when our stack holds something for the
        /// slide currently in view, so editing elsewhere in the deck keeps native undo intact.
        /// </summary>
        /// <remarks>
        /// This cannot be perfect. PowerPoint exposes no "document changed" event, so if the user makes
        /// a manual edit on this slide after an AlignPro command, Ctrl+Z reverses our command rather
        /// than their edit. That is recoverable - Redo puts it back - and is a far better failure than
        /// native undo discarding an unbounded coalesced entry.
        /// </remarks>
        public bool CanUndoOnCurrentSlide() =>
            Undo.CanUndo && TryGetActiveSlideId() == _undoSlideId;

        /// <summary>
        /// Everything needed to tell apart the ways interception can decline: an empty stack, a slide
        /// mismatch, or an unreadable active slide.
        /// </summary>
        public string DescribeUndoState()
        {
            var active = TryGetActiveSlideId();
            return string.Format(
                CultureInfo.InvariantCulture,
                "canUndo={0} undoDepth={1} activeSlideId={2} undoSlideId={3} next='{4}'",
                Undo.CanUndo,
                Undo.UndoDepth,
                active?.ToString(CultureInfo.InvariantCulture) ?? "null",
                _undoSlideId,
                Undo.NextUndoLabel ?? "-");
        }

        public bool CanRedoOnCurrentSlide() =>
            Undo.CanRedo && TryGetActiveSlideId() == _undoSlideId;

        private int? TryGetActiveSlideId()
        {
            PowerPoint.DocumentWindow? window = null;
            PowerPoint.Slide? slide = null;
            try
            {
                if (_app.Windows.Count == 0) return null;

                window = _app.ActiveWindow;
                slide = window.View.Slide as PowerPoint.Slide;
                return slide?.SlideID;
            }
            catch (COMException)
            {
                // No slide in view - the master, a notes page, or a slideshow.
                return null;
            }
            finally
            {
                Com.Release(slide);
                Com.Release(window);
            }
        }

        public CommandResult UndoLast()
        {
            if (!Undo.TryUndo(out var inverse)) return CommandResult.Failed("Nothing for AlignPro to undo.");

            var outcome = ChangeApplier.Apply(_app, _undoSlideId, inverse.Changes);
            return outcome.Applied == 0
                ? CommandResult.Failed("The shapes from that operation are no longer on the slide.")
                : CommandResult.Ok(Skipped(outcome));
        }

        public CommandResult RedoLast()
        {
            if (!Undo.TryRedo(out var transaction)) return CommandResult.Failed("Nothing for AlignPro to redo.");

            var outcome = ChangeApplier.Apply(_app, _undoSlideId, transaction.Changes);
            return outcome.Applied == 0
                ? CommandResult.Failed("The shapes from that operation are no longer on the slide.")
                : CommandResult.Ok(Skipped(outcome));
        }

        private static bool IsMatchSize(AlignVerb verb) =>
            verb == AlignVerb.MatchWidth || verb == AlignVerb.MatchHeight || verb == AlignVerb.MatchBoth;

        private static string? BuildNote(SolveResult solved, ApplyOutcome outcome)
        {
            var diagnostics = solved.Diagnostics.Count > 0 ? string.Join(" ", solved.Diagnostics) : null;
            var skipped = Skipped(outcome);

            if (diagnostics == null) return skipped;
            return skipped == null ? diagnostics : diagnostics + " " + skipped;
        }

        private static string? Skipped(ApplyOutcome outcome) =>
            outcome.Missing == 0
                ? null
                : outcome.Missing == 1
                    ? "1 shape could not be changed."
                    : outcome.Missing + " shapes could not be changed.";
    }
}
