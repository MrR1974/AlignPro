using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using AlignPro.Geometry;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace AlignPro.AddIn
{
    /// <summary>The result of a ribbon command, in a form the ribbon can report without knowing why.</summary>
    internal sealed class CommandResult
    {
        private CommandResult(bool succeeded, string? message, bool notable)
        {
            Succeeded = succeeded;
            Message = message;
            Notable = notable;
        }

        public bool Succeeded { get; }

        /// <summary>Why nothing happened, or what was skipped. Null when there is nothing to say.</summary>
        public string? Message { get; }

        /// <summary>
        /// True when the message must actually reach the user rather than being logged quietly -
        /// a refusal, or a success that deliberately skipped part of the selection.
        /// </summary>
        public bool Notable { get; }

        public static CommandResult Ok(string? message = null) => new CommandResult(true, message, false);

        /// <summary>Succeeded, but part of what was asked for did not happen.</summary>
        public static CommandResult Note(string message) => new CommandResult(true, message, true);

        public static CommandResult Failed(string message) => new CommandResult(false, message, true);
    }

    /// <summary>Which way the match-size margin goes.</summary>
    internal enum SizeDirection
    {
        /// <summary>Each step sits inside the anchor.</summary>
        Shrink,

        /// <summary>Each step sits outside the anchor.</summary>
        Grow
    }

    /// <summary>
    /// Orchestrates one operation: read the selection, solve, apply. Holds the
    /// settings the ribbon's dropdowns bind to.
    /// </summary>
    internal sealed class AlignProController
    {
        private readonly PowerPoint.Application _app;

        public AlignProController(PowerPoint.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
        }

        public ReferenceTarget Reference { get; set; } = ReferenceTarget.SelectionBounds;

        public BoundsModel Bounds { get; set; } = BoundsModel.ShapeFrame;

        public DistributeMode DistributeMode { get; set; } = DistributeMode.Gap;

        /// <summary>Null means "even out whatever space is already there".</summary>
        public double? ExactSpacing { get; set; }

        /// <summary>Inset for the slide-margins reference, in points.</summary>
        public double Margin { get; set; } = 36;

        public ResizeOrigin ResizeOrigin { get; set; } = ResizeOrigin.TopLeft;

        /// <summary>
        /// The match-size margin, measured per side in points. Always zero or more: which way it
        /// goes is <see cref="SizeDirection"/>'s job, not the sign's.
        /// </summary>
        public double SizeMargin { get; set; }

        /// <summary>Whether the margin shrinks the shapes inside the anchor or grows them past it.</summary>
        public SizeDirection SizeDirection { get; set; } = SizeDirection.Shrink;

        /// <summary>Whether the match-size margin applies, and whether it cascades.</summary>
        public SizeMarginMode SizeMarginMode { get; set; } = SizeMarginMode.None;

        /// <summary>Null means a near-square grid.</summary>
        public int? GridColumns { get; set; }

        /// <summary>How far each duplicate step moves right, in points.</summary>
        public double DuplicateX { get; set; } = 20;

        /// <summary>How far each duplicate step moves down, in points.</summary>
        public double DuplicateY { get; set; } = 20;

        /// <summary>How far each duplicate step turns, in degrees clockwise.</summary>
        public double DuplicateAngle { get; set; }

        public int DuplicateCopies { get; set; } = 1;

        public DuplicatePivot DuplicatePivot { get; set; } = DuplicatePivot.OwnCentre;

        /// <summary>
        /// Shared by Duplicate and Distribute along curve: turn shapes with the step or the curve, or
        /// leave their angle alone.
        /// </summary>
        public bool RotateShapes { get; set; } = true;

        /// <summary>How close, in points, shapes must be for Tidy to treat them as meant to line up.</summary>
        public double TidyTolerance { get; set; } = TidyRequest.DefaultTolerance;

        public CommandResult Run(AlignVerb verb, string label)
        {
            var selection = SelectionReader.TryRead(_app, Margin, out var problem);
            if (selection == null) return CommandResult.Failed(problem ?? "Nothing to align.");

            // Match-size, match-rotation and grid are only meaningful against an anchor and a
            // rectangle respectively, so supply a sensible reference rather than refusing on a
            // technicality.
            var reference = Reference;
            if (IsMatchSize(verb) || verb == AlignVerb.MatchRotation) reference = ReferenceTarget.Anchor;

            var request = new AlignRequest(
                verb,
                reference,
                Bounds,
                selection.Anchor,
                ExactSpacing,
                DistributeMode,
                ResizeOrigin,
                allowGroupResize: false,
                sizeMargin: SignedSizeMargin,
                sizeMarginMode: SizeMarginMode,
                gridColumns: GridColumns);

            foreach (var snapshot in selection.Shapes)
            {
                Diagnostics.LogVerbose(string.Format(
                    CultureInfo.InvariantCulture,
                    "  read {0} '{1}' {2} rot={3:F0} group={4}",
                    snapshot.Key, snapshot.Name ?? "?", snapshot.Frame, snapshot.Rotation, snapshot.IsGroup));
            }

            // Inside a rotated group the solve runs along the group's own axes; anywhere else this is
            // the align solver as it always was.
            var solved = GroupSpace.Solve(request, selection.Shapes, selection.Slide, selection.GroupRotation);
            return ApplySolved(label, selection, solved);
        }

        /// <summary>
        /// Applies a geometry solve. Shared by every verb whose result is a set of frame and angle
        /// changes, whichever solver produced it.
        /// </summary>
        private CommandResult ApplySolved(string label, SelectionSnapshot selection, SolveResult solved)
        {
            if (!solved.Succeeded)
            {
                return CommandResult.Failed(string.Join(" ", solved.Diagnostics));
            }

            var changes = solved.EffectiveChanges.ToList();
            if (changes.Count == 0)
            {
                return CommandResult.Note("Nothing to change - the selection is already in place.");
            }

            foreach (var change in changes)
            {
                Diagnostics.LogVerbose("  change " + change);
            }

            var outcome = ChangeApplier.Apply(_app, selection.SlideId, changes);
            if (outcome.Applied == 0)
            {
                return CommandResult.Failed("None of the selected shapes could be moved.");
            }

            Diagnostics.Log(string.Format(
                CultureInfo.InvariantCulture,
                "Ran '{0}': applied={1} missing={2} slideId={3}",
                label, outcome.Applied, outcome.Missing, selection.SlideId));

            var note = BuildNote(solved, outcome);
            return solved.Notable || outcome.Missing > 0
                ? CommandResult.Note(note ?? "Part of the selection was skipped.")
                : CommandResult.Ok(note);
        }

        /// <summary>
        /// Restacks the selection. Parallel to <see cref="Run"/> rather than another verb inside it:
        /// this changes no geometry, so it has nothing to say to the align solver.
        /// </summary>
        public CommandResult RunOrder(OrderVerb verb, string label)
        {
            var selection = SelectionReader.TryRead(_app, Margin, out var problem, includeSlideOrder: true);
            if (selection == null) return CommandResult.Failed(problem ?? "Nothing to reorder.");

            if (selection.SlideOrder == null)
            {
                return CommandResult.Failed("Could not read the slide's stacking order.");
            }

            var keys = new List<ShapeKey>(selection.Shapes.Count);
            foreach (var snapshot in selection.Shapes) keys.Add(snapshot.Key);

            var solved = ZOrderSolver.Solve(selection.SlideOrder, keys, verb);
            if (!solved.Succeeded)
            {
                return CommandResult.Failed(string.Join(" ", solved.Diagnostics));
            }

            if (solved.Change!.IsNoOp)
            {
                return CommandResult.Note("Nothing to change - the shapes are already in that order.");
            }

            var outcome = ChangeApplier.ApplyOrder(_app, selection.SlideId, solved.Change.NewOrder);
            if (outcome.Applied == 0)
            {
                return CommandResult.Failed("None of the selected shapes could be restacked.");
            }

            Diagnostics.Log(string.Format(
                CultureInfo.InvariantCulture,
                "Ran '{0}': restacked={1} missing={2} slideId={3}",
                label, outcome.Applied, outcome.Missing, selection.SlideId));

            var skipped = Skipped(outcome);
            return skipped == null ? CommandResult.Ok() : CommandResult.Note(skipped);
        }

        /// <summary>
        /// Duplicates the selection as a unit, <see cref="DuplicateCopies"/> times, each copy one step
        /// further on. A third entry point beside <see cref="Run"/> and <see cref="RunOrder"/>,
        /// because it makes shapes rather than editing them.
        /// </summary>
        public CommandResult RunDuplicate(string label)
        {
            var selection = SelectionReader.TryRead(_app, Margin, out var problem);
            if (selection == null) return CommandResult.Failed(problem ?? "Nothing to duplicate.");

            var request = new DuplicateRequest(
                DuplicateX, DuplicateY, DuplicateAngle, DuplicateCopies, DuplicatePivot, RotateShapes, selection.Anchor);

            // Inside a rotated group the step runs along the group's own axes.
            var solved = GroupSpace.SolveDuplicate(request, selection.Shapes, selection.Slide, selection.GroupRotation);
            if (!solved.Succeeded)
            {
                return CommandResult.Failed(string.Join(" ", solved.Diagnostics));
            }

            var originals = new List<ShapeKey>(selection.Shapes.Count);
            foreach (var snapshot in selection.Shapes) originals.Add(snapshot.Key);

            var outcome = ShapeCreator.Create(_app, selection.SlideId, originals, solved.Copies);
            if (outcome.Created.Count == 0)
            {
                return CommandResult.Failed("None of the selected shapes could be duplicated.");
            }

            Diagnostics.Log(string.Format(
                CultureInfo.InvariantCulture,
                "Ran '{0}': created={1} missing={2} slideId={3}",
                label, outcome.Created.Count, outcome.Missing, selection.SlideId));

            var notes = new List<string>(solved.Diagnostics);
            var expected = originals.Count * solved.Copies.Count;
            if (outcome.Created.Count < expected)
            {
                notes.Add(string.Format(
                    CultureInfo.CurrentCulture, "{0} of {1} copies could not be made.",
                    expected - outcome.Created.Count, expected));
            }

            var message = notes.Count > 0 ? string.Join(" ", notes) : null;
            return solved.Notable || outcome.Created.Count < expected
                ? CommandResult.Note(message ?? "Part of the selection was skipped.")
                : CommandResult.Ok(message);
        }

        /// <summary>
        /// Lines up what is nearly lined up and evens out what is nearly even, deciding for itself
        /// what to move. With nothing selected it works on every visible shape on the slide. Always
        /// reports, since a 2pt fix is invisible - and says when it looked at the whole slide, so a
        /// click with nothing selected by accident is obvious, and one Ctrl+Z away.
        /// </summary>
        public CommandResult RunTidy(string label)
        {
            var wholeSlide = SelectionReader.NothingSelected(_app);
            var selection = wholeSlide
                ? SelectionReader.TryReadSlide(_app, Margin, out var problem)
                : SelectionReader.TryRead(_app, Margin, out problem);
            if (selection == null) return CommandResult.Failed(problem ?? "Nothing to tidy.");

            if (wholeSlide && selection.Shapes.Count(s => !s.IsConnector) < 2)
            {
                return CommandResult.Note("Nothing was selected, and this slide has fewer than two shapes to tidy.");
            }

            var request = new TidyRequest(TidyTolerance, Bounds);
            var solved = GroupSpace.SolveTidy(request, selection.Shapes, selection.Slide.GroupBounds, selection.GroupRotation);

            if (wholeSlide && solved.Succeeded)
            {
                var said = new List<string> { "Nothing was selected, so Tidy looked at the whole slide." };
                said.AddRange(solved.Diagnostics);
                solved = SolveResult.OkWithNotice(solved.Changes, said.ToArray());
            }

            // Nothing moved is an answer, not a failure, and Tidy's own words say why.
            if (solved.Succeeded && !solved.EffectiveChanges.Any())
            {
                return CommandResult.Note(string.Join(" ", solved.Diagnostics));
            }

            return ApplySolved(label, selection, solved);
        }

        /// <summary>
        /// Places the selection along the curve the anchor defines. The solve is its own entry point,
        /// but the result is ordinary geometry, so it applies like any align.
        /// </summary>
        public CommandResult RunCurve(string label)
        {
            var selection = SelectionReader.TryRead(_app, Margin, out var problem, includeAnchorCurve: true);
            if (selection == null) return CommandResult.Failed(problem ?? "Nothing to place.");

            if (selection.Anchor == null || selection.AnchorCurve == null)
            {
                return CommandResult.Failed(selection.AnchorCurveProblem ?? "Select the curve last.");
            }

            var request = new CurveRequest(selection.Anchor.Value, selection.AnchorCurve, ExactSpacing, RotateShapes);
            var solved = CurveSolver.Solve(request, selection.Shapes);

            return ApplySolved(label, selection, solved);
        }

        /// <summary>
        /// The margin as the solver wants it. The solver has always grown shapes on a negative
        /// margin, so the direction dropdown only has to supply the sign.
        /// </summary>
        private double SignedSizeMargin =>
            SizeDirection == SizeDirection.Grow ? -SizeMargin : SizeMargin;

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
