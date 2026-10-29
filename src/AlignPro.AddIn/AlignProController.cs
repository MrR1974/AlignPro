using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
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

    /// <summary>Which number of the duplicate step a box sets.</summary>
    internal enum DuplicateStep { X, Y, Angle }

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
    /// settings the ribbon's dropdowns bind to, and remembers them between sessions.
    /// </summary>
    internal sealed class AlignProController
    {
        private readonly PowerPoint.Application _app;

        /// <summary>Null while the settings are being loaded or reset, so that does not write them back.</summary>
        private SettingsStore? _store;

        public AlignProController(PowerPoint.Application app) : this(app, new SettingsStore())
        {
        }

        /// <param name="store">Where settings are remembered, or null for a controller that only holds the defaults.</param>
        private AlignProController(PowerPoint.Application app, SettingsStore? store)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            if (store == null) return;

            Load(store.ReadAll());
            _store = store;
        }

        // -- settings -------------------------------------------------------------------------------
        // Each property's name is also the name it is saved under, so renaming one forgets what users
        // had chosen for it.

        private ReferenceTarget _reference = ReferenceTarget.SelectionBounds;
        public ReferenceTarget Reference { get => _reference; set => Remember(ref _reference, value); }

        private BoundsModel _bounds = BoundsModel.ShapeFrame;
        public BoundsModel Bounds { get => _bounds; set => Remember(ref _bounds, value); }

        private DistributeMode _distributeMode = DistributeMode.Gap;
        public DistributeMode DistributeMode { get => _distributeMode; set => Remember(ref _distributeMode, value); }

        private double? _exactSpacing;

        /// <summary>Null means "even out whatever space is already there".</summary>
        public double? ExactSpacing { get => _exactSpacing; set => Remember(ref _exactSpacing, value); }

        private double _margin = 36;

        /// <summary>Inset for the slide-margins reference, in points.</summary>
        public double Margin { get => _margin; set => Remember(ref _margin, value); }

        private ResizeOrigin _resizeOrigin = ResizeOrigin.TopLeft;
        public ResizeOrigin ResizeOrigin { get => _resizeOrigin; set => Remember(ref _resizeOrigin, value); }

        private double _sizeMargin;

        /// <summary>
        /// The match-size margin, measured per side in points. Always zero or more: which way it
        /// goes is <see cref="SizeDirection"/>'s job, not the sign's.
        /// </summary>
        public double SizeMargin { get => _sizeMargin; set => Remember(ref _sizeMargin, value); }

        private SizeDirection _sizeDirection = SizeDirection.Shrink;

        /// <summary>Whether the margin shrinks the shapes inside the anchor or grows them past it.</summary>
        public SizeDirection SizeDirection { get => _sizeDirection; set => Remember(ref _sizeDirection, value); }

        private SizeMarginMode _sizeMarginMode = SizeMarginMode.None;

        /// <summary>Whether the match-size margin applies, and whether it cascades.</summary>
        public SizeMarginMode SizeMarginMode { get => _sizeMarginMode; set => Remember(ref _sizeMarginMode, value); }

        private int? _gridColumns;

        /// <summary>Null means a near-square grid.</summary>
        public int? GridColumns { get => _gridColumns; set => Remember(ref _gridColumns, value); }

        private double _duplicateX = 20;

        /// <summary>How far each duplicate step moves right, in points.</summary>
        public double DuplicateX { get => _duplicateX; set => Remember(ref _duplicateX, value); }

        private double _duplicateY = 20;

        /// <summary>How far each duplicate step moves down, in points.</summary>
        public double DuplicateY { get => _duplicateY; set => Remember(ref _duplicateY, value); }

        private double _duplicateAngle;

        /// <summary>How far each duplicate step turns, in degrees clockwise.</summary>
        public double DuplicateAngle { get => _duplicateAngle; set => Remember(ref _duplicateAngle, value); }

        private int _duplicateCopies = 1;
        public int DuplicateCopies { get => _duplicateCopies; set => Remember(ref _duplicateCopies, value); }

        private DuplicatePivot _duplicatePivot = DuplicatePivot.OwnCentre;
        public DuplicatePivot DuplicatePivot { get => _duplicatePivot; set => Remember(ref _duplicatePivot, value); }

        private bool _rotateShapes = true;

        /// <summary>
        /// Shared by Duplicate and Distribute along curve: turn shapes with the step or the curve, or
        /// leave their angle alone.
        /// </summary>
        public bool RotateShapes { get => _rotateShapes; set => Remember(ref _rotateShapes, value); }

        private double _tidyTolerance = TidyRequest.DefaultTolerance;

        /// <summary>How close, in points, shapes must be for Tidy to treat them as meant to line up.</summary>
        public double TidyTolerance { get => _tidyTolerance; set => Remember(ref _tidyTolerance, value); }

        /// <summary>Puts every setting back to its default and forgets what was saved.</summary>
        public void ResetSettings()
        {
            var store = _store;
            _store = null;
            try
            {
                CopySettingsFrom(new AlignProController(_app, store: null));
                store?.Clear();
            }
            finally
            {
                _store = store;
            }
        }

        private void CopySettingsFrom(AlignProController other)
        {
            Reference = other.Reference;
            Bounds = other.Bounds;
            DistributeMode = other.DistributeMode;
            ExactSpacing = other.ExactSpacing;
            Margin = other.Margin;
            ResizeOrigin = other.ResizeOrigin;
            SizeMargin = other.SizeMargin;
            SizeDirection = other.SizeDirection;
            SizeMarginMode = other.SizeMarginMode;
            GridColumns = other.GridColumns;
            DuplicateX = other.DuplicateX;
            DuplicateY = other.DuplicateY;
            DuplicateAngle = other.DuplicateAngle;
            DuplicateCopies = other.DuplicateCopies;
            DuplicatePivot = other.DuplicatePivot;
            RotateShapes = other.RotateShapes;
            TidyTolerance = other.TidyTolerance;
        }

        private void Remember<T>(ref T field, T value, [CallerMemberName] string name = "")
        {
            field = value;
            _store?.Write(name, Stored(value));
        }

        /// <summary>
        /// A setting as it is saved: invariant culture, so it reads back the same after a change of
        /// locale, and choices by name, so reordering an enum cannot remap them. Blank is "none".
        /// </summary>
        private static string Stored(object? value) => value switch
        {
            null => string.Empty,
            double number => number.ToString("R", CultureInfo.InvariantCulture),
            int whole => whole.ToString(CultureInfo.InvariantCulture),
            _ => value.ToString()
        };

        /// <summary>
        /// Takes what was saved through the same checks a typed value gets, so a value saved by an
        /// older version, or edited by hand, that is no longer acceptable leaves the default in place.
        /// Anything missing or unreadable is the default too.
        /// </summary>
        private void Load(IReadOnlyDictionary<string, string> saved)
        {
            LoadChoice<ReferenceTarget>(saved, nameof(Reference), v => Reference = v);
            LoadChoice<BoundsModel>(saved, nameof(Bounds), v => Bounds = v);
            LoadChoice<DistributeMode>(saved, nameof(DistributeMode), v => DistributeMode = v);
            LoadChoice<ResizeOrigin>(saved, nameof(ResizeOrigin), v => ResizeOrigin = v);
            LoadChoice<SizeDirection>(saved, nameof(SizeDirection), v => SizeDirection = v);
            LoadChoice<SizeMarginMode>(saved, nameof(SizeMarginMode), v => SizeMarginMode = v);
            LoadChoice<DuplicatePivot>(saved, nameof(DuplicatePivot), v => DuplicatePivot = v);

            LoadNumber(saved, nameof(ExactSpacing), v => Taken(ExactSpacing = v), blankAllowed: true);
            LoadNumber(saved, nameof(Margin), v => SetMargin(v!.Value));
            LoadNumber(saved, nameof(SizeMargin), v => SetSizeMargin(v!.Value));
            LoadNumber(saved, nameof(TidyTolerance), v => SetTidyTolerance(v!.Value));
            LoadNumber(saved, nameof(DuplicateX), v => Taken(DuplicateX = v!.Value));
            LoadNumber(saved, nameof(DuplicateY), v => Taken(DuplicateY = v!.Value));
            LoadNumber(saved, nameof(DuplicateAngle), v => Taken(DuplicateAngle = v!.Value));

            LoadWhole(saved, nameof(GridColumns), v => SetGridColumns(v), blankAllowed: true);
            LoadWhole(saved, nameof(DuplicateCopies), v => SetDuplicateCopies(v!.Value));

            if (saved.TryGetValue(nameof(RotateShapes), out var rotate))
            {
                if (bool.TryParse(rotate, out var parsed)) RotateShapes = parsed;
                else Diagnostics.Log("Ignored saved RotateShapes '" + rotate + "'.");
            }
        }

        /// <summary>For a setting any number suits, so there is nothing to refuse.</summary>
        private static string? Taken<T>(T _) => null;

        private static void LoadChoice<T>(IReadOnlyDictionary<string, string> saved, string name, Action<T> apply)
            where T : struct, Enum
        {
            if (!saved.TryGetValue(name, out var text)) return;

            // IsDefined as well as TryParse, which would happily take "7" for an enum with three values.
            if (Enum.TryParse(text, ignoreCase: false, out T value) && Enum.IsDefined(typeof(T), value)) apply(value);
            else Diagnostics.Log("Ignored saved " + name + " '" + text + "'.");
        }

        private static void LoadNumber(
            IReadOnlyDictionary<string, string> saved, string name, Func<double?, string?> apply, bool blankAllowed = false)
        {
            if (!saved.TryGetValue(name, out var text)) return;

            double? value = null;
            if (text.Length > 0)
            {
                // Invariant only. TryParsePoints tries the user's locale first, which on a German machine
                // reads the saved "1.5" as fifteen.
                if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) &&
                    !double.IsNaN(parsed) && !double.IsInfinity(parsed))
                {
                    value = parsed;
                }
                else
                {
                    Diagnostics.Log("Ignored saved " + name + " '" + text + "'.");
                    return;
                }
            }
            else if (!blankAllowed)
            {
                Diagnostics.Log("Ignored saved " + name + ", which was blank.");
                return;
            }

            var refused = apply(value);
            if (refused != null) Diagnostics.Log("Ignored saved " + name + " '" + text + "': " + refused);
        }

        private static void LoadWhole(
            IReadOnlyDictionary<string, string> saved, string name, Func<int?, string?> apply, bool blankAllowed = false)
        {
            if (!saved.TryGetValue(name, out var text)) return;

            int? value = null;
            if (text.Length > 0)
            {
                if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
                {
                    value = parsed;
                }
                else
                {
                    Diagnostics.Log("Ignored saved " + name + " '" + text + "'.");
                    return;
                }
            }
            else if (!blankAllowed)
            {
                Diagnostics.Log("Ignored saved " + name + ", which was blank.");
                return;
            }

            var refused = apply(value);
            if (refused != null) Diagnostics.Log("Ignored saved " + name + " '" + text + "': " + refused);
        }

        // -- settings typed as text -----------------------------------------------------------------
        // The ribbon's boxes and the automation surface both set these from text, through here, so the
        // two cannot accept different things. Each returns null when the value was taken, or what was
        // wrong with it, and leaves the setting alone.

        /// <summary>Exact spacing in points; blank means "even out what is already there".</summary>
        public string? SetExactSpacing(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                ExactSpacing = null;
                return null;
            }

            if (!TryParsePoints(text, out var value))
            {
                return $"'{text}' is not a number of points. Leave it blank to even out the existing spacing.";
            }

            ExactSpacing = value;
            return null;
        }

        /// <summary>
        /// The match-size margin in points, per side; blank is zero. Never negative: the sign belongs
        /// to <see cref="SizeDirection"/>, and two settings that can cancel each other out - Grow with
        /// -10, shrinking - would be a puzzle.
        /// </summary>
        public string? SetSizeMargin(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                SizeMargin = 0;
                return null;
            }

            if (!TryParsePoints(text, out var value)) return $"'{text}' is not a margin in points.";
            return SetSizeMargin(value);
        }

        /// <inheritdoc cref="SetSizeMargin(string)"/>
        public string? SetSizeMargin(double value)
        {
            if (value < 0)
            {
                return string.Format(CultureInfo.CurrentCulture,
                    "The margin is always a positive number. To make the shapes larger than the anchor, set Direction to Grow and enter {0:0.##}.",
                    -value);
            }

            SizeMargin = value;
            return null;
        }

        /// <summary>The slide-margins inset, in points: zero or more.</summary>
        public string? SetMargin(string text)
        {
            if (!TryParsePoints(text, out var value)) return MarginRefusal(text);
            return SetMargin(value, text);
        }

        /// <inheritdoc cref="SetMargin(string)"/>
        public string? SetMargin(double value) => SetMargin(value, value.ToString("0.##", CultureInfo.CurrentCulture));

        private string? SetMargin(double value, string asTyped)
        {
            if (value < 0) return MarginRefusal(asTyped);

            Margin = value;
            return null;
        }

        private static string MarginRefusal(string asTyped) =>
            $"'{asTyped}' is not a margin in points. It must be zero or more.";

        /// <summary>Tidy's tolerance in points, within the range Tidy accepts; blank restores the default.</summary>
        public string? SetTidyTolerance(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                TidyTolerance = TidyRequest.DefaultTolerance;
                return null;
            }

            if (!TryParsePoints(text, out var value)) return ToleranceRefusal(text);
            return SetTidyTolerance(value, text);
        }

        /// <inheritdoc cref="SetTidyTolerance(string)"/>
        public string? SetTidyTolerance(double value) => SetTidyTolerance(value, value.ToString("0.##", CultureInfo.CurrentCulture));

        private string? SetTidyTolerance(double value, string asTyped)
        {
            if (value < TidyRequest.MinTolerance || value > TidyRequest.MaxTolerance) return ToleranceRefusal(asTyped);

            TidyTolerance = value;
            return null;
        }

        private static string ToleranceRefusal(string asTyped) => string.Format(CultureInfo.CurrentCulture,
            "'{0}' is not a tolerance Tidy can use. Enter a number of points from {1} to {2}.",
            asTyped, TidyRequest.MinTolerance, TidyRequest.MaxTolerance);

        /// <summary>Grid columns, one or more; blank means a near-square grid.</summary>
        public string? SetGridColumns(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                GridColumns = null;
                return null;
            }

            if (!TryParseWhole(text, out var columns)) return ColumnsRefusal(text);
            return SetGridColumns(columns, text);
        }

        /// <inheritdoc cref="SetGridColumns(string)"/>
        public string? SetGridColumns(int? columns) => columns == null
            ? SetGridColumns(string.Empty)
            : SetGridColumns(columns.Value, columns.Value.ToString(CultureInfo.CurrentCulture));

        private string? SetGridColumns(int columns, string asTyped)
        {
            if (columns < 1) return ColumnsRefusal(asTyped);

            GridColumns = columns;
            return null;
        }

        private static string ColumnsRefusal(string asTyped) =>
            $"'{asTyped}' is not a column count. Leave it blank for a near-square grid.";

        /// <summary>One number of the duplicate step - X, Y or Angle. Blank is zero, which is inert in a step.</summary>
        public string? SetDuplicateStep(DuplicateStep part, string text)
        {
            var value = 0.0;
            if (!string.IsNullOrWhiteSpace(text) && !TryParsePoints(text, out value))
            {
                return $"'{text}' is not a number.";
            }

            switch (part)
            {
                case DuplicateStep.X: DuplicateX = value; break;
                case DuplicateStep.Y: DuplicateY = value; break;
                default: DuplicateAngle = value; break;
            }

            return null;
        }

        /// <summary>How many copies Duplicate makes: a whole number from one to the limit.</summary>
        public string? SetDuplicateCopies(string text)
        {
            if (!TryParseWhole(text, out var copies)) return CopiesRefusal(text);
            return SetDuplicateCopies(copies, text);
        }

        /// <inheritdoc cref="SetDuplicateCopies(string)"/>
        public string? SetDuplicateCopies(int copies) => SetDuplicateCopies(copies, copies.ToString(CultureInfo.CurrentCulture));

        private string? SetDuplicateCopies(int copies, string asTyped)
        {
            if (copies < 1 || copies > DuplicateRequest.MaxCopies) return CopiesRefusal(asTyped);

            DuplicateCopies = copies;
            return null;
        }

        private static string CopiesRefusal(string asTyped) =>
            $"'{asTyped}' is not a number of copies. Enter a whole number from 1 to {DuplicateRequest.MaxCopies}.";

        /// <summary>
        /// A number as the user's locale writes it, or failing that as script does - so "1,5" works on
        /// a French machine and "1.5" works everywhere.
        /// </summary>
        private static bool TryParsePoints(string text, out double value) =>
            double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out value) ||
            double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value);

        private static bool TryParseWhole(string text, out int value) =>
            int.TryParse(text, NumberStyles.Integer, CultureInfo.CurrentCulture, out value) ||
            int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);

        public CommandResult Run(AlignVerb verb, string label)
        {
            var timer = Diagnostics.StartStages();
            // Match size and rotation are only meaningful against an anchor, so supply one rather than
            // refusing on a technicality - decided before reading, which needs to know.
            var reference = IsMatchSize(verb) || verb == AlignVerb.MatchRotation ? ReferenceTarget.Anchor : Reference;
            var selection = SelectionReader.TryRead(
                _app, Margin, out var problem,
                textBounds: Bounds == BoundsModel.TextBounds,
                placeholderBounds: reference == ReferenceTarget.PlaceholderBounds);
            timer.Mark("read");
            if (selection == null) return CommandResult.Failed(problem ?? "Nothing to align.");

            using (selection)
            {
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
                timer.Mark("solve");
                return ApplySolved(label, selection, solved, timer);
            }
        }

        /// <summary>
        /// Applies a geometry solve. Shared by every verb whose result is a set of frame and angle
        /// changes, whichever solver produced it.
        /// </summary>
        private CommandResult ApplySolved(string label, SelectionSnapshot selection, SolveResult solved, StageTimer timer)
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

            var outcome = ChangeApplier.Apply(_app, selection.Live, changes);
            timer.Mark("apply");
            if (outcome.Applied == 0)
            {
                return CommandResult.Failed("None of the selected shapes could be moved.");
            }

            Diagnostics.Log(string.Format(
                CultureInfo.InvariantCulture,
                "Ran '{0}': applied={1} missing={2} slideId={3} {4}",
                label, outcome.Applied, outcome.Missing, selection.SlideId, timer));

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
            var timer = Diagnostics.StartStages();
            var selection = SelectionReader.TryRead(_app, Margin, out var problem, includeSlideOrder: true);
            timer.Mark("read");
            if (selection == null) return CommandResult.Failed(problem ?? "Nothing to reorder.");

            using (selection)
            {
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

                timer.Mark("solve");
                var outcome = ChangeApplier.ApplyOrder(_app, selection.Live, solved.Change.OldOrder, solved.Change.NewOrder);
                timer.Mark("apply");
                if (outcome.Applied == 0)
                {
                    return CommandResult.Failed("None of the selected shapes could be restacked.");
                }

                Diagnostics.Log(string.Format(
                    CultureInfo.InvariantCulture,
                    "Ran '{0}': restacked={1} missing={2} slideId={3} {4}",
                    label, outcome.Applied, outcome.Missing, selection.SlideId, timer));

                var skipped = Skipped(outcome);
                return skipped == null ? CommandResult.Ok() : CommandResult.Note(skipped);
            }
        }

        /// <summary>
        /// Duplicates the selection as a unit, <see cref="DuplicateCopies"/> times, each copy one step
        /// further on. A third entry point beside <see cref="Run"/> and <see cref="RunOrder"/>,
        /// because it makes shapes rather than editing them.
        /// </summary>
        public CommandResult RunDuplicate(string label)
        {
            var timer = Diagnostics.StartStages();
            var selection = SelectionReader.TryRead(_app, Margin, out var problem);
            timer.Mark("read");
            if (selection == null) return CommandResult.Failed(problem ?? "Nothing to duplicate.");

            using (selection)
            {
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

                timer.Mark("solve");
                var outcome = ShapeCreator.Create(_app, selection.SlideId, selection.Live, originals, solved.Copies);
                timer.Mark("apply");
                if (outcome.Created.Count == 0)
                {
                    return CommandResult.Failed("None of the selected shapes could be duplicated.");
                }

                Diagnostics.Log(string.Format(
                    CultureInfo.InvariantCulture,
                    "Ran '{0}': created={1} missing={2} slideId={3} {4}",
                    label, outcome.Created.Count, outcome.Missing, selection.SlideId, timer));

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
        }

        /// <summary>
        /// Lines up what is nearly lined up and evens out what is nearly even, deciding for itself
        /// what to move. With nothing selected it works on every visible shape on the slide. Always
        /// reports, since a 2pt fix is invisible - and says when it looked at the whole slide, so a
        /// click with nothing selected by accident is obvious, and one Ctrl+Z away.
        /// </summary>
        public CommandResult RunTidy(string label)
        {
            var timer = Diagnostics.StartStages();
            var wholeSlide = SelectionReader.NothingSelected(_app);
            var textBounds = Bounds == BoundsModel.TextBounds;
            var selection = wholeSlide
                ? SelectionReader.TryReadSlide(_app, Margin, out var problem, textBounds)
                : SelectionReader.TryRead(_app, Margin, out problem, textBounds: textBounds);
            timer.Mark("read");
            if (selection == null) return CommandResult.Failed(problem ?? "Nothing to tidy.");

            using (selection)
            {
                if (wholeSlide && selection.Shapes.Count(s => !s.IsConnector) < 2)
                {
                    return CommandResult.Note("Nothing was selected, and this slide has fewer than two shapes to tidy.");
                }

                var request = new TidyRequest(TidyTolerance, Bounds);
                var solved = GroupSpace.SolveTidy(request, selection.Shapes, selection.Slide.GroupBounds, selection.GroupRotation);
                timer.Mark("solve");

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

                return ApplySolved(label, selection, solved, timer);
            }
        }

        /// <summary>
        /// Places the selection along the curve the anchor defines. The solve is its own entry point,
        /// but the result is ordinary geometry, so it applies like any align.
        /// </summary>
        public CommandResult RunCurve(string label)
        {
            var timer = Diagnostics.StartStages();
            var selection = SelectionReader.TryRead(_app, Margin, out var problem, includeAnchorCurve: true);
            timer.Mark("read");
            if (selection == null) return CommandResult.Failed(problem ?? "Nothing to place.");

            using (selection)
            {
                if (selection.Anchor == null || selection.AnchorCurve == null)
                {
                    return CommandResult.Failed(selection.AnchorCurveProblem ?? "Select the curve last.");
                }

                var request = new CurveRequest(selection.Anchor.Value, selection.AnchorCurve, ExactSpacing, RotateShapes);
                var solved = CurveSolver.Solve(request, selection.Shapes);
                timer.Mark("solve");

                return ApplySolved(label, selection, solved, timer);
            }
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
