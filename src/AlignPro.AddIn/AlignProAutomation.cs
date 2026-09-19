using System;
using System.Globalization;
using System.Runtime.InteropServices;
using AlignPro.Geometry;

namespace AlignPro.AddIn
{
    /// <summary>
    /// Drives AlignPro from outside PowerPoint, reached via
    /// <c>Application.COMAddIns.Item("AlignPro.AddIn").Object</c>.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This exists so the add-in can be tested without a human clicking the ribbon. Ribbon callbacks
    /// are only reachable by clicking, which made every behavioural question - especially about undo -
    /// an expensive round trip. With this, a PowerShell harness can build a deck, select shapes, run a
    /// verb and inspect the result, end to end.
    /// </para>
    /// <para>
    /// It doubles as a macro surface: anything AlignPro can do from the ribbon can be scripted.
    /// </para>
    /// <para>
    /// Every method returns a status string rather than throwing, because COM exceptions across a
    /// late-bound boundary lose almost everything useful. An empty string means success.
    /// </para>
    /// </remarks>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class AlignProAutomation
    {
        private readonly Func<AlignProController> _resolveController;

        internal AlignProAutomation(Func<AlignProController> resolveController)
        {
            _resolveController = resolveController;
        }

        private AlignProController Controller => _resolveController();

        /// <summary>Where the trace log is written.</summary>
        public string LogPath => Diagnostics.LogPath;

        /// <summary>
        /// Turns per-shape and per-change tracing on or off. Noisy, but it shows exactly what was read
        /// from PowerPoint and what the solver decided, which is how the grid ordering bug was found.
        /// </summary>
        public string SetVerboseLogging(bool enabled)
        {
            Diagnostics.Verbose = enabled;
            return string.Empty;
        }

        /// <summary>The current settings, for a harness to assert against.</summary>
        public string Describe() => string.Format(
            CultureInfo.InvariantCulture,
            "reference={0} bounds={1} distributeMode={2} exactSpacing={3} margin={4} gridColumns={5} " +
            "sizeMargin={6} sizeMarginMode={7}",
            Controller.Reference,
            Controller.Bounds,
            Controller.DistributeMode,
            Controller.ExactSpacing?.ToString(CultureInfo.InvariantCulture) ?? "none",
            Controller.Margin.ToString(CultureInfo.InvariantCulture),
            Controller.GridColumns?.ToString(CultureInfo.InvariantCulture) ?? "auto",
            Controller.SizeMargin.ToString(CultureInfo.InvariantCulture),
            Controller.SizeMarginMode);

        /// <summary>
        /// Runs a verb against the current selection. Verb names match <see cref="AlignVerb"/>, e.g.
        /// "AlignLeft", "DistributeH", "MatchBoth", "GridArrange".
        /// </summary>
        /// <returns>Empty on success, otherwise why nothing happened.</returns>
        public string RunVerb(string verb)
        {
            if (!Enum.TryParse(verb, ignoreCase: true, result: out AlignVerb parsed))
            {
                return "Unknown verb '" + verb + "'.";
            }

            // Returns the message whether the operation was refused or merely skipped something, so a
            // harness sees the same information the ribbon shows. Empty means a clean, silent success.
            var result = Controller.Run(parsed, verb);
            return result.Message ?? string.Empty;
        }

        /// <summary>
        /// Reverses the last AlignPro operation, exactly as the ribbon's Undo button does.
        /// </summary>
        /// <remarks>
        /// AlignPro keeps its own undo stack, so a harness cannot reach it through
        /// <c>CommandBars.ExecuteMso("Undo")</c> - that drives PowerPoint's, which is a different
        /// thing entirely and the reason the stack exists. Ordering has no geometry for native undo to
        /// restore either, so without this there is no way to test that path from script at all.
        /// </remarks>
        /// <returns>Empty on success, otherwise why nothing happened.</returns>
        public string Undo() => Controller.UndoLast().Message ?? string.Empty;

        /// <summary>Reapplies the operation just undone, as the ribbon's Redo button does.</summary>
        /// <returns>Empty on success, otherwise why nothing happened.</returns>
        public string Redo() => Controller.RedoLast().Message ?? string.Empty;

        /// <summary>
        /// Restacks the current selection. Verb names match <see cref="OrderVerb"/>, i.e.
        /// "StackFirstOnTop" or "StackFirstOnBottom".
        /// </summary>
        /// <returns>Empty on success, otherwise why nothing happened.</returns>
        public string RunOrder(string verb)
        {
            if (!Enum.TryParse(verb, ignoreCase: true, result: out OrderVerb parsed))
            {
                return "Unknown order verb '" + verb + "'.";
            }

            var result = Controller.RunOrder(parsed, verb);
            return result.Message ?? string.Empty;
        }

        /// <summary>Sets what to align against. Names match <see cref="ReferenceTarget"/>.</summary>
        public string SetReference(string reference)
        {
            if (!Enum.TryParse(reference, ignoreCase: true, result: out ReferenceTarget parsed))
            {
                return "Unknown reference '" + reference + "'.";
            }

            Controller.Reference = parsed;
            return string.Empty;
        }

        /// <summary>Sets which rectangle to measure. Names match <see cref="BoundsModel"/>.</summary>
        public string SetBoundsModel(string model)
        {
            if (!Enum.TryParse(model, ignoreCase: true, result: out BoundsModel parsed))
            {
                return "Unknown bounds model '" + model + "'.";
            }

            Controller.Bounds = parsed;
            return string.Empty;
        }

        /// <summary>Sets the distribute reference point. Names match <see cref="DistributeMode"/>.</summary>
        public string SetDistributeMode(string mode)
        {
            if (!Enum.TryParse(mode, ignoreCase: true, result: out DistributeMode parsed))
            {
                return "Unknown distribute mode '" + mode + "'.";
            }

            Controller.DistributeMode = parsed;
            return string.Empty;
        }

        /// <summary>Sets exact spacing in points. An empty string clears it.</summary>
        public string SetExactSpacing(string points)
        {
            if (string.IsNullOrWhiteSpace(points))
            {
                Controller.ExactSpacing = null;
                return string.Empty;
            }

            if (!double.TryParse(points, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
            {
                return "'" + points + "' is not a number.";
            }

            Controller.ExactSpacing = value;
            return string.Empty;
        }

        /// <summary>
        /// Sets the match-size margin in points, measured per side. Negative grows the shapes. An
        /// empty string clears it.
        /// </summary>
        public string SetSizeMargin(string points)
        {
            if (string.IsNullOrWhiteSpace(points))
            {
                Controller.SizeMargin = 0;
                return string.Empty;
            }

            if (!double.TryParse(points, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
            {
                return "'" + points + "' is not a number.";
            }

            Controller.SizeMargin = value;
            return string.Empty;
        }

        /// <summary>Sets how the margin applies. Names match <see cref="SizeMarginMode"/>.</summary>
        public string SetSizeMarginMode(string mode)
        {
            if (!Enum.TryParse(mode, ignoreCase: true, result: out SizeMarginMode parsed))
            {
                return "Unknown size margin mode '" + mode + "'.";
            }

            Controller.SizeMarginMode = parsed;
            return string.Empty;
        }

        /// <summary>Sets the grid column count. An empty string means a near-square grid.</summary>
        public string SetGridColumns(string columns)
        {
            if (string.IsNullOrWhiteSpace(columns))
            {
                Controller.GridColumns = null;
                return string.Empty;
            }

            if (!int.TryParse(columns, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) || value < 1)
            {
                return "'" + columns + "' is not a column count.";
            }

            Controller.GridColumns = value;
            return string.Empty;
        }
    }
}
