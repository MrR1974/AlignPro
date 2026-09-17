using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AlignPro.Geometry;
using Office = Microsoft.Office.Core;

namespace AlignPro.AddIn
{
    /// <summary>
    /// The AlignPro ribbon tab. Every callback is wrapped, because an exception escaping into
    /// PowerPoint's ribbon host is reported to the user as a broken add-in and can disable it.
    /// </summary>
    [ComVisible(true)]
    public sealed class AlignProRibbon : Office.IRibbonExtensibility
    {
        private readonly Func<AlignProController> _resolveController;
        private Office.IRibbonUI? _ribbon;

        internal AlignProRibbon(Func<AlignProController> resolveController)
        {
            _resolveController = resolveController;
        }

        private AlignProController Controller => _resolveController();

        public string GetCustomUI(string ribbonId)
        {
            var resource = typeof(AlignProRibbon).Namespace + ".Ribbon.xml";
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resource))
            {
                if (stream == null) throw new InvalidOperationException($"Embedded resource '{resource}' is missing.");
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbon) => _ribbon = ribbon;

        // -- verbs -------------------------------------------------------------------------------

        public void OnAlign(Office.IRibbonControl control) => Guard(() =>
        {
            var (verb, label) = control.Id switch
            {
                "btnAlignLeft" => (AlignVerb.AlignLeft, "Align left"),
                "btnAlignCentreH" => (AlignVerb.AlignCentreH, "Align centre"),
                "btnAlignRight" => (AlignVerb.AlignRight, "Align right"),
                "btnAlignTop" => (AlignVerb.AlignTop, "Align top"),
                "btnAlignCentreV" => (AlignVerb.AlignCentreV, "Align middle"),
                "btnAlignBottom" => (AlignVerb.AlignBottom, "Align bottom"),
                _ => throw new ArgumentOutOfRangeException(nameof(control), control.Id, "Unknown align button.")
            };

            Report(Controller.Run(verb, label), label);
        });

        public void OnDistribute(Office.IRibbonControl control) => Guard(() =>
        {
            var horizontal = control.Id == "btnDistributeH";
            var label = horizontal ? "Distribute horizontally" : "Distribute vertically";
            Report(Controller.Run(horizontal ? AlignVerb.DistributeH : AlignVerb.DistributeV, label), label);
        });

        public void OnMatchSize(Office.IRibbonControl control) => Guard(() =>
        {
            var (verb, label) = control.Id switch
            {
                "btnMatchWidth" => (AlignVerb.MatchWidth, "Match width"),
                "btnMatchHeight" => (AlignVerb.MatchHeight, "Match height"),
                "btnMatchBoth" => (AlignVerb.MatchBoth, "Match size"),
                _ => throw new ArgumentOutOfRangeException(nameof(control), control.Id, "Unknown match button.")
            };

            Report(Controller.Run(verb, label), label);
        });

        public void OnGrid(Office.IRibbonControl control) => Guard(() =>
            Report(Controller.Run(AlignVerb.GridArrange, "Arrange in grid"), "Arrange in grid"));

        // -- undo --------------------------------------------------------------------------------
        // AlignPro keeps its own undo because PowerPoint's cannot be trusted after an object-model
        // change: its entry may cover an unbounded amount of earlier work. Use these buttons rather
        // than Ctrl+Z after an AlignPro command. See UndoBoundary for the attempt to fix the native
        // behaviour and why it is not currently usable.

        public void OnUndo(Office.IRibbonControl control) => Guard(() =>
            Report(Controller.UndoLast(), "Undo"));

        public void OnRedo(Office.IRibbonControl control) => Guard(() =>
            Report(Controller.RedoLast(), "Redo"));

        public bool GetUndoEnabled(Office.IRibbonControl control) => Controller.Undo.CanUndo;

        public bool GetRedoEnabled(Office.IRibbonControl control) => Controller.Undo.CanRedo;

        public string GetUndoLabel(Office.IRibbonControl control) =>
            Controller.Undo.NextUndoLabel is string label ? "Undo " + label.ToLowerInvariant() : "Undo";

        public string GetRedoLabel(Office.IRibbonControl control) =>
            Controller.Undo.NextRedoLabel is string label ? "Redo " + label.ToLowerInvariant() : "Redo";

        // -- reference and bounds ----------------------------------------------------------------

        private static readonly ReferenceTarget[] References =
        {
            ReferenceTarget.Anchor,
            ReferenceTarget.SelectionBounds,
            ReferenceTarget.Slide,
            ReferenceTarget.SlideMargins,
            ReferenceTarget.PlaceholderBounds
        };

        private static readonly BoundsModel[] BoundsModels =
        {
            BoundsModel.ShapeFrame,
            BoundsModel.VisualBounds,
            BoundsModel.TextBounds
        };

        private static readonly DistributeMode[] DistributeModes =
        {
            DistributeMode.LeadingEdge,
            DistributeMode.Centre,
            DistributeMode.TrailingEdge,
            DistributeMode.Gap
        };

        public int GetReferenceIndex(Office.IRibbonControl control) =>
            Math.Max(0, Array.IndexOf(References, Controller.Reference));

        public void OnReferenceChange(Office.IRibbonControl control, string selectedId, int selectedIndex) =>
            Guard(() => Controller.Reference = References[selectedIndex]);

        public int GetBoundsIndex(Office.IRibbonControl control) =>
            Math.Max(0, Array.IndexOf(BoundsModels, Controller.Bounds));

        public void OnBoundsChange(Office.IRibbonControl control, string selectedId, int selectedIndex) =>
            Guard(() => Controller.Bounds = BoundsModels[selectedIndex]);

        public int GetDistributeModeIndex(Office.IRibbonControl control) =>
            Math.Max(0, Array.IndexOf(DistributeModes, Controller.DistributeMode));

        public void OnDistributeModeChange(Office.IRibbonControl control, string selectedId, int selectedIndex) =>
            Guard(() => Controller.DistributeMode = DistributeModes[selectedIndex]);

        // -- numeric boxes -----------------------------------------------------------------------

        public string GetSpacingText(Office.IRibbonControl control) =>
            Controller.ExactSpacing?.ToString("0.##", CultureInfo.CurrentCulture) ?? string.Empty;

        public void OnSpacingChange(Office.IRibbonControl control, string text) => Guard(() =>
        {
            // Blank means "even out what is already there", which is a meaningful setting rather than
            // a validation failure.
            if (string.IsNullOrWhiteSpace(text))
            {
                Controller.ExactSpacing = null;
                return;
            }

            if (TryParsePoints(text, out var value))
            {
                Controller.ExactSpacing = value;
            }
            else
            {
                Warn($"'{text}' is not a number of points. Leave the box blank to even out the existing spacing.");
            }

            Invalidate();
        });

        public string GetMarginText(Office.IRibbonControl control) =>
            Controller.Margin.ToString("0.##", CultureInfo.CurrentCulture);

        public void OnMarginChange(Office.IRibbonControl control, string text) => Guard(() =>
        {
            if (TryParsePoints(text, out var value) && value >= 0)
            {
                Controller.Margin = value;
            }
            else
            {
                Warn($"'{text}' is not a margin in points. It must be zero or more.");
            }

            Invalidate();
        });

        public string GetGridColumnsText(Office.IRibbonControl control) =>
            Controller.GridColumns?.ToString(CultureInfo.CurrentCulture) ?? string.Empty;

        public void OnGridColumnsChange(Office.IRibbonControl control, string text) => Guard(() =>
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                Controller.GridColumns = null;
                return;
            }

            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.CurrentCulture, out var columns) && columns >= 1)
            {
                Controller.GridColumns = columns;
            }
            else
            {
                Warn($"'{text}' is not a column count. Leave the box blank for a near-square grid.");
            }

            Invalidate();
        });

        public bool GetResizeFromCentre(Office.IRibbonControl control) =>
            Controller.ResizeOrigin == ResizeOrigin.Centre;

        public void OnResizeFromCentreChange(Office.IRibbonControl control, bool pressed) =>
            Guard(() => Controller.ResizeOrigin = pressed ? ResizeOrigin.Centre : ResizeOrigin.TopLeft);

        // -- plumbing ----------------------------------------------------------------------------

        private static bool TryParsePoints(string text, out double value) =>
            double.TryParse(
                text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out value) ||
            double.TryParse(
                text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value);

        /// <summary>
        /// Refresh the ribbon's dynamic state - undo labels, enabled states, the text of the edit boxes
        /// after a value was rejected.
        /// </summary>
        private void Invalidate() => _ribbon?.Invalidate();

        /// <summary>
        /// Only speaks up when something did not happen. A successful align says nothing: a dialog after
        /// every click would be intolerable in normal use.
        /// </summary>
        private void Report(CommandResult result, string title)
        {
            Invalidate();

            if (result.Succeeded && result.Message == null) return;

            if (result.Succeeded)
            {
                // Worth mentioning but not worth interrupting for.
                return;
            }

            MessageBox.Show(
                result.Message, "AlignPro - " + title,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void Warn(string message) =>
            MessageBox.Show(message, "AlignPro", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        private static void Guard(Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                // Never let an exception cross back into the ribbon host.
                MessageBox.Show(
                    ex.Message, "AlignPro encountered a problem",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
