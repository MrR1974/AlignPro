using System;
using System.Collections.Generic;
using System.Linq;

namespace AlignPro.Geometry
{
    /// <summary>
    /// One undoable operation: everything a single ribbon click or hotkey changed, under a label the
    /// UI can show.
    /// </summary>
    public sealed class AlignTransaction
    {
        public AlignTransaction(string label, IReadOnlyList<GeometryChange> changes)
        {
            if (string.IsNullOrWhiteSpace(label)) throw new ArgumentException("A transaction needs a label.", nameof(label));
            if (changes is null) throw new ArgumentNullException(nameof(changes));

            Label = label;
            Changes = changes;
        }

        /// <summary>Shown in the UI, e.g. "Align left" or "Distribute horizontally".</summary>
        public string Label { get; }

        public IReadOnlyList<GeometryChange> Changes { get; }

        /// <summary>Changes that would actually move something.</summary>
        public IReadOnlyList<GeometryChange> EffectiveChanges =>
            Changes.Where(c => !c.IsNoOp).ToList();

        /// <summary>True when nothing in the transaction would move, so it is not worth recording.</summary>
        public bool IsEmpty => EffectiveChanges.Count == 0;

        /// <summary>The transaction running backwards, ready to apply.</summary>
        public AlignTransaction Inverted() =>
            new AlignTransaction(Label, Changes.Select(c => c.Inverted()).ToList());

        /// <summary>Builds a transaction from a solve, dropping changes that would not move anything.</summary>
        public static AlignTransaction FromResult(string label, SolveResult result)
        {
            if (result is null) throw new ArgumentNullException(nameof(result));
            return new AlignTransaction(label, result.EffectiveChanges.ToList());
        }
    }

    /// <summary>
    /// AlignPro's own undo stack.
    /// </summary>
    /// <remarks>
    /// <para>
    /// PowerPoint cannot do this for us. It has no <c>UndoRecord</c> equivalent,
    /// <c>CommandBars.ExecuteMso</c> fails outright, and shape geometry changes made through the
    /// object model do not appear to reach its undo stack at all - one measured Ctrl+Z after a batch
    /// of object-model work removed an entire slide rather than undoing the last shape move. See
    /// <c>docs/object-model-findings.md</c>.
    /// </para>
    /// <para>
    /// That is also why the add-in intercepts Ctrl+Z: without it, the keystroke reaches past our
    /// changes into the user's own earlier edits. Interception is conditional on
    /// <see cref="CanUndo"/>, so when this stack is empty the keystroke passes through to PowerPoint
    /// and normal editing is untouched.
    /// </para>
    /// <para>
    /// Pure and Office-free: the journal decides <em>what</em> to change and hands back a transaction,
    /// but never touches PowerPoint. Applying is the add-in's job, which is also where a shape that
    /// has since been deleted gets skipped. Not thread-safe by design - it is only ever touched from
    /// the UI thread, and the keyboard hook posts rather than calling in.
    /// </para>
    /// </remarks>
    public sealed class UndoManager
    {
        /// <summary>How many operations are remembered before the oldest is dropped.</summary>
        public const int DefaultCapacity = 20;

        // Front of each list is the most recent entry.
        private readonly LinkedList<AlignTransaction> _undo = new LinkedList<AlignTransaction>();
        private readonly LinkedList<AlignTransaction> _redo = new LinkedList<AlignTransaction>();

        public UndoManager(int capacity = DefaultCapacity)
        {
            if (capacity < 1) throw new ArgumentOutOfRangeException(nameof(capacity), capacity, "Capacity must be at least one.");
            Capacity = capacity;
        }

        public int Capacity { get; }

        public bool CanUndo => _undo.Count > 0;

        public bool CanRedo => _redo.Count > 0;

        public int UndoDepth => _undo.Count;

        public int RedoDepth => _redo.Count;

        /// <summary>The label of the operation Ctrl+Z would reverse, or null when there is none.</summary>
        public string? NextUndoLabel => _undo.First?.Value.Label;

        /// <summary>The label of the operation Ctrl+Y would replay, or null when there is none.</summary>
        public string? NextRedoLabel => _redo.First?.Value.Label;

        /// <summary>
        /// Records an operation that has just been applied. Transactions that would not move anything
        /// are ignored, so Ctrl+Z never consumes a step that does nothing visible.
        /// </summary>
        /// <returns>True when the transaction was recorded.</returns>
        public bool Push(AlignTransaction transaction)
        {
            if (transaction is null) throw new ArgumentNullException(nameof(transaction));
            if (transaction.IsEmpty) return false;

            _undo.AddFirst(transaction);
            while (_undo.Count > Capacity) _undo.RemoveLast();

            // A new operation invalidates anything that was undone before it.
            _redo.Clear();
            return true;
        }

        /// <summary>
        /// Takes the most recent operation off the undo stack and hands back the inverse, ready to
        /// apply. The caller applies it, then it becomes redoable.
        /// </summary>
        public bool TryUndo(out AlignTransaction inverse)
        {
            if (_undo.First is null)
            {
                inverse = null!;
                return false;
            }

            var transaction = _undo.First.Value;
            _undo.RemoveFirst();
            _redo.AddFirst(transaction);

            inverse = transaction.Inverted();
            return true;
        }

        /// <summary>
        /// Takes the most recently undone operation and hands it back in its original direction.
        /// </summary>
        public bool TryRedo(out AlignTransaction transaction)
        {
            if (_redo.First is null)
            {
                transaction = null!;
                return false;
            }

            transaction = _redo.First.Value;
            _redo.RemoveFirst();
            _undo.AddFirst(transaction);
            return true;
        }

        /// <summary>
        /// Forgets everything. Call when the active presentation changes: shape keys are only
        /// meaningful within one deck, so a stale stack would move the wrong shapes.
        /// </summary>
        public void Clear()
        {
            _undo.Clear();
            _redo.Clear();
        }
    }
}
