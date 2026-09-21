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
        public AlignTransaction(
            string label, IReadOnlyList<GeometryChange> changes, ZOrderChange? order = null)
        {
            if (string.IsNullOrWhiteSpace(label)) throw new ArgumentException("A transaction needs a label.", nameof(label));
            if (changes is null) throw new ArgumentNullException(nameof(changes));

            Label = label;
            Changes = changes;
            Order = order;
        }

        /// <summary>Shown in the UI, e.g. "Align left" or "Distribute horizontally".</summary>
        public string Label { get; }

        public IReadOnlyList<GeometryChange> Changes { get; }

        /// <summary>
        /// The stacking order this operation set, or null for the geometry verbs - which is all of
        /// them bar ordering.
        /// </summary>
        /// <remarks>
        /// Carried alongside the geometry rather than folded into a common "change" abstraction. No
        /// operation produces both kinds at once, so a shared base type would buy nothing and cost a
        /// rewrite of every existing change site.
        /// </remarks>
        public ZOrderChange? Order { get; }

        /// <summary>Changes that would actually move something.</summary>
        public IReadOnlyList<GeometryChange> EffectiveChanges =>
            Changes.Where(c => !c.IsNoOp).ToList();

        /// <summary>True when nothing in the transaction would change, so it is not worth recording.</summary>
        public bool IsEmpty => EffectiveChanges.Count == 0 && (Order is null || Order.IsNoOp);

        /// <summary>The transaction running backwards, ready to apply.</summary>
        public AlignTransaction Inverted() =>
            new AlignTransaction(Label, Changes.Select(c => c.Inverted()).ToList(), Order?.Inverted());

        /// <summary>Builds a transaction from a solve, dropping changes that would not move anything.</summary>
        public static AlignTransaction FromResult(string label, SolveResult result)
        {
            if (result is null) throw new ArgumentNullException(nameof(result));
            return new AlignTransaction(label, result.EffectiveChanges.ToList());
        }

        /// <summary>Builds a transaction that only restacks, changing no geometry.</summary>
        public static AlignTransaction FromOrder(string label, ZOrderChange order)
        {
            if (order is null) throw new ArgumentNullException(nameof(order));
            return new AlignTransaction(label, Array.Empty<GeometryChange>(), order);
        }
    }

    /// <summary>
    /// AlignPro's own undo stack.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This is a labelled journal, not a substitute for PowerPoint's undo. Object-model changes
    /// <em>do</em> reach PowerPoint's undo stack; they are coalesced into one open entry, and the
    /// add-in ends that entry per operation through <c>Application.StartNewUndoEntry</c>, so Ctrl+Z
    /// reverses exactly one AlignPro command. See <c>docs/object-model-findings.md</c>.
    /// </para>
    /// <para>
    /// What this adds over the native stack is a name. PowerPoint's Undo cannot say which operation
    /// it is about to reverse, so the ribbon reads "Undo align left" and this journal is what knows
    /// that. No keystroke is intercepted and no keyboard hook exists.
    /// </para>
    /// <para>
    /// The two stacks are independent, and a native Ctrl+Z does not pop this one. That is survivable
    /// rather than correct: <see cref="GeometryChange"/> holds absolute frames, so undoing here after
    /// a native undo rewrites coordinates the shapes already occupy. The button's label and enabled
    /// state can still be a step ahead of the document.
    /// </para>
    /// <para>
    /// Pure and Office-free: the journal decides <em>what</em> to change and hands back a transaction,
    /// but never touches PowerPoint. Applying is the add-in's job, which is also where a shape that
    /// has since been deleted gets skipped. Not thread-safe by design - it is only ever touched from
    /// the UI thread.
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
