# PowerPoint object-model findings

Measured by `tools/Probe-ShapeGeometry.ps1`. These answers drive the design of
`AlignPro.Geometry` - do not re-derive them from documentation.

- Date: 2026-09-17
- PowerPoint: version 16.0, build 20326
- PowerShell: 7.6.6

## 1. Does Shape.Left/Top/Width/Height return the unrotated frame or the rotated visual bbox?

**UNROTATED FRAME - the OM ignores rotation. AlignPro must compute visual bounds itself.**

```text
before rotation          L=  100.00  T=  100.00  W=  100.00  H=   50.00  Rot=   0.0
after Rotation=45        L=  100.00  T=  100.00  W=  100.00  H=   50.00  Rot=  45.0
centre before (150.00, 125.00)  ->  centre after (150.00, 125.00)
a 100x50 rect at 45 deg has a visual bbox of 106.07 x 106.07
centre invariant under rotation: True
```

## 2. Does Selection.ShapeRange preserve selection order?

**SELECTION ORDER PRESERVED - "last selected is the anchor" is viable.**

```text
added in z-order:  ProbeA, ProbeB, ProbeC
selected in order: ProbeC, ProbeA, ProbeB
ShapeRange returned: ProbeC, ProbeA, ProbeB
NOTE: this selects programmatically. Confirm by clicking the three shapes by hand too -
the UI and the OM do not have to agree.
```

## 3. Does a group behave as one rigid object, preserving internal spacing?

**RIGID ON TRANSLATE, PROPORTIONAL ON RESIZE - align/distribute are safe for groups as-is. Match-size rescales internal spacing, so it needs an explicit decision.**

```text
group before translate   L=   20.00  T=  300.00  W=  260.00  H=   40.00  Rot=   0.0
    child ProbeA   L=   20.00 T=  300.00 W=   60.00
    child ProbeB   L=  120.00 T=  300.00 W=   60.00
    child ProbeC   L=  220.00 T=  300.00 W=   60.00
    left-edge gaps: 100.00, 100.00
group after translate    L=   80.00  T=  270.00  W=  260.00  H=   40.00  Rot=   0.0
    left-edge gaps: 100.00, 100.00  (preserved: True)
group after 1.5x width   L=   80.00  T=  270.00  W=  390.00  H=   40.00  Rot=   0.0
    child ProbeA   L=   80.00 T=  270.00 W=   90.00
    child ProbeB   L=  230.00 T=  270.00 W=   90.00
    child ProbeC   L=  380.00 T=  270.00 W=   90.00
    left-edge gaps: 150.00, 150.00  (scaled: True)
```

## 4. Does an object-model change reach PowerPoint's undo stack?

**NEEDS MANUAL CHECK - see the instructions printed at the end.**

```text
undo probe at start      L=  400.00  T=  300.00  W=  120.00  H=   60.00  Rot=   0.0
moved from script        L=  500.00  T=  380.00  W=  120.00  H=   60.00  Rot=   0.0
CommandBars.ExecuteMso failed: Unexpected HRESULT has been returned from a call to a COM component.
```

## 5. Are TextRange2.BoundLeft/Top/Width/Height reliable?

**USABLE - TextRange2.Bound* returns real geometry. Viable as the TextBounds model.**

```text
textbox frame            L=  300.00  T=  100.00  W=  200.00  H=   29.08  Rot=   0.0
AutoSize=0 WordWrap= -1  BoundL=  307.20 BoundT=  103.60 BoundW=   70.25 BoundH=   21.60
AutoSize=0 WordWrap=  0  BoundL=  307.20 BoundT=  103.60 BoundW=   70.25 BoundH=   21.60
AutoSize=1 WordWrap= -1  BoundL=  307.20 BoundT=  103.60 BoundW=   70.25 BoundH=   21.60
AutoSize=1 WordWrap=  0  BoundL=  307.20 BoundT=  103.60 BoundW=   70.25 BoundH=   21.60
```

## 6. Can layout placeholder bounds be read, for "align to content area"?

**YES - CustomLayout placeholders expose geometry. PlaceholderBounds reference is viable.**

```text
layout: Title Slide
  ph[1] type=3           L=  120.00  T=   88.38  W=  720.00  H=  188.00  Rot=   0.0
  ph[2] type=4           L=  120.00  T=  283.63  W=  720.00  H=  130.37  Rot=   0.0
  ph[3] type=16          L=   66.00  T=  500.50  W=  216.00  H=   28.75  Rot=   0.0
  ph[4] type=15          L=  318.00  T=  500.50  W=  324.00  H=   28.75  Rot=   0.0
  ph[5] type=13          L=  678.00  T=  500.50  W=  216.00  H=   28.75  Rot=   0.0
```

## Undo: measured over three experiments

Object-model changes **do** reach PowerPoint's undo stack. They are **coalesced into one entry per
automation burst**, where a burst appears to be bounded by user interaction with the UI.

Three observations, all consistent with that one rule:

| Experiment | What ran | One Ctrl+Z did |
|---|---|---|
| `Probe-ShapeGeometry.ps1` probe 4 | One long script: `Slides.Add`, dozens of shape adds, deletes, groups and property writes, no user interaction anywhere | **Removed the whole slide** |
| `Probe-UndoInteraction.ps1 -Stage Align` | A human drag, then one scripted move of `ProbeC` (two property writes) | **Reverted exactly that move** — `ProbeC` (300, 310) → (500, 250) |
| `Probe-UndoInteraction.ps1 -Stage Granularity` | Two scripted moves back to back, no interaction between them | **Reverted both** — `ProbeA` and `ProbeC` together |
| Real add-in, `New-TestDeck.ps1` deck | Scripted deck creation, then two ribbon commands, with selection and dropdown changes in between | **Removed every slide** |

The first result looked like "geometry changes don't register", and was recorded that way at first.
It is better explained by coalescing: that script was a single uninterrupted burst that *began* with
`Slides.Add`, so the one entry covering it took the slide with it when undone. `Presentations.Add()`
yields a deck with no slides, which is why undoing the insertion emptied the deck.

The second result rules out "geometry changes don't register" directly, and shows a human edit closes
the group. The third shows that without an intervening interaction, consecutive operations merge.

### Fourth experiment: a ribbon click is not a boundary

Measured with the real add-in loaded. `New-TestDeck.ps1` built a four-slide deck through the object
model; the user then selected shapes, changed the Measure dropdown, clicked **Align left**, then
clicked **Align top**; then pressed Ctrl+Z once.

**Every slide disappeared.** The whole session — the scripted deck creation and both ribbon
operations — was a single undo entry.

That refines the rule. The coalescing group is not closed by user *interaction*; it is closed by a
document *modification* that originates in the UI. The drag in experiment 2 was such a modification.
Selecting shapes, changing a dropdown and clicking a ribbon button are not, so none of them broke the
group, and our two operations joined an entry that had been open since the first scripted slide.

### What this means for AlignPro

An AlignPro operation can be coalesced with an arbitrary amount of preceding object-model work, and
one Ctrl+Z discards all of it.

- In a **human-authored deck**, the practical blast radius is consecutive AlignPro operations: the
  user's own last manual edit closed the previous group, so Ctrl+Z undoes every AlignPro command since
  then, as one step. Surprising, and still wrong, but bounded.
- Where **object-model changes precede** ours — a macro, another add-in, our own scripts — the blast
  radius is unbounded. That is the case measured above, and it destroyed four slides.

`UndoManager` is therefore load-bearing after all, not a convenience, and Ctrl+Z needs intercepting so
the reflex reaches our per-operation stack instead of PowerPoint's coalesced entry.

**One limitation interception cannot fix.** We can claim the keystroke, but not PowerPoint's own Undo
button on the ribbon or the Quick Access Toolbar — that is PowerPoint's own command, and clicking it
still hits the coalesced entry. So interception narrows the hazard rather than removing it.

### Fifth experiment: repurposing does not work in PowerPoint

`customUI` lets an add-in repurpose a built-in command — `<command idMso="Undo" onAction="..."/>` — which
would have covered the buttons as well as the keystroke. It was implemented, the ribbon parsed it (the
tab kept working), and it had no effect.

The add-in's own log settles why. After an Align left, the log holds
`Ran 'Align left': applied=3 missing=0 slideId=256 undoDepth=1` and **no callback line at all**. The
repurposed `onAction` is never invoked. PowerPoint accepts the markup and ignores it. This matches
[MS-CUSTOMUI](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/21316865-fbce-4a3f-a9ff-a8277cce5f2d)
allowing `onAction` only on commands that are simple buttons — PowerPoint's Undo is a split button with
a history dropdown — but the practical answer is simply that it cannot be used.

### Sixth experiment: closing the group ourselves

Since only a *modifying* command closes the group, the search was for one whose effect is invisible.
Measured with `tools/Probe-UndoGrouping.ps1`:

| Candidate | Closes group | Visible change |
|---|---|---|
| `Bold` once | yes | text goes bold |
| `Bold` / `Italic` twice (self-cancelling) | yes | none, but lossy on a *mixed* selection |
| `Bold` on a shape with no text | yes | font state still changes |
| **`Italic` then `ExecuteMso('Undo')`** | **yes** | **none** |

The last is the one to use. A real undo restores the previous formatting exactly, so unlike a double
toggle it is lossless even when the selection's formatting is mixed — and the group boundary survives
undoing the command that created it.

This inverts the whole approach: rather than intercepting undo, AlignPro can close the group
immediately before it applies anything, so PowerPoint's own entry contains exactly one AlignPro
operation. Native Ctrl+Z, the QAT button and the ribbon button then all behave correctly, and no
keyboard hook is needed.

**The one dangerous failure mode, and its guard.** If the `Italic` silently fails to create an undo
entry, the following `Undo` would revert whatever came *before* — the user's own last edit, or the open
coalesced group. That is precisely the disaster being fixed. So the formatting state is read before and
after the toggle, and the `Undo` is only issued when the state actually changed. If it did not, the
closer is skipped and we simply do without the boundary.


## Design implications

**Probe 1 is the load-bearing one.** `Left/Top/Width/Height` is the *unrotated frame*, and the shape's
centre is invariant under rotation. Two consequences:

- Native PowerPoint align really does misalign rotated shapes, because it aligns the frame. The
  `VisualBounds` model is a headline feature, not a nicety.
- Because the centre is invariant, visual bounds are derived purely from frame + rotation:
  `halfW' = (w·|cos θ| + h·|sin θ|)/2`, `halfH' = (w·|sin θ| + h·|cos θ|)/2`, centred on the frame
  centre. No extra OM calls needed.

**Probe 2** clears "last selected is the anchor". Selection order survived (`ProbeC, ProbeA, ProbeB`
against a z-order of A, B, C). This was measured programmatically — still worth clicking three shapes
by hand once, since the UI and the OM need not agree.

**Probe 3** confirms groups are rigid on translate (gaps held at 100, 100) and proportional on resize
(gaps scaled to 150, 150 at 1.5× width). So align and distribute are safe on groups with no special
handling, and `MatchWidth/Height/Both` must **refuse groups by default** — silently rescaling a
diagram's internal spacing is exactly the damage we're trying to prevent.

**Probe 4** needed two follow-up experiments to read correctly — see *Undo: measured over three
experiments* below. Short version: object-model changes do reach the undo stack, coalesced one entry
per uninterrupted automation burst. Separately, `CommandBars.ExecuteMso('Bold')` failed outright with
an HRESULT error, so the "push a neutral command to seed an undo entry" trick is unavailable — but
with coalescing understood, we no longer need it.

**Probe 5** shows `TextRange2.Bound*` returns real geometry, and the 7.2pt offset between frame left
(300.00) and bound left (307.20) is the default 0.1" internal margin — a useful sanity check that
these are true text bounds rather than the frame. Values did not vary across autosize/wrap because
"Align me" fits on one line either way; re-probe with wrapping text before trusting the model for
multi-line shapes.

**Probe 6** confirms `PlaceholderBounds` is viable, and exposes `PlaceholderFormat.Type` so we can
target the body placeholder specifically rather than the whole layout.
