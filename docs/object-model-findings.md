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

**Confirmed by hand on 2026-09-19** — see *Probe 2 by hand* below. The caveat above is closed.

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

## Undo: measured over seven experiments

Object-model changes **do** reach PowerPoint's undo stack. They are **coalesced into one entry per
automation burst**, where a burst appears to be bounded by user interaction with the UI.

Four observations, all consistent with that one rule:

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

The second rules out "geometry changes don't register" directly, and shows a human edit closes the
group. The third shows that without an intervening interaction, consecutive operations merge. The
fourth is examined below.

### Fourth experiment: a ribbon click is not a boundary

Measured with the real add-in loaded. A four-slide deck was built through the object model, then
shapes were selected, a dropdown changed, **Align left** clicked, **Align top** clicked, and Ctrl+Z
pressed once.

**Every slide disappeared.** The scripted deck creation and both ribbon operations were a single undo
entry.

That refines the rule — though the first wording of it here was wrong, and worth correcting rather
than quietly fixing. It said the group is closed by "a document modification originating in the UI",
and that clicking a ribbon button is not one. Native **Align left** is a ribbon button, and it *does*
get its own undo entry: click it twice and the two undo one at a time.

The real distinction is **PowerPoint's own command versus automation writes**. Every built-in command
brackets its own undo entry — the dispatcher opens one, does the work in native code, closes it. A
click on *our* button runs no command at all: PowerPoint invokes a callback into managed code, which
then writes through the object model, indistinguishable from a macro. The drag in experiment 2 closed
the group because it was a native operation taking its own entry, not because a human's hand was on
the mouse. So our two operations joined an entry open since the first scripted slide.

### What this means for AlignPro

An AlignPro operation can be coalesced with an arbitrary amount of preceding object-model work, and
one Ctrl+Z discards all of it.

- In a **human-authored deck**, the practical blast radius is consecutive AlignPro operations: the
  user's own last manual edit closed the previous group, so Ctrl+Z undoes every AlignPro command since
  then, as one step. Surprising, and still wrong, but bounded.
- Where **object-model changes precede** ours — a macro, another add-in, our own scripts — the blast
  radius is unbounded. That is the case measured above, and it destroyed four slides.

That made `UndoManager` load-bearing rather than a convenience, and three attempts follow below. The
fifth and seventh failed; the **ninth succeeded**, and native Ctrl+Z is now safe — so the paragraphs
between here and there describe the problem as it stood, not as it stands.

Worth keeping for the reasoning: interception was considered and rejected, because claiming the
keystroke cannot claim PowerPoint's own Undo button on the ribbon or the Quick Access Toolbar. Those
are PowerPoint's own command and would still hit the coalesced entry, so interception could only ever
narrow the hazard. Fixing the entry itself covers all three routes at once, which is what the ninth
experiment does.

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

That appeared to invert the whole approach: rather than intercepting undo, close the group immediately
before applying anything, so PowerPoint's own entry contains exactly one AlignPro operation.

**It does not work from a ribbon callback.** See the seventh experiment below. `UndoBoundary` is kept
in the source, documented and unused, because the finding is real and might yet be rescued.

**The dangerous failure mode it was built to avoid.** If the `Italic` silently fails to create an undo
entry, the following `Undo` reverts whatever came *before* — the last manual edit, or the open coalesced
group. That is precisely the disaster being fixed, so the formatting state is read before and after the
toggle and the `Undo` is only issued when the state actually changed.

### Seventh experiment: ExecuteMso is deferred inside a ribbon callback

With the group-closer wired in, clicking **Align Left** silently did nothing. The add-in's log showed
the operation running and reporting `applied=3`, with every shape still in its original position.

PowerPoint will not run an undo while a command is executing, so `ExecuteMso("Undo")` inside a ribbon
callback is *deferred*. The italic toggle applied synchronously, the geometry was written, and only
then did the deferred undo run — reverting the geometry instead of the toggle.

Driving exactly the same code through the automation surface works perfectly. That difference is the
whole bug, and it carries a second lesson: a harness calling in over cross-process COM has no command
in flight, so it cannot reproduce a ribbon callback's context. All seven checks of the COM harness
passed while the add-in was visibly broken. `tools/Test-RibbonClicks.ps1` exists because of this, and
clicks the real ribbon through UI Automation.

**Where that left undo, at the time.** Neither mechanism tried made native undo safe: repurposing is
ignored, and the toggle-and-undo boundary cannot be established from the context the add-in runs in.
That conclusion stood until the ninth experiment below, which reaches the same goal with a single
object-model call and makes this whole line of attack unnecessary.


## Eighth experiment: Shape.ZOrder is *not* deferred in a ribbon callback

Asked because of the seventh: `ExecuteMso("Undo")` is deferred while a command is executing, which
silently broke Align left. Ordering is built entirely on `Shape.ZOrder`, so the same question had to
be settled before trusting it — and the COM harness cannot answer it, for the reason that experiment
gives.

Measured by `tools/Test-RibbonClicks.ps1`, clicking the real button through UI Automation, against
slide 7 of the sample deck.

```text
Clicking Stack puts the first selected on top   Card1=7 Card2=5 Card3=3  (higher is nearer the front)
Unselected shapes keep their layer              BarA 4->4, BarB 6->6
Clicking Reverse flips the stack                Card1=3 Card3=7
Undo restores the previous stacking             Card1=7 Card3=3
```

**`Shape.ZOrder` applies synchronously from a ribbon callback**, like the `Left`/`Top` writes and
unlike `ExecuteMso`. The distinction is between an object-model write, which is executed in place, and
a *command*, which PowerPoint queues until the current command finishes.

The second line is the one worth keeping: `BarA` and `BarB` were not selected and held z-positions 4
and 6 throughout, so the restacking really does happen within the slots the selection already
occupied. That is what separates this from Bring to front.

`ZOrderPosition` is read-only, so an ordering is realised by calling `BringToFront` on each shape in
turn from the back of the target list forwards. Applied to a complete ordering that lands exactly on
it, and the previous ordering is then a complete instruction for undoing it.

## Probe 2 by hand: selection order survives a mouse

Probe 2 measured selection order with the shapes selected **programmatically**, and flagged that the
UI and the object model need not agree. That mattered more once ordering and the match-size cascade
were built, because both rest entirely on selection order — and every harness here selects over COM,
including the ribbon-click one, which clicks the *buttons* for real but still selects the *shapes*
through the object model. Nothing automated could close it.

Measured by clicking three shapes by hand on slide 7 of the sample deck, in an order deliberately
unlike their z-order, then reading the range back over COM.

```text
clicked in order:    Card3, Card1, Card2
ShapeRange returned: Card3, Card1, Card2
                     z=7    z=3    z=5
slide z-order, back to front: Card1(3), BarA(4), Card2(5), BarB(6), Card3(7)
```

**SELECTION ORDER PRESERVED FOR A HAND-CLICKED SELECTION.** A fallback to z-order would have returned
`Card1, Card2, Card3`; the range came back in click order instead, with the z-positions scrambled
relative to it.

So "the anchor is the last shape you selected" holds for real users and not just for automation, and
the same goes for Order and for the cascade. The explicit *pin anchor* button that
`Probe-ShapeGeometry.ps1` names as the fallback is not needed.

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


## Ninth experiment: `StartNewUndoEntry` is the API this needed all along

Prompted by a question the earlier write-up could not answer: if a ribbon click is not a boundary, why
does native **Align left** followed by native **Align middle** undo one step at a time? It does, and
the answer is above — built-in commands bracket their own entries. Stated that way the requirement was
never "manufacture a UI modification", only "end the open automation entry" — and PowerPoint has a
first-class method for exactly that, found by reflecting over the interop assembly the add-in already
references:

```text
Microsoft.Office.Interop.PowerPoint._Application
  void StartNewUndoEntry()      // dispid 2067, no arguments
```

It appears nowhere in the project's history. `UndoManager`'s remark that PowerPoint "has no
`UndoRecord` equivalent" is what steered past it: true of the name, since Word and Excel spell it
differently, and wrong in substance.

Measured first over COM, with no ribbon involved:

| What ran | One Ctrl+Z did |
|---|---|
| Plain OM write, no boundary | reverted it |
| `StartNewUndoEntry()`, then a write | reverted it — the call does not cost undoability |
| Boundary, write, boundary, write | **reverted only the second write; the first held** |

Then through a real ribbon click, which is the context that broke the sixth experiment's approach.
Both of PowerPoint's own undo routes were driven: its Undo button clicked through UI Automation, and
the Ctrl+Z keystroke.

```text
clicked 'Align left':  Plain1 420 -> 330
PowerPoint's own Undo button:  Plain1 back to 420   REVERTED
Ctrl+Z keystroke:              Plain1 back to 420   REVERTED
deck intact: 13 slides
```

**`StartNewUndoEntry` works from a ribbon callback**, where `ExecuteMso("Undo")` does not. That is the
eighth experiment's distinction again, and it is the whole reason this succeeds where the sixth failed:
an object-model *method* executes in place, a *command* is queued until the current one finishes.

`ChangeApplier` now calls it before writing, in both the geometry and the restacking path. `Ctrl+Z`,
the ribbon Undo button and the Quick Access Toolbar each reverse exactly one AlignPro operation, and
`tools/Test-RibbonClicks.ps1` covers all of it.

### A harness trap worth knowing about

The first two runs of the new Ctrl+Z cases failed, and the product was fine. **Invoking a ribbon button
through UI Automation leaves keyboard focus on that button, and a Ctrl+Z sent in that state never
reaches the document.** It silently does nothing, which is indistinguishable from a broken undo —
`Send-NativeUndo` sends `{ESC}` first, and without it the harness reports a failure that does not
exist.

Two other ways to lose an afternoon here, both hit while measuring this:

- **Never click a verb on a selection it cannot change.** "Nothing to change" is a modal message box,
  and a modal dialog blocks PowerPoint's UI thread, so every COM call afterwards hangs. Displace a
  shape first so the verb has real work. The harness docstring already warned about this.
- `CommandBars.ExecuteMso('Undo')` over COM was not the culprit when a probe hung, despite probe 4's
  HRESULT failure making it the obvious suspect. The blocked UI thread from a dialog was.


## Duplicates and paths: probes 7 to 10

The spike planned as step 8, before Duplicate and Distribute along curve were built.

Measured by `tools/Probe-DuplicateAndPaths.ps1` on 2026-09-23, PowerPoint 16.0 build 20326.

### 7. Does Shape.Duplicate stay inside the entry StartNewUndoEntry opened?

**YES - one undo removed the copy and its edits, and the move before the boundary held.**

```text
shapes before undo: 2, after: 1
copy gone: True
DupA.Left after undo: 200.0 (moved to 200 before the boundary)
Measured over COM. The ribbon-click harness repeats this through a real click.
```

### 8. Where does a duplicate land in the z-order, and does its offset need undoing?

**ON TOP, OFFSET (12.0, 12.0) - a duplicate lands frontmost; writing an absolute position straight after simply overrides the offset.**

```text
Back                       L=  100.00 T=  300.00 W=   60.00 H=   40.00 Rot=   0.0 Z=1
Middle                     L=  180.00 T=  300.00 W=   60.00 H=   40.00 Rot=  20.0 Z=2
Front                      L=  260.00 T=  300.00 W=   60.00 H=   40.00 Rot=   0.0 Z=3
Middle.Duplicate()         L=  192.00 T=  312.00 W=   60.00 H=   40.00 Rot=  20.0 Z=4
offset from original: (12.00, 12.00); copy on top: True; rotation copied: True
after Left=500 Top=200     L=  500.00 T=  200.00 W=   60.00 H=   40.00 Rot=  20.0 Z=4
Range order: Front, Back, Middle
its Duplicate() lefts, in returned order: 112, 192, 272  (sources at Front=260, Back=100, Middle=180)
```

### 9. Which coordinate space do Shape.Nodes report for a rotated or flipped freeform?

**ROTATED AND FLIPPED - nodes unchanged by rotation: False; unchanged by a horizontal flip: False.**

```text
freeform                   L=  100.00 T=  100.00 W=  300.00 H=  160.00 Rot=   0.0 Z=1
nodes as built:   (100.0,100.0) (300.0,100.0) (300.0,200.0) (320.0,260.0)c (380.0,260.0)c (400.0,200.0)c
segment types:    0,0,0,1,1,1   (1 = curve)
after Rotation=90          L=  100.00 T=  100.00 W=  300.00 H=  160.00 Rot=  90.0 Z=1
nodes rotated:    (330.0,30.0) (330.0,230.0) (230.0,230.0) (170.0,250.0)c (170.0,310.0)c (230.0,330.0)c
HorizontalFlip = -1
nodes flipped H:  (400.0,100.0) (200.0,100.0) (200.0,200.0) (180.0,260.0)c (120.0,260.0)c (100.0,200.0)c
closed triangle:  (500.0,100.0) (600.0,100.0) (600.0,200.0)
AddLine(100,400 -> 300,350) L=  100.00 T=  350.00 W=  200.00 H=   50.00 Rot=   0.0 Z=3
line Type=9 HorizontalFlip=0 VerticalFlip=-1
line nodes: none - Nodes.Count is 0
```

### 10. How do the Arc autoshape's adjustments map to start and end angles?

**ADJ1 = START, ADJ2 = END, degrees clockwise from three o'clock, as DIRECTIONS from the ellipse's centre. The frame is NOT the ellipse's box: it is the box of the arc together with the centre. Rotating an arc changes its reported frame, so a rotated arc cannot be rebuilt.**

```text
AddShape 300x100       L=  100.00 T=  200.00 W=  300.00 H=  100.00 rot=    0 flipH= 0 adj=-90.00,0.00
adjusted to 0..45      L=  100.00 T=  300.00 W=  300.00 H=   94.87 rot=    0 flipH= 0 adj=0.00,45.00
  direction predicts L=100 T=300 W=300 H=94.87; parameter predicts H=70.71
adjusted to -30..200   L= -200.05 T=  213.40 W=  600.10 H=  186.62 rot=    0 flipH= 0 adj=-30.00,-160.00
rotated 30             L=  100.00 T=  300.00 W=  259.81 H=  173.21 rot=   30 flipH= 0 adj=0.00,45.00
flipped horizontally   L= -200.00 T=  300.00 W=  300.00 H=   94.87 rot=    0 flipH=-1 adj=0.00,45.00
Width = 150            L=  100.00 T=  300.00 W=  150.00 H=   94.87 rot=    0 flipH= 0 adj=0.00,63.43
```

### What this means for AlignPro

- **Probe 7:** `ShapeCreator` needs nothing beyond the `StartNewUndoEntry` every verb already calls.
  One Ctrl+Z removes every copy a Duplicate made and nothing before it. The ribbon-click harness
  repeats this through a real click.
- **Probe 8:** `ShapeRange.Duplicate` hands its copies back in **z-order, not in the order of the
  range** - `Front, Back, Middle` came back as copies of `Back, Middle, Front`. Matching copies to
  placements by position would therefore put copies in the wrong place whenever the selection
  order differs from the stacking, which is the usual case. So `ShapeCreator` duplicates one shape
  at a time, back to front, and each copy's identity is certain. The 12pt offset needs no undoing,
  because the placement written straight after is absolute.
- **Probe 9:** nodes come back **exactly where they are drawn**, rotation and flips already applied.
  The plan had assumed the opposite, and turning them again would have turned a rotated path twice.
  `CurveSolver` takes path nodes as they are and transforms only the kinds it builds from the frame.
  Two smaller points: a straight line reports no nodes at all, so it is its own curve kind, built
  corner to corner from the frame and its flips; and a closed freeform does **not** repeat its first
  node, so there is no way to tell it from an open one. Freeforms are followed from their first node
  to their last, and the closing segment is not used - use an oval to go all the way round.
- **Probe 10** was the surprise. An Arc's frame is **not** the box of its ellipse, which is what the
  plan assumed. It is the box of the arc *together with the ellipse's centre* - a pie slice's box -
  and PowerPoint resizes it whenever the angles change, keeping the ellipse fixed. So the default
  quarter has its centre at the frame's bottom-left corner and an ellipse twice the frame each way.
  The angles are directions from the centre, not the ellipse's parameter: the 94.87pt height fits
  only the direction reading. `CurveSolver.SampleArc` rebuilds the ellipse from the frame and the two
  angles - one unknown, the ratio of the radii, solved by bisection - and reproduces all three frames
  above exactly. Flipping mirrors the frame about the ellipse's centre, so the arc is still the
  unflipped arc mirrored within its current frame. **Rotation changes the reported frame's size**,
  which no longer describes the ellipse at all, so a rotated arc is refused with the reason rather
  than followed wrongly. Resizing rewrites the angles to match, so the frame and angles always agree.

  The first version of this probe tried to answer the question by exporting the slide and reading
  pixels where each reading put the stroke. It found no stroke anywhere, because the stroke was not
  in the frame at all. Rendering by eye is what showed why.
