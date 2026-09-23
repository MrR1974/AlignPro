# AlignPro — development notes

The companion to [`README.md`](../README.md), which is the guide for people who just want to use
AlignPro. This file is for people changing it: how the pieces fit, how to build and test, and what was
measured rather than assumed.

## Why

PowerPoint's align and distribute tools have no anchor/key-object alignment, no exact spacing, no
"make same size", no grid tidying, and they align the raw object-model rectangle — which
[ignores rotation entirely](object-model-findings.md), so rotated shapes end up visually wrong.

## Layout

| Path | What's in it |
|---|---|
| [`src/AlignPro.Geometry/`](../src/AlignPro.Geometry/) | The solver and the undo journal. Pure logic, **zero Office references**, `netstandard2.0` |
| [`src/AlignPro.AddIn/`](../src/AlignPro.AddIn/) | The VSTO add-in: ribbon, selection adapter, apply pipeline, undo boundary, automation surface. `net48` |
| [`tests/AlignPro.Geometry.Tests/`](../tests/AlignPro.Geometry.Tests/) | xUnit suite on `net8.0` |
| [`install.ps1`](../install.ps1) | The one-line remote installer. Downloads a release, verifies its checksum, hands it to the installer inside |
| [`tools/Install-AlignPro.ps1`](../tools/Install-AlignPro.ps1) | The installer itself. Ships inside the release zip; `install.ps1` calls this |
| [`tools/Uninstall-AlignPro.ps1`](../tools/Uninstall-AlignPro.ps1) | Reverses it, and nothing else. Installed alongside the add-in so Add/Remove Programs can call it |
| [`tools/Install AlignPro.cmd`](../tools/) | Double-click shim for the no-terminal route; ships inside the zip |
| [`tools/New-Package.ps1`](../tools/New-Package.ps1) | Packages a built Release into the release zip and its `.sha256` |
| [`tools/New-Installer.ps1`](../tools/New-Installer.ps1) | Builds the MSI, for managed deployment |
| [`tools/New-Release.ps1`](../tools/New-Release.ps1) | Tests, builds Release, and produces all three release assets |
| [`tools/New-DevSigningCertificate.ps1`](../tools/New-DevSigningCertificate.ps1) | Creates the machine-local certificate VSTO needs to build |
| [`tools/Probe-ShapeGeometry.ps1`](../tools/Probe-ShapeGeometry.ps1) | Measures PowerPoint's object model; doubles as the integration harness |
| [`tools/Probe-UndoGrouping.ps1`](../tools/Probe-UndoGrouping.ps1) | How PowerPoint groups undo entries, and what closes a group |
| [`tools/Probe-DuplicateAndPaths.ps1`](../tools/Probe-DuplicateAndPaths.ps1) | Duplicate's undo and z-order, freeform nodes, and what an Arc's frame really is |
| [`tools/New-SampleDeck.ps1`](../tools/New-SampleDeck.ps1) | Builds the **saved** 17-slide sample deck, one slide per capability |
| [`tools/New-TestDeck.ps1`](../tools/New-TestDeck.ps1) | Builds a throwaway scratch deck, never saved |
| [`tools/Test-AlignProEndToEnd.ps1`](../tools/Test-AlignProEndToEnd.ps1) | Drives the add-in inside PowerPoint and asserts the results, no clicking |
| [`tools/Test-RibbonClicks.ps1`](../tools/Test-RibbonClicks.ps1) | Clicks the real ribbon through UI Automation and asserts the results |
| [`object-model-findings.md`](object-model-findings.md) | What the probe measured, and what it means for the design |

## Installing a development build

`install.ps1` installs a *published release*, which is not what you want while working on the add-in.
Build the solution and run the installer against the build output instead:

```powershell
.\tools\Install-AlignPro.ps1 -Source .\src\AlignPro.AddIn\bin\Debug
```

It skips the sample deck and the uninstaller when they are not beside the add-in, which is the case in
a build folder, and therefore writes no Add/Remove Programs entry. Everything else — registration,
trust — is identical to what an end user gets.

Visual Studio also grants trust to what it builds, so an F5 debug session needs none of this.

## How it is distributed, and why

The release carries three assets: `AlignPro-<version>.zip`, its `.sha256`, and an MSI.

The zip is the real distribution. The one-liner in the README fetches `install.ps1`, which downloads
the zip, checks it against the published hash, and runs the installer inside it. **Nothing goes through
the browser**, which is the entire point: SmartScreen's application-reputation check applies to
executables a browser downloads, and an unsigned file has no reputation. Microsoft documents that an
unsigned file's reputation "must build for each new version of your files, starting with zero", and
that a self-signed certificate behaves exactly like no signature — so an unsigned MSI is warned about
on every release forever. That is not a warning any amount of download volume fixes.

The same route also avoids the mark of the web, which was the other half of the problem. Measured on
Windows 11:

| | Result |
|---|---|
| `Invoke-WebRequest` output file | no `Zone.Identifier` stream — no mark |
| `Expand-Archive` from a marked zip | extracted files carry no mark (Explorer's extractor *does* propagate it) |
| A marked `.ps1` run by path | `AuthorizationManager check failed`, **even at `Unrestricted`** |
| The same text via `Invoke-Expression` | runs |

That last pair is why the old "download the zip and run the script" instructions were fragile, and why
`install.ps1` hands over to the packaged installer through `[ScriptBlock]::Create` rather than by path:
script *files* are governed by the execution policy, script blocks are not.

The MSI is kept for Intune and Group Policy, where an installer package is what the tooling expects and
no browser is involved. It is not what a person downloads.

## Detecting the prerequisites

All three prerequisite checks read HKLM through
`[Microsoft.Win32.RegistryKey]::OpenBaseKey(..., 'Registry64' | 'Registry32')` rather than through an
`HKLM:` path, and that is not fussiness. A path is resolved relative to the bitness of whichever
PowerShell the user launched: 64-bit PowerShell sees the native view, 32-bit PowerShell is silently
redirected into `WOW6432Node` and cannot see the native view at all. Office, the VSTO runtime and
Windows do not agree on which view they register in, so a path-based check passes or fails depending
on which shell someone happened to open.

Measured on one 64-bit Click-to-Run machine:

| Key under `SOFTWARE` | Native (64-bit) | `WOW6432Node` (32-bit) |
|---|---|---|
| `Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE` | present | present |
| `Microsoft\NET Framework Setup\NDP\v4\Full` | present | present |
| `Microsoft\VSTO Runtime Setup\v4R` | **absent** | present, 10.0.60910 |
| `Microsoft\VSTO Runtime Setup\v4` | **absent** | present, 10.0.60910 |

The runtime check is also the loosest of the three: it accepts `v4R` or `v4` in either view, and falls
back to looking for `VSTOInstaller.exe` under either Program Files tree.

**And when it still finds nothing, it warns and installs anyway.** A user reported the old check
refusing to install on a machine that had Office — and pointing at a Microsoft download URL that had
since started returning 404. Weigh the two failure modes: a false negative blocks someone whose
machine is fine and leaves them nowhere to go, while a false positive installs some files and a few
HKCU values onto a machine where the tab then does not appear, which the closing note tells them how
to fix. The second is plainly the better one to be wrong in.

PowerPoint is checked first for the same reason. It used to be checked last, so a machine with no
PowerPoint at all was told the VSTO runtime was missing — true, but not the reason, and not something
installing the runtime would fix.

The runtime download link is `https://aka.ms/VSTORuntimeDownload`, a redirector, deliberately: the
numeric Download Center ID it replaced stopped resolving.

ClickOnce deployment from a URL was considered and rejected. The ClickOnce trust prompt is governed by
`HKLM\SOFTWARE\MICROSOFT\.NETFramework\Security\TrustManager\PromptingLevel`, whose default for the
`Internet` zone is `AuthenticodeRequired` — a trust prompt is only offered for a certificate that
identifies the publisher. A self-signed manifest served over HTTPS gets no prompt at all; it simply
fails to load.

## Two constraints worth knowing up front

**The add-in targets `net48`, not .NET 8.** VSTO cannot target .NET Core/5+ — the two runtimes cannot
share a process, and Microsoft will not be updating the COM add-in platform. This is permanent, and
it is the one place this repo deviates from its .NET 8 convention.

**The dotnet CLI cannot touch the solution — only the two portable projects.** `AlignPro.AddIn` is an
old-style VSTO project whose `$(VSToolsPath)` import resolves into Visual Studio, so
`dotnet build AlignPro.sln` and `dotnet test AlignPro.sln` both fail to even load it. That is expected,
and the split is deliberate: the solver is `netstandard2.0`, so the engine and its whole test suite
build and run on the dotnet CLI with no Visual Studio and no PowerPoint.

```powershell
# The engine and its tests - no Visual Studio needed. Target the PROJECT, not the solution.
dotnet test tests\AlignPro.Geometry.Tests\AlignPro.Geometry.Tests.csproj

# Everything, including the add-in - needs Visual Studio's MSBuild
& "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" `
    AlignPro.sln /p:Configuration=Debug /p:VisualStudioVersion=17.0

# Measure PowerPoint's own behaviour (opens a new, never-saved presentation)
.\tools\Probe-ShapeGeometry.ps1 -KeepOpen -OutFile docs\object-model-findings.md
```

### First build after cloning

VSTO refuses to build without signed ClickOnce manifests, and the signing certificate is machine-local
so it is not in the repository. Generate one once:

```powershell
.\tools\New-DevSigningCertificate.ps1
```

That writes `AlignPro.AddIn_TemporaryKey.pfx` and `Signing.props` into the add-in project, both
gitignored. Skip it and the build fails with *"Cannot build because the ClickOnce manifest signing
option is not selected"*. The certificate is self-signed and trusted only on the machine that made it
— shipping to colleagues needs a real code-signing certificate, which is a separate step.

## Testing

Three layers, because each catches what the others cannot.

```powershell
dotnet test tests\AlignPro.Geometry.Tests\AlignPro.Geometry.Tests.csproj   # 231, no PowerPoint
.\tools\Test-AlignProEndToEnd.ps1                                          # 34, PowerPoint via COM
.\tools\Test-RibbonClicks.ps1                                              # 25, real ribbon clicks
```

The third exists because the second is blind to a whole class of bug. It drives the add-in over
cross-process COM, where no Office command is in flight - which is **not** the context a ribbon
callback runs in. PowerPoint defers some operations while a command is executing, and a shipped bug
where Align Left silently did nothing when clicked passed all seven of those checks.

`Test-RibbonClicks.ps1` uses COM only for setup and assertions and clicks the actual ribbon button
through UI Automation, so Office dispatches the command exactly as it would for a person. Reintroducing
that bug deliberately confirms the split is real: the COM harness still reports 7 of 7 passing, while
the click harness reports *"Clicking Align Left moves shapes: FAIL - 0 of 4 shapes moved"*.

It deliberately avoids operations that raise a message box, since a modal dialog blocks PowerPoint's
UI thread; those paths stay covered through the automation surface.

## Cutting a release

```powershell
.\tools\New-Release.ps1 -Version 1.0.0
```

Runs the geometry tests, builds Release, and writes three files to `dist\`:

| File | What it is |
|---|---|
| `AlignPro-<version>.zip` | The payload: add-in, sample deck, both scripts, the `.cmd` shim and a readme |
| `AlignPro-<version>.zip.sha256` | Its checksum, sha256sum-style, which `install.ps1` verifies |
| `AlignPro-<version>.msi` | The same install for Intune and Group Policy. `-SkipInstaller` omits it, and with it the WiX dependency |

**The zip and its `.sha256` must both be attached to the GitHub release, under exactly those names.**
`install.ps1` looks them up by name through the releases API and refuses to install without the
checksum, so a release missing it is a broken release.

The checksum is an integrity check, not a supply-chain one — it travels in the same release as the file
it describes. The trust anchor is HTTPS and GitHub, the same as for `install.ps1` itself. Worth being
honest about that in any release notes rather than implying more.

`New-Release.ps1` deliberately publishes nothing. It writes the files and prints the `gh release create`
command; running it stays a separate, deliberate act.

Two things to get right around it:

- **Bump `AssemblyVersion` and `AssemblyFileVersion`** in `src/AlignPro.AddIn/Properties/AssemblyInfo.cs`
  to match, and commit that before cutting the release, so the installed DLL says which version it is.
  Nothing enforces this - 1.0.1 shipped with the assembly still reading 1.0.0.0, the tag having served
  as the version of record.
- **Name the repository when you create the release.** `gh release create` defaults to the remote it
  infers from the checkout, which is not necessarily the one being released. Pass `--repo` explicitly,
  and `--target <commit>` so the tag lands on the commit the assets were built from rather than on
  whatever the default branch points at by then.

There is no CI build for the add-in, and that is not an oversight: a VSTO project needs Visual Studio
with the Office/SharePoint workload, which stock hosted runners do not have. The geometry engine and
its tests build on the dotnet CLI alone and could be run in CI happily.

## How the solver is shaped

Every command is one request: `(Verb, Reference, BoundsModel, Options)`.

- **Verb** — `AlignLeft/Right/Top/Bottom/CentreH/CentreV`, `DistributeH/V`,
  `MatchWidth/Height/Both`, `MatchRotation`, `GridArrange`
- **Reference** — `Anchor`, `SelectionBounds`, `Slide`, `SlideMargins`, `PlaceholderBounds`
- **BoundsModel** — `ShapeFrame` (PowerPoint's own), `VisualBounds` (rotation-aware), `TextBounds`
- **DistributeMode** — what the distribute verbs actually space evenly: `LeadingEdge`
  (left-to-left horizontally, top-to-top vertically), `Centre`, `TrailingEdge` (right-to-right,
  bottom-to-bottom), or `Gap` (the visible space between shapes). Identical when every shape is the
  same size; they diverge the moment sizes differ
- **SizeMarginMode** — whether the match-size margin applies at all (`None`), once to every shape
  (`Uniform`), or cumulatively along the selection (`Cascade`). The margin is measured **per side**,
  so one step takes twice it off each dimension — which is what makes `Cascade` plus
  `ResizeOrigin.Centre` come out as even concentric rings. Note there are now two settings called a
  margin: this one and the slide inset behind `ReferenceTarget.SlideMargins`. They are unrelated.
  The ribbon only ever sends a positive margin and a **Direction** (Shrink / Grow); Grow is the
  controller negating it, because the solver has always grown shapes on a negative margin

**Ordering is the one verb that lives outside this request.** `ZOrderSolver` is a separate entry
point with its own `OrderVerb`, because the align solver is defined by emitting frame coordinates and
restacking emits no geometry at all. It works in terms of the slide's whole stacking order, back to
front, and rewrites only the slots the selection already occupies — so unselected shapes keep their
layer. `ZOrderPosition` is read-only in the object model, so the applier realises an ordering by
calling `BringToFront` on each shape in turn from the back of the target list forwards; the ordering
that was there before is itself the complete undo instruction.

**Two more verbs live outside it for the same kind of reason.** `DuplicateSolver` emits a recipe for
shapes that do not exist yet - per copy, per original, a frame and an angle - so it cannot be a list
of changes to existing shapes. One step is a rigid transform, turn about the pivot then move, and copy
*k* is the step applied *k* times. `CurveSolver` does emit ordinary changes, but needs the anchor's
curve as well as its frame, which no other verb does. It flattens every kind of curve - oval, Arc,
line, freeform - to one polyline and places by distance along it, so one placement rule covers all
four.

**A change can carry an angle.** `GeometryChange` has optional `OldRotation`/`NewRotation`: null means
"leave the angle alone", which is what every align verb sends, so none of them can clobber a rotation
they never read. Match rotation and Distribute along curve set it. PowerPoint turns a shape about its
frame's centre and reports the unrotated frame whatever the angle, so a rotation-only change leaves
the frame exactly where it was.

**A transaction can carry created shapes.** `ShapeCreation` holds a duplicate's recipe and the keys
of the shapes it made. Undo deletes them; redo re-runs the recipe from the originals and, since
PowerPoint gives recreated shapes new ids, rewrites the keys through `Rekey` - the one mutable object
in the journal, shared between a transaction and its inverse so the next undo sees the new keys.

The six align edges are the verbs; anchor, slide and rotation awareness are the two orthogonal axes
crossed over them. That is why a large feature list comes out of one small engine.

**The central invariant:** the solver reasons in whichever bounds space was asked for, but always
emits *frame* coordinates, because the frame is the only thing PowerPoint lets us write. Translation
is rigid, so a delta measured in visual or text space is the same delta in frame space — which is how
align and distribute get rotation- and text-awareness for free.

Three design decisions that came out of measurement rather than preference:

- **Groups are always one object.** The solver never descends into a group, so translation preserves
  internal spacing for free. Resize verbs *refuse* groups by default, because scaling a group scales
  the gaps between its children.
- **Match-size works in frame space** whatever bounds model is requested. Matching a rotated shape's
  visual width is ill-posed — at 90° it is driven entirely by the frame's height — and "make these the
  same size as that one" means frame size anyway.
- **Every operation opens its own native undo entry.** PowerPoint's built-in commands each bracket
  their own undo entry — which is why native Align left then Align middle undo one at a time — but
  [object-model writes get no such bracket](object-model-findings.md) and accumulate into one open
  entry. A click on AlignPro's button is not a PowerPoint command; it is a callback into managed code
  that then writes through the object model. So operations used to join whatever entry was already
  open, and one Ctrl+Z could discard an unbounded amount of earlier work. Measured: two commands on a
  scripted deck, one Ctrl+Z, every slide gone.

  `Application.StartNewUndoEntry` ends that entry, and `ChangeApplier` calls it before writing.
  Ctrl+Z, the ribbon Undo button and the Quick Access Toolbar now each reverse exactly one AlignPro
  operation. Two earlier attempts failed and are worth not repeating: repurposing the built-in Undo
  (PowerPoint parses `<command idMso="Undo">` and never invokes the callback), and closing the group
  with a formatting toggle plus `ExecuteMso("Undo")`, which works from outside PowerPoint but not
  from a ribbon callback, because a *command* is deferred while another is executing. The distinction
  that matters is that `StartNewUndoEntry` is an object-model *method*, so it executes in place.

- **AlignPro still keeps its own undo, for the label.** PowerPoint's Undo cannot say which operation
  it is about to reverse; the ribbon reads "Undo align left" because `UndoManager` knows. The two
  stacks are independent and a native Ctrl+Z does not pop ours, so the button's label can be a step
  ahead of the document. Harmless today — changes carry absolute frames — but it is the loose end.

## Status

| Phase | State |
|---|---|
| 0. Object-model spike | **Done** — six probes plus two follow-up undo experiments |
| 1. Geometry engine + tests | **Done** — 161 tests passing |
| 1b. Undo journal | **Done** — `UndoManager` and `AlignTransaction`, pure and fully tested |
| 2. VSTO shell: ribbon, selection adapter, apply pipeline | **Done** — add-in loads and connects in PowerPoint |
| 3. Verbs wired to the ribbon | **Done** — all twelve verbs, reference/measure/spacing controls, confirmed by hand against the sample deck |
| 3b. Undo coalescing | **Done** — `Application.StartNewUndoEntry` gives each operation its own native undo entry; verified by real ribbon clicks plus Ctrl+Z |
| 3c. Automated end-to-end tests | **Done** — 231 unit tests, 34 COM checks, and 25 real ribbon clicks |
| 4. Keyboard hook and bindings | **Not doing** — a deliberate decision, not an omission. It was originally how Ctrl+Z would be protected, and that need went away; as pure convenience it does not justify a global keyboard hook, the riskiest component in the plan. Revisit if daily use makes the ribbon feel slow |
| 5. Distribution | **Done** - one-line remote install, a zip for the no-terminal route, and an MSI for managed deployment. Installer and uninstaller tested end to end. A signed channel is deferred until there is demand, and is not currently available to this publisher |
| 6. Grow or shrink by the match-size margin | **Done** — Direction dropdown; a negative margin is refused |
| 7. Rotation in changes, then Match rotation | **Done** |
| 8. Object-model spike for duplicate and paths | **Done** — probes 7 to 10; two of the plan's assumptions were wrong, see [Next features](#next-features) |
| 9. Created shapes in transactions, then Duplicate | **Done** |
| 10. Distribute on a circle, arc or path | **Done** — rotated arcs refused, closed freeforms followed as open |
| 11. Tidy: snap near-alignments | **Planned** — see [Tidy](#tidy) |
| 12. Tidy: even out near-even spacing | **Planned** |
| 13. Tidy on the ribbon | **Planned** |
| 14. Tidy the whole slide | **Planned** — only after 13 has had real use |

### Next features

Planned 2026-09-23 and built the same day, in this order. Each step has unit tests, COM and
ribbon-click checks, and a sample-deck slide (slides 10 and 15 to 17). What follows records the decisions, and
where the build departed from the plan because the spike said otherwise.

**6. Grow or shrink by the margin.** Ribbon-only, as planned: a **Direction** dropdown (Shrink / Grow)
beside **Apply margin**. The margin box and `SetSizeMargin` refuse a negative number and point at
Grow, so there are never two controls that can cancel each other. With Cascade and Grow the shape
selected first ends largest, which the screentips and README say.

**7. Rotation in changes, then Match rotation.** `GeometryChange` carries optional angles and
`ChangeApplier` writes `Rotation` only when one is present. **Rotation** sits in the Match group: every
shape takes the anchor's angle, flips untouched, groups allowed.

**8. Object-model spike.** [Probes 7 to 10](object-model-findings.md#duplicates-and-paths-probes-7-to-10).
Duplicate stays inside the undo entry and lands on top, as hoped. Two assumptions in the plan were
wrong: freeform nodes are reported *as drawn*, already rotated and flipped, and an Arc's frame is
the box of the arc and its centre, not of its ellipse.

**9. Created shapes in transactions, then Duplicate.** Built as planned: `ShapeCreation`, undo
deletes, redo re-runs the recipe and rekeys. Decided along the way:
- **Own centre** means each shape's own centre, so for a multi-shape selection the angle spins every
  shape in place while X and Y move them all together. **Selection centre** is the pivot that turns
  the selection as one rigid piece. It is the centre of the selection's *visual* bounds.
- `ShapeRange.Duplicate` returns copies in z-order rather than selection order (probe 8), so
  `ShapeCreator` duplicates one shape at a time, back to front, which keeps each copy's identity
  certain and the copies' stacking the same as the originals'.
- A step that would leave every copy on its original is refused rather than quietly stacking shapes.
  At most 100 copies. Copies that land wholly off the slide are reported.
- Settings persist like the others: X 20, Y 20, Angle 0, Copies 1, Own centre, Rotate shapes on.

**10. Distribute along a curve.** One **Along curve** button in Distribute, with the curve selected
last. Ovals, Arcs, lines and freeforms are all flattened to a polyline and placed by distance.
- **Oval:** from twelve o'clock clockwise, n even slots, spaced by arc length.
- **Arc:** end to end, with the ellipse rebuilt from the frame and angles (probe 10). **A rotated arc
  is refused.** Its reported frame changes size with the rotation and stops describing the ellipse.
  A flipped arc works.
- **Freeform:** taken exactly as its nodes report. A closed freeform cannot be told from an open one
  (the closing node is not repeated), so it is followed first node to last without the closing
  segment. Use an oval to go all the way round.
- **Line:** its own kind, because lines report no nodes: corner to corner of the frame, flips deciding
  which diagonal.
- The tangent is measured across a short span centred on the point. The segment the point sits on
  was off by half a flattening step, which showed as 89.8° where 90° was meant.
- **Exact (pt)** is the distance between centres along the curve. If it does not fit, the command is
  refused and says how long the curve is.

**The ribbon now needs `autoScale`.** With two more groups the tab no longer fitted a 1680px window,
and Office collapsed Duplicate and AlignPro undo into single drop-downs. The ribbon-click harness
found this: it could not find the buttons. `autoScale="true"` on the groups with large buttons lets
Office shrink those buttons first, so on a narrow window Align shows small buttons and no group
collapses. Distribute later moved to small buttons anyway: its three verbs stack in one column, with
Space by and Exact (pt) in a second. Duplicate opens a gap between the step and its options with a
label of one Braille blank (U+2800). Three simpler ways failed: Office leaves out a label that is only
non-breaking or figure spaces, a disabled spacer button draws a grey square, and a
separator adds a line where only space was wanted.

### Tidy

Planned 2026-09-23, not yet built. **Tidy** finds shapes that are *nearly* aligned or *nearly* evenly
spaced and makes them exact. It fixes a slide that was built by eye: the box 2pt left of its
neighbours, or the row whose gaps are 18, 20 and 21pt. It never guesses a layout that is not already
almost there. Each step below ships on its own with unit tests, and steps 13 and 14 also need COM and
ribbon-click checks and a sample-deck slide.

**Not using Copilot.** Considered and rejected for now:
- Only Office.js add-ins that use the unified manifest can act as a Copilot skill, and that is in
  preview. A VSTO add-in cannot.
- Copilot's view of a PowerPoint slide is its text, not shape positions, so it would only add a
  natural-language way in.
- Language models are unreliable at exact coordinate arithmetic, and this problem is deterministic.
- It would need a Copilot licence for every user and a hosted web add-in. That breaks the no-admin
  install.

The one thing a model could add is telling a sloppy layout from a deliberate one. If that is ever
wanted, an optional call to a model API with a picture of the slide and the snapshots would do it,
without Copilot. It is not part of this plan.

**The shape of it.** `TidySolver` is a separate entry point beside `CurveSolver`, because it has no
verb or reference: it decides for itself what to align. It takes a `TidyRequest` (tolerance, bounds
model, what to fix) and the snapshots. It returns an ordinary `SolveResult` of translations, so it
gets the rest without extra work: one native undo entry, labelled undo, the Measure setting and rigid
groups. It **only moves shapes**. It never resizes or rotates, and its changes carry no angle.

**11. Snap near-alignments.** Each axis is solved on its own, since a translation on one axis
cannot disturb the other.
1. **Find candidates.** For each of the three features on the axis (left, centre and right, or top,
   middle and bottom), sort every shape's value in the requested bounds. Sweep the sorted values
   into clusters, adding a value while it is within the tolerance of the cluster's *lowest* value.
   The spread of a cluster therefore never exceeds the tolerance, so a chain of values each 2pt
   apart cannot grow into one 10pt cluster. A cluster needs two or more shapes.
2. **Lock what is already exact.** A cluster whose spread is within `RectD.Epsilon` is an alignment
   the user already has. Its members are locked on that axis. Tidy must never break an exact
   alignment to make a near one.
3. **Pick a target.** Snap to a *member's* value, never a computed one, so at least one shape stays
   still. Use the member that minimises total movement (the median), and break ties by `ShapeId`.
   A placeholder holds still: moving it overrides the layout's position. A cluster containing
   a placeholder therefore snaps to it, and one containing two placeholders that disagree is
   skipped.
4. **Settle conflicts.** A shape moves at most once per axis. Take clusters by member count, largest
   first, then by spread, tightest first. Accept a cluster only if every member is free or already
   at the target. An accepted cluster's members then count as moved. The rest are skipped, not
   partly applied.

Tolerance defaults to 3pt, accepts 0.5 to 20pt, and a value outside that range is refused. The
algorithm is a single pass. **Tidy must be idempotent:** a second run straight after the first
finds nothing to do. That is a test, and the lock rule is what makes it hold.

**12. Even out near-even spacing.** This runs after step 11, on the moved positions.
- A **row** is three or more shapes in one accepted or locked top/middle/bottom cluster that do not
  overlap horizontally, sorted by X. A **column** is the same on the other axis.
- Measure the gaps and the centre-to-centre pitches. If one of them varies by no more than the
  tolerance, make it exact, holding the outer two shapes still, as Distribute already does.
  If both qualify, use the one that varies less. Leading and trailing edge are not considered, since
  they only differ from centre when sizes differ, and then the gap is what the eye reads.
- Spacing loses to alignment. A row is skipped if evening it would move a shape that step 11 already
  moved or locked on that axis. In a grid, the column alignments usually win and the row spacing is
  left alone.
  **Open question:** spacing the *column targets* evenly, rather than the shapes, would fix both
  at once. Try it only if grids come out looking unfinished.

**13. Tidy on the ribbon.**
- In the **Arrange** group: a **Tidy** button and a **Tolerance (pt)** box that persists like the
  other settings. Recheck that the tab still fits in a 1680px window with `autoScale`.
- Selection only, two or more shapes. Measure is honoured like every other verb: with **Shape frame**,
  a rotated shape is judged by its frame. The README should point to Visual bounds, just as it does
  for Align.
- **Connectors are left out.** They are neither measured nor moved, because they re-route when the
  shapes they join move. `ShapeSnapshot` gains `IsConnector`, which the reader fills from
  `Shape.Connector`. Plain lines take part.
- **It always reports**, because a 2pt fix is invisible. "Tidied 5 shapes: 3 alignments, 1 row", or
  "Nothing to tidy: no shapes are within 3pt of lining up". It uses `OkWithNotice` even on
  success. Connectors left out and clusters skipped over a conflict are both listed.
- The undo label is "Undo tidy".
- A sample-deck slide (18) should include: a jittered 3×3 grid; a near-aligned row with uneven gaps;
  a deliberate 6pt stagger that Tidy must leave alone at 3pt; a rotated shape that is flush only by
  Visual bounds; a connector; and a placeholder that others snap to.

**14. Tidy the whole slide.** With nothing selected, Tidy works on every top-level shape. The reader
already walks `slide.Shapes` for z-order, and would read snapshots in the same walk. This is where a
false positive costs most, because the user did not choose the shapes, so it waits until 13 has had
real use. It may need a lower default tolerance.

**Tests to write first** (`TidyTests.cs`):
- One shape 2pt off a left edge snaps to the other two, and the two do not move.
- With a 3pt tolerance, values of 0, 2, 4 and 6pt do not collapse into one cluster.
- A 6pt stagger is left alone at 3pt.
- An exact centre alignment survives, even when a near left alignment would break it.
- Two clusters that compete for a shape: the larger wins and the other is skipped whole.
- A placeholder holds still, and two disagreeing placeholders skip the cluster.
- Connectors are neither measured nor moved.
- Groups translate rigidly.
- No change carries a rotation.
- A rotated shape is judged by its visual bounds when asked.
- Gap spacing is chosen over centre spacing when it varies less, and the reverse.
- A row whose alignment moves would conflict is skipped.
- A regular grid jittered by up to ±1pt comes back exactly aligned and evenly spaced.
- A second run is a no-op.
- A tolerance of zero or outside 0.5 to 20pt is refused.

### Known limitations

**AlignPro's Undo button can be a step ahead of the document.** Ctrl+Z is safe as of 1.2.0 and
reverses one AlignPro operation, but it does not pop AlignPro's own stack, so after a native undo the
ribbon may still offer to undo what PowerPoint already reversed. Doing so is harmless — changes carry
absolute frames, so it rewrites coordinates the shapes already occupy — but the label misleads.

**Settings are sticky across slides, and that changes what a verb does.** Reference, Measure,
Space by, Exact (pt), Margin (pt), Direction and the Duplicate step persist until you change them. A `Reference` left on **Anchor** makes Grid lay
out inside a single shape's bounds, which packs the whole selection into that shape's footprint — it
looks like the shapes have collapsed into a corner. Grid now falls back to the selection's extent and
says so, but the general trap remains: when a result looks wrong, check Reference and Measure first.

**Don't keep the sample deck in OneDrive.** PowerPoint enables AutoSave for OneDrive-backed files, so
every experiment is written straight back into the fixture. `New-SampleDeck.ps1` therefore defaults to
`sample\` beside the project, which is local and gitignored.

Manual verification: run [`tools/New-SampleDeck.ps1`](../tools/New-SampleDeck.ps1), which writes a saved
17-slide deck to `sample\` and reopens it with a clean undo history. Each slide is
captioned with what to try. The headline check is slide 2 — align left with **Measure = Shape frame**
(what PowerPoint does, and the rotated shape lands wrong) against **Measure = Visual bounds** (flush).
