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
| [`tools/New-SampleDeck.ps1`](../tools/New-SampleDeck.ps1) | Builds the **saved** 13-slide sample deck, one slide per capability |
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
dotnet test tests\AlignPro.Geometry.Tests\AlignPro.Geometry.Tests.csproj   # 161, no PowerPoint
.\tools\Test-AlignProEndToEnd.ps1                                          # 7, PowerPoint via COM
.\tools\Test-RibbonClicks.ps1                                              # 8, real ribbon clicks
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
  `MatchWidth/Height/Both`, `GridArrange`
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
  margin: this one and the slide inset behind `ReferenceTarget.SlideMargins`. They are unrelated

**Ordering is the one verb that lives outside this request.** `ZOrderSolver` is a separate entry
point with its own `OrderVerb`, because the align solver is defined by emitting frame coordinates and
restacking emits no geometry at all. It works in terms of the slide's whole stacking order, back to
front, and rewrites only the slots the selection already occupies — so unselected shapes keep their
layer. `ZOrderPosition` is read-only in the object model, so the applier realises an ordering by
calling `BringToFront` on each shape in turn from the back of the target list forwards; the ordering
that was there before is itself the complete undo instruction.

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
| 3c. Automated end-to-end tests | **Done** — 161 unit tests, 13 COM checks, and 13 real ribbon clicks |
| 4. Keyboard hook and bindings | **Not doing** — a deliberate decision, not an omission. It was originally how Ctrl+Z would be protected, and that need went away; as pure convenience it does not justify a global keyboard hook, the riskiest component in the plan. Revisit if daily use makes the ribbon feel slow |
| 5. Distribution | **Done** - one-line remote install, a zip for the no-terminal route, and an MSI for managed deployment. Installer and uninstaller tested end to end. A signed channel is deferred until there is demand, and is not currently available to this publisher |
| 6. Grow or shrink by the match-size margin | **Planned** — see [Next features](#next-features) |
| 7. Rotation in changes, then Match rotation | **Planned** |
| 8. Object-model spike for duplicate and paths | **Planned** |
| 9. Created shapes in transactions, then Duplicate | **Planned** |
| 10. Distribute on a circle, arc or path | **Planned** |

### Next features

Planned 2026-09-23, in build order. Each step ships on its own with unit tests, COM and ribbon-click
checks, and a sample-deck slide.

**6. Grow or shrink by the margin.** The solver already grows shapes on a negative `SizeMargin`, in
both `Uniform` and `Cascade` (`NegativeMargin_GrowsTheShapesInstead`,
`Cascade_GrowsShapesSelectedAfterTheAnchor`), so this is ribbon-only. Add a **Direction** dropdown
(Shrink / Grow) beside **Apply margin**; the margin box accepts only positive values and the
direction supplies the sign. A negative entry is refused with a pointer to Grow. With Cascade and
Grow, the shape selected first ends **largest** — the screentips and README must say so.

**7. Rotation in changes, then Match rotation.** `GeometryChange` carries frames only, and its
`Inverted()` relies on "no verb changes rotation". Add `OldRotation`/`NewRotation`, swap them in
`Inverted()`, and have `ChangeApplier` write `Rotation`. PowerPoint rotates about the frame's centre,
so a rotation-only change leaves the frame untouched. Match rotation is then a **Rotation** button in
Match size: every shape takes the anchor's angle. Angle only, not flips. Groups are allowed, since
rotation is rigid.

**8. Object-model spike.** Measure before building, and record in `object-model-findings.md`:
- Does `Shape.Duplicate()` / `ShapeRange.Duplicate()` stay inside the entry `StartNewUndoEntry` opened?
- Where do duplicates land in the z-order, and does the default duplicate offset need undoing before
  the absolute position is written?
- Which coordinate space do `Shape.Nodes` report for a rotated or flipped freeform?
- How do the Arc autoshape's adjustments map to start and end angles?

**9. Created shapes in transactions, then Duplicate.** `AlignTransaction` gains a list of created
shapes; undo deletes them. Redo re-runs the duplicate from the originals and rewrites the keys, since
recreated shapes get new ids. `DuplicateSolver` is a separate pure entry point, shaped like
`ZOrderSolver`: selection plus `dx`, `dy`, `angle`, pivot and copies in, per-copy frame and rotation
out, copy *k* receiving the step *k* times. Decided:
- One step may both rotate and translate: rotate about the pivot, then translate. A zero field is inert.
- **Pivot** dropdown: own centre, selection centre, anchor centre, slide centre.
- **Rotate shapes** toggle, default on: copies turn with the step; off, they stay upright and only
  their positions follow the rotation.
- A multi-shape selection is duplicated as a unit, keeping its internal layout; groups stay groups.
- Originals and copies are left selected, ready for the next verb.
- New ribbon group: X, Y, Angle, Copies, Pivot, Rotate shapes, Duplicate. Build translation first,
  then rotation.

**10. Distribute on a circle, arc or path.** The anchor defines the curve; the other shapes' centres
are placed along it in selection order.
- **Ellipse:** a full circuit of n even slots, no overlap at 360°. Spaced by arc length, not angle,
  so non-circular ellipses do not bunch at the ends.
- **Arc autoshape:** end to end, both ends included, angles from its adjustments.
- **Any path:** freeform, line or curve read from `Nodes`, flattened to a polyline and placed by arc
  length. The geometry stays in `AlignPro.Geometry`; the reader only converts.
- **Exact (pt)** is reused as the spacing along the curve.
- The same **Rotate shapes** toggle as Duplicate orients each shape to the tangent.
- Circle and arc first, then arbitrary paths. Duplicate then distribute-on-circle gives radial arrays.

### Known limitations

**AlignPro's Undo button can be a step ahead of the document.** Ctrl+Z is safe as of 1.2.0 and
reverses one AlignPro operation, but it does not pop AlignPro's own stack, so after a native undo the
ribbon may still offer to undo what PowerPoint already reversed. Doing so is harmless — changes carry
absolute frames, so it rewrites coordinates the shapes already occupy — but the label misleads.

**Settings are sticky across slides, and that changes what a verb does.** Reference, Measure,
Space by, Exact (pt) and Margin (pt) persist until you change them. A `Reference` left on **Anchor** makes Grid lay
out inside a single shape's bounds, which packs the whole selection into that shape's footprint — it
looks like the shapes have collapsed into a corner. Grid now falls back to the selection's extent and
says so, but the general trap remains: when a result looks wrong, check Reference and Measure first.

**Don't keep the sample deck in OneDrive.** PowerPoint enables AutoSave for OneDrive-backed files, so
every experiment is written straight back into the fixture. `New-SampleDeck.ps1` therefore defaults to
`sample\` beside the project, which is local and gitignored.

Manual verification: run [`tools/New-SampleDeck.ps1`](../tools/New-SampleDeck.ps1), which writes a saved
13-slide deck to `sample\` and reopens it with a clean undo history. Each slide is
captioned with what to try. The headline check is slide 2 — align left with **Measure = Shape frame**
(what PowerPoint does, and the rotated shape lands wrong) against **Measure = Visual bounds** (flush).
