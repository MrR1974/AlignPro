# AlignPro

A PowerPoint add-in for aligning, distributing, sizing and tidying shapes, doing the things the
native tools won't: aligning to a designated anchor shape, exact numeric spacing, matching sizes,
grid arrangement, and alignment on what you actually *see* rather than on PowerPoint's own rectangle.

## Why

PowerPoint's align and distribute tools have no anchor/key-object alignment, no exact spacing, no
"make same size", no grid tidying, and they align the raw object-model rectangle — which
[ignores rotation entirely](docs/object-model-findings.md), so rotated shapes end up visually wrong.

## Layout

| Path | What's in it |
|---|---|
| [`src/AlignPro.Geometry/`](src/AlignPro.Geometry/) | The solver and the undo journal. Pure logic, **zero Office references**, `netstandard2.0` |
| [`src/AlignPro.AddIn/`](src/AlignPro.AddIn/) | The VSTO add-in: ribbon, selection adapter, apply pipeline, undo boundary, automation surface. `net48` |
| [`tests/AlignPro.Geometry.Tests/`](tests/AlignPro.Geometry.Tests/) | xUnit suite on `net8.0` |
| [`tools/Probe-ShapeGeometry.ps1`](tools/Probe-ShapeGeometry.ps1) | Measures PowerPoint's object model; doubles as the integration harness |
| [`tools/New-SampleDeck.ps1`](tools/New-SampleDeck.ps1) | Builds a **saved** 10-slide sample deck, one slide per capability |
| [`tools/New-TestDeck.ps1`](tools/New-TestDeck.ps1) | Builds a throwaway scratch deck, never saved |
| [`tools/Test-AlignProEndToEnd.ps1`](tools/Test-AlignProEndToEnd.ps1) | Drives the add-in inside PowerPoint and asserts the results, no clicking |
| [`tools/Probe-UndoGrouping.ps1`](tools/Probe-UndoGrouping.ps1) | How PowerPoint groups undo entries, and what closes a group |
| [`tools/New-DevSigningCertificate.ps1`](tools/New-DevSigningCertificate.ps1) | Creates the machine-local certificate VSTO needs to build |
| [`docs/object-model-findings.md`](docs/object-model-findings.md) | What the probe measured, and what it means for the design |

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
- **We fix PowerPoint's undo rather than replacing it.** PowerPoint coalesces object-model changes
  into [one undo entry and keeps it open until a modifying command arrives from the UI](docs/object-model-findings.md)
  — and a ribbon click is not one. So an AlignPro operation joins whatever entry is already open, and
  one Ctrl+Z can discard an unbounded amount of earlier work. Measured: two commands on a scripted
  deck, one Ctrl+Z, every slide gone.

  Repurposing the built-in Undo would have been the tidy fix, but PowerPoint parses
  `<command idMso="Undo">` and never invokes the callback. Instead `UndoBoundary` closes the group
  itself before each apply — toggle italic on the selection, verify the state actually changed, then
  undo that toggle, which restores the formatting exactly. PowerPoint's entry then holds exactly one
  AlignPro operation, so **Ctrl+Z, the ribbon Undo button and the QAT button are all correct**, with no
  keyboard hook anywhere. The verification-before-undo matters: an unverified undo would revert the
  user's own last edit, which is the very disaster being fixed.

  Because native undo is now correct, AlignPro deliberately has **no undo button of its own** — a
  second stack would disagree with PowerPoint's and misplace shapes. `UndoManager` is retained and
  tested, but the ribbon does not use it.

## Status

| Phase | State |
|---|---|
| 0. Object-model spike | **Done** — six probes plus two follow-up undo experiments |
| 1. Geometry engine + tests | **Done** — 126 tests passing |
| 1b. Undo journal | **Done** — `UndoManager` and `AlignTransaction`, pure and fully tested |
| 2. VSTO shell: ribbon, selection adapter, apply pipeline | **Done** — add-in loads and connects in PowerPoint |
| 3. Verbs wired to the ribbon | **Done** — all twelve verbs, reference/measure/spacing controls |
| 3b. Undo coalescing | **Fixed and verified** — `UndoBoundary` makes native undo per-operation |
| 3c. Automated end-to-end tests | **Done** — 7 checks green via the automation surface |
| 4. Keyboard hook and bindings | Not started |
| 5. ClickOnce packaging and signing | Not started |

### Known limitations

**The undo boundary needs text formatting to exist.** `UndoBoundary` closes PowerPoint's coalescing
group by toggling italic on the selection and undoing it. For a selection with no usable text
formatting — some pictures, lines and placeholders — the state cannot be read, so no boundary is
established and that operation may merge into the previously open undo entry. It degrades to the old
behaviour rather than failing, and it is logged.

**PowerPoint's redo stack is cleared** by an AlignPro operation, as it would be by any edit.

Manual verification: run [`tools/New-SampleDeck.ps1`](tools/New-SampleDeck.ps1), which writes a saved
10-slide deck to your Documents folder and reopens it with a clean undo history. Each slide is
captioned with what to try. The headline check is slide 2 — align left with **Measure = Shape frame**
(what PowerPoint does, and the rotated shape lands wrong) against **Measure = Visual bounds** (flush).
