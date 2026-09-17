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
| `src/AlignPro.AddIn/` | The VSTO add-in: ribbon, selection adapter, undo. `net48` *(not built yet)* |
| [`tests/AlignPro.Geometry.Tests/`](tests/AlignPro.Geometry.Tests/) | xUnit suite on `net8.0` |
| [`tools/Probe-ShapeGeometry.ps1`](tools/Probe-ShapeGeometry.ps1) | Measures PowerPoint's object model; doubles as the integration harness |
| [`tools/New-TestDeck.ps1`](tools/New-TestDeck.ps1) | Builds a scratch deck for exercising the ribbon by hand |
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
- **We keep our own undo journal.** PowerPoint does put object-model changes on its undo stack, but
  coalesces them into [one entry per uninterrupted automation burst](docs/object-model-findings.md),
  and it has no `UndoRecord` equivalent to control that. `UndoManager` gives labelled, per-operation
  undo and redo instead — every `GeometryChange` carries its `OldFrame`, so it is its own undo record.
  Whether we also need to *intercept* Ctrl+Z depends on whether a ribbon click closes the coalescing
  group, which can only be measured once the ribbon exists.

## Status

| Phase | State |
|---|---|
| 0. Object-model spike | **Done** — six probes plus two follow-up undo experiments |
| 1. Geometry engine + tests | **Done** — 107 tests passing |
| 1b. Undo journal | **Done** — `UndoManager` and `AlignTransaction`, pure and fully tested |
| 2. VSTO shell: ribbon, selection adapter, apply pipeline | **Done** — add-in loads and connects in PowerPoint |
| 3. Verbs wired to the ribbon | **Done** — all twelve verbs, both dropdowns, our own undo/redo |
| 3b. Measure undo coalescing against a real ribbon click | Open — see [findings](docs/object-model-findings.md) |
| 4. Keyboard hook and bindings | Not started |
| 5. ClickOnce packaging and signing | Not started |

Manual verification: run [`tools/New-TestDeck.ps1`](tools/New-TestDeck.ps1) and follow the steps it
prints. The headline check is slide 1 — align left with **Measure = Shape frame** (what PowerPoint
does, and the rotated shape lands wrong) against **Measure = Visual bounds** (flush).
