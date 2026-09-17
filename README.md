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
| [`tools/Install-AlignPro.ps1`](tools/Install-AlignPro.ps1) | What an end user runs; ships inside the release zip |
| [`tools/Uninstall-AlignPro.ps1`](tools/Uninstall-AlignPro.ps1) | Reverses it, and nothing else |
| [`tools/New-Release.ps1`](tools/New-Release.ps1) | Tests, builds Release, packages the zip |
| [`docs/object-model-findings.md`](docs/object-model-findings.md) | What the probe measured, and what it means for the design |

## Installing

Download the latest zip from Releases, extract it, close PowerPoint and run:

```powershell
.\Install-AlignPro.ps1
```

No administrator rights, no Visual Studio, no compiler and no certificate. The script copies five
files to `%LOCALAPPDATA%\AlignPro` and writes one registry value under `HKEY_CURRENT_USER`; that is
the entire installation. `Uninstall-AlignPro.ps1` reverses exactly that and nothing else.

**AlignPro is not signed by a certificate authority**, so Windows cannot tell you who published it.
That is a deliberate trade-off: publicly trusted code-signing certificates
[can no longer be issued as an exportable `.pfx`](https://knowledge.digicert.com/alerts/code-signing-changes-in-2023)
and must live on a hardware token, and
[Azure Trusted Signing does not support ClickOnce](https://learn.microsoft.com/en-au/answers/questions/1791756/how-to-sign-clickonce-application-manifest-file-wi)
at all. The source is here to be read and built by anyone who would rather verify than trust.

Signing matters only at *build* time - VSTO refuses to emit unsigned manifests - and the certificate
that satisfies it is self-signed and never leaves the machine that built it. Installation does not
involve a certificate: the manifest is registered with the `|vstolocal` suffix, so the add-in is
loaded straight from disk rather than installed as a ClickOnce package, and trust comes from the file
being local. Measured, not assumed: deleting the signing certificate entirely and restarting
PowerPoint leaves the add-in loading exactly as before.

What it needs is already on any machine that runs Office - .NET Framework 4.8 (part of Windows) and
the VSTO runtime (part of Office). The installer checks both and says plainly if either is missing.

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

## Cutting a release

```powershell
.\tools\New-Release.ps1 -Version 1.0.0
```

Runs the geometry tests, builds Release, and writes `dist\AlignPro-<version>.zip` - the five add-in
files, both scripts and a readme, about 62KB. It deliberately publishes nothing; attaching the zip to
a GitHub release stays a separate, deliberate act.

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
- **AlignPro keeps its own undo, because PowerPoint's cannot be trusted.** PowerPoint coalesces
  object-model changes into [one undo entry and keeps it open until a modifying command arrives from
  the UI](docs/object-model-findings.md) — and a ribbon click is not one. So an AlignPro operation
  joins whatever entry is already open, and one Ctrl+Z can discard an unbounded amount of earlier
  work. Measured: two commands on a scripted deck, one Ctrl+Z, every slide gone. **Use the AlignPro
  Undo button, not Ctrl+Z.**

  Two fixes were tried and neither survives contact with a ribbon click. Repurposing the built-in
  Undo: PowerPoint parses `<command idMso="Undo">` and never invokes the callback. Closing the
  coalescing group ourselves (`UndoBoundary`): works from outside PowerPoint, but not from a ribbon
  callback, because PowerPoint defers `ExecuteMso("Undo")` while a command is executing — so the
  undo of our own formatting toggle landed *after* the geometry writes and reverted them, making the
  button silently do nothing. `UndoBoundary` is kept, documented and unused.

## Status

| Phase | State |
|---|---|
| 0. Object-model spike | **Done** — six probes plus two follow-up undo experiments |
| 1. Geometry engine + tests | **Done** — 136 tests passing |
| 1b. Undo journal | **Done** — `UndoManager` and `AlignTransaction`, pure and fully tested |
| 2. VSTO shell: ribbon, selection adapter, apply pipeline | **Done** — add-in loads and connects in PowerPoint |
| 3. Verbs wired to the ribbon | **Done** — all twelve verbs, reference/measure/spacing controls, confirmed by hand against the sample deck |
| 3b. Undo coalescing | **Understood, not solved** — two fixes tried and reverted; AlignPro's own undo is the answer for now |
| 3c. Automated end-to-end tests | **Partly** — geometry is covered; the harness cannot reproduce a ribbon-callback context, which is how a real bug got through |
| 4. Keyboard hook and bindings | Not started |
| 5. Distribution | **Done** - unsigned local install, tested end to end. A signed channel is deferred until there is demand |

### Known limitations

**PowerPoint's own undo is unsafe after an AlignPro command.** Use the AlignPro Undo button. Ctrl+Z
and the Quick Access Toolbar reach PowerPoint's coalesced entry, which may cover far more than your
last action — up to and including everything a script did to build the deck.

**Settings are sticky across slides, and that changes what a verb does.** Reference, Measure,
Space by and Exact (pt) persist until you change them. A `Reference` left on **Anchor** makes Grid lay
out inside a single shape's bounds, which packs the whole selection into that shape's footprint — it
looks like the shapes have collapsed into a corner. Grid now falls back to the selection's extent and
says so, but the general trap remains: when a result looks wrong, check Reference and Measure first.

**Don't keep the sample deck in OneDrive.** PowerPoint enables AutoSave for OneDrive-backed files, so
every experiment is written straight back into the fixture. `New-SampleDeck.ps1` therefore defaults to
`sample\` beside the project, which is local and gitignored.

Manual verification: run [`tools/New-SampleDeck.ps1`](tools/New-SampleDeck.ps1), which writes a saved
10-slide deck to `sample\` and reopens it with a clean undo history. Each slide is
captioned with what to try. The headline check is slide 2 — align left with **Measure = Shape frame**
(what PowerPoint does, and the rotated shape lands wrong) against **Measure = Visual bounds** (flush).
