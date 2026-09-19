# AlignPro

A PowerPoint add-in for aligning, distributing, sizing and tidying shapes — doing the things the
built-in tools won't.

PowerPoint's own align and distribute have no anchor alignment, no exact spacing, no "make the same
size", no grid tidying — and they align the raw object-model rectangle, which **ignores rotation**. A
rotated shape aligned by PowerPoint is visibly out of line. AlignPro can align on what you actually
see.

---

## Installing

Close PowerPoint, then paste this into PowerShell:

```powershell
irm https://raw.githubusercontent.com/MrR1974/AlignPro/main/install.ps1 | iex
```

Start PowerPoint — there will be an **AlignPro** tab on the ribbon.

No administrator rights: it installs for you alone, under your own user account. **To remove it:**
Settings → Apps → Installed apps → **AlignPro** → Uninstall.

### Why a command rather than a download

Because a downloaded installer does not work well here, and it is worth being straight about why.

Windows checks downloaded programs against SmartScreen's reputation database, and an unsigned file has
no reputation to check. Microsoft's own guidance is explicit that for an unsigned file, reputation
"must build for each new version of your files, starting with zero" — so an unsigned installer is
warned about on *every* release, no matter how many people have installed the previous one, and a
self-signed certificate counts as no signature at all. That is not a warning AlignPro can grow out of.

The command above sidesteps it rather than arguing with it: nothing is downloaded by the browser, so
nothing goes through that check. It fetches this repository's `install.ps1` over HTTPS, and that script
downloads the release, verifies it against the SHA-256 published beside it, and installs it. You can
[read the script](install.ps1) before running it — it is the same file the command fetches.

### If you would rather not paste a command

Download `AlignPro-<version>.zip` from [Releases](../../releases), then:

1. **Right-click the zip → Properties → tick Unblock → OK.** Do this *before* extracting
2. Extract it
3. Close PowerPoint and double-click **Install AlignPro.cmd**

Step 1 matters. Windows marks files that came from the internet, Explorer's extractor copies that mark
onto everything inside, and Windows then refuses to run the installer script — without ever explaining
why. Unblocking the zip first clears the mark for everything it contains.

For managed deployment there is also an MSI on the release, which Intune and Group Policy can install
per-user without any of this applying.

### What the installer does, and what it asks of you

Read this bit — installing an Office add-in involves a trust decision, and you should know which one.

It copies the add-in, the sample deck and its own uninstaller to `%LOCALAPPDATA%\AlignPro`, and writes
a few registry values under `HKEY_CURRENT_USER`: one telling PowerPoint where to find the add-in, and
one listing AlignPro in Add/Remove Programs so you can uninstall it like anything else.

Then it **grants trust**, because it has to. VSTO will not load an add-in unless the machine trusts the
certificate that signed its manifest — without it, PowerPoint sets `LoadBehavior` to 2 and the add-in
silently never appears. AlignPro is signed by a self-signed certificate that your machine has never
heard of.

There were three ways to solve that, and this is the narrowest:

| | What you would be trusting |
|---|---|
| **A VSTO inclusion-list entry** ← what the installer does | This one add-in, at this one path, signed by this one key |
| Importing the certificate | Anything ever signed by that certificate |
| A publicly trusted certificate | Anything the publisher ever signs |

So the installer adds a single entry to [VSTO's inclusion list](https://learn.microsoft.com/en-us/visualstudio/vsto/trusting-office-solutions-by-using-inclusion-lists?view=vs-2022),
holding AlignPro's manifest path and the public key that signed it. **Nothing is added to any
certificate store**, and uninstalling revokes the entry again.

It is still a trust decision, just a precise one. If you would rather not make it, build from source
instead — Visual Studio grants trust to what it builds.

### What this does not give you

Windows cannot tell you who published AlignPro. Only a certificate from a certificate authority fixes
that, and the economics do not work for a tool like this: since June 2023 a traditional code-signing
certificate's private key
[must live on hardware](https://knowledge.digicert.com/alerts/code-signing-changes-in-2023). Microsoft's
own [Artifact Signing](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/code-signing-options)
service removed both the token and most of the cost — it starts at about $10 a month — but individual
identity validation is limited to the United States and Canada, and the organisation route wants a
legal entity with three years of trading history. Neither is available here.

It would not solve the whole problem anyway: [Azure Trusted Signing does not support ClickOnce](https://learn.microsoft.com/en-au/answers/questions/1791756/how-to-sign-clickonce-application-manifest-file-wi),
which is how VSTO manifests are signed, so the trust decision described above would remain either way.

The source is here to read and build if you would rather verify than trust.

### Requirements

Already present on any machine that runs Office:

- Windows with PowerPoint (desktop; not Microsoft 365 for the web)
- .NET Framework 4.8, which ships with Windows 10 1903 and later
- The VSTO runtime, which ships with Office

The installer checks all three. If it cannot find the VSTO runtime it says so and installs anyway
rather than refusing: the runtime registers itself differently on different machines, and a check
that is sometimes wrong should not be allowed to block a machine that is fine. If the AlignPro tab
then does not appear, [install the runtime](https://aka.ms/VSTORuntimeDownload) and run the
installer again.
---

## Using it

Select some shapes, then use the **AlignPro** tab. Two dropdowns carry the state the six align buttons
are crossed with, so a small set of buttons covers a large matrix.

### Align to

| Reference | What the verbs measure against |
|---|---|
| **Anchor** | The shape you selected **last**. Select the others first, then Ctrl+click the one to align to |
| **Selection bounds** | The outline of everything selected — what PowerPoint does natively |
| **Slide** | The slide edges |
| **Slide margins** | The slide inset by the **Margin (pt)** box |
| **Content placeholder** | The body placeholder from the slide's layout |

### Measure

| Model | Which rectangle of each shape is aligned |
|---|---|
| **Shape frame** | PowerPoint's own rectangle. Ignores rotation — this is the native behaviour |
| **Visual bounds** | What you actually see. Use this whenever rotated shapes are involved |
| **Text bounds** | The text inside the shape rather than the shape itself. Lines up captions whose boxes differ |

### Align

Six buttons — **Left, Centre, Right, Top, Middle, Bottom** — applied to whichever Reference and
Measure are selected.

### Distribute

**Space by** decides what actually gets spaced evenly. All four agree when the shapes are the same
size, and diverge the moment they differ:

| Mode | Spaces |
|---|---|
| **Leading edges** | Left edges horizontally, top edges vertically, at a constant pitch |
| **Centres** | Centre to centre |
| **Trailing edges** | Right edges horizontally, bottom edges vertically |
| **Space between** | The visible gaps — the only mode that equalises what you see between shapes |

**Exact (pt)** sets the spacing precisely; leave it blank to even out whatever is already there. In the
three edge and centre modes the number is a *pitch*; in Space between it is the *gap*. Press **Enter**
to commit the value.

### Match size

**Width**, **Height** or **Both**, matched to the anchor — select the others first, then the shape to
match **last**. **From centre** holds each shape's centre while resizing instead of its top-left
corner.

Groups are skipped, and AlignPro tells you why: resizing a group rescales the gaps between its
children, which silently distorts a diagram. A group makes a perfectly good *anchor*, though — it's
only measured, never resized.

**Margin (pt)** makes the shapes a set amount *smaller* than the anchor rather than the same size. It
is measured **per side**, so a margin of 10 leaves a 10pt border showing all the way round and takes 20
off each dimension. Negative numbers make the shapes larger instead. (This is the match-size margin —
not the slide margin in **Align to**, which is a separate setting that happens to share the name.)

**Apply margin** decides how far that goes:

| Setting | What it does |
|---|---|
| **Off** | Ignore the margin and match the anchor exactly. |
| **All the same** | Every shape ends one margin inside the anchor, so they all come out the same size. |
| **Cascade** | The margin accumulates along the order you selected in, so the shapes tier — and the shape you selected **first** ends smallest. |

For concentric rings, turn **From centre** on, use **Cascade**, and select the outermost shape last.
Slide 9 of the sample deck walks through it.

### Order

**Stack** restacks the selection in the order you selected it: the shape you selected **first** ends on
top, the one you selected last at the bottom. **Reverse** is the mirror image. Either replaces a run of
Bring to front and Send to back with one click, for any number of shapes.

The reordering happens **in place**. The selected shapes are redistributed across the layers they
already occupied between them, so anything you did *not* select stays on exactly the layer it was on —
unlike Bring to front, which hauls the whole selection over the top of everything else.

Select the shapes that sit directly on the slide, not shapes inside a group: a grouped shape is
stacked within its group rather than against the slide, so AlignPro refuses rather than guess.

### Arrange

**Grid** tidies a scatter into even rows and columns, keeping each shape near where it already was.
Sizes are never changed, so it's safe on groups. Leave **Columns** blank for a near-square grid.

---

## Two things worth knowing

**Use AlignPro's Undo button, not Ctrl+Z.** PowerPoint groups changes made by an add-in into a single
undo entry that can cover far more than your last action — in testing, one Ctrl+Z removed four slides.
AlignPro's own Undo and Redo reverse exactly one operation at a time and say what they'll reverse.

**The settings are sticky.** Reference, Measure, Space by, Exact (pt) and Margin (pt) persist until you change them,
including across slides. If a result looks wrong, check those two dropdowns first — a `Reference` left
on **Anchor** makes Grid lay out inside a single shape's bounds, which looks like the shapes have
collapsed into a corner.

---

## Try it

A thirteen-slide sample deck is **installed alongside the add-in**, at
`%LOCALAPPDATA%\AlignPro\AlignPro-Sample.pptx`. It is also here in the repository at
[`sample/AlignPro-Sample.pptx`](sample/AlignPro-Sample.pptx).

One slide per capability, each captioned with what to try and what should happen. Slide 2 is the one
to start with — align left with **Measure = Shape frame**, undo, then again with **Visual bounds**,
and watch the rotated shape. Slides 7 to 9 cover ordering and the match-size margin, ending with a
concentric-rings slide that uses both together.

---
## Building from source
<a id="building-from-source"></a>

Needs Visual Studio 2022 with the **Office/SharePoint development** workload — VSTO projects cannot be
built by the .NET CLI.

```powershell
.\tools\New-DevSigningCertificate.ps1     # once: VSTO will not build unsigned manifests
```

That creates a self-signed certificate used only to satisfy the build. It never leaves your machine
and plays no part in installation.

```powershell
# The geometry engine and its tests - no Visual Studio, no PowerPoint
dotnet test tests\AlignPro.Geometry.Tests\AlignPro.Geometry.Tests.csproj

# Everything, including the add-in
& "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" `
    AlignPro.sln /p:Configuration=Debug /p:VisualStudioVersion=17.0

# Test, build Release, and package the zip, its checksum and the MSI
# (the MSI needs: dotnet tool install --global wix --version 5.*; omit it with -SkipInstaller)
.\tools\New-Release.ps1 -Version 1.0.0
```

That writes `dist\AlignPro-<version>.zip`, `dist\AlignPro-<version>.zip.sha256` and
`dist\AlignPro-<version>.msi`. The zip and the checksum must both be attached to the GitHub release
and keep those names: `install.ps1` looks them up by name and refuses to install without the checksum.

[`docs/development.md`](docs/development.md) is the companion to this file for anyone changing
AlignPro: how the pieces fit, how the distribution route was chosen, and what was measured rather
than assumed.

### How it's put together

`AlignPro.Geometry` targets `netstandard2.0` and has **no Office references at all** — it is pure
geometry, so the whole engine and its test suite build and run on the .NET CLI with no Visual Studio
and no PowerPoint. `AlignPro.AddIn` is a thin `net48` adapter: it reads the selection into immutable
snapshots, hands them to the solver, and writes the results back.

The solver reasons in whichever bounds space you asked for but always emits *frame* coordinates,
because the frame is the only thing PowerPoint lets you write. Translation is rigid, so a delta
measured in visual or text space is the same delta in frame space — which is how align and distribute
get rotation- and text-awareness essentially for free.

[`docs/object-model-findings.md`](docs/object-model-findings.md) records what was measured about
PowerPoint's object model — rotation, selection order, group behaviour, undo coalescing — and why the
code is shaped the way it is. Worth reading before changing anything in that area; several of those
behaviours are not what the documentation implies.

### Testing

Three layers, because each catches what the others cannot:

```powershell
dotnet test tests\AlignPro.Geometry.Tests\AlignPro.Geometry.Tests.csproj   # 136, no PowerPoint
.\tools\Test-AlignProEndToEnd.ps1                                          # 7, PowerPoint via COM
.\tools\Test-RibbonClicks.ps1                                              # 8, real ribbon clicks
```

The third exists because the second is blind to a whole class of bug. It drives the add-in over
cross-process COM, where no Office command is in flight — which is not the context a ribbon callback
runs in. A bug where Align Left silently did nothing when clicked passed all seven of those checks.
`Test-RibbonClicks.ps1` clicks the actual ribbon through UI Automation, so Office dispatches the
command exactly as it would for a person.

---

## Licence

MIT — see [LICENSE](LICENSE).

