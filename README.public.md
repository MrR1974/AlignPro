<!--
  The public-facing README from the github.com/MrR1974/AlignPro repository, kept here after that
  repository was deleted. It is the installation and usage guide written for strangers rather than
  the development notes in README.md beside it.

  Preserved so republishing later is a copy rather than a rewrite. If the project goes public again,
  this becomes README.md in that repository, with LICENSE alongside it.
-->
# AlignPro

A PowerPoint add-in for aligning, distributing, sizing and tidying shapes — doing the things the
built-in tools won't.

PowerPoint's own align and distribute have no anchor alignment, no exact spacing, no "make the same
size", no grid tidying — and they align the raw object-model rectangle, which **ignores rotation**. A
rotated shape aligned by PowerPoint is visibly out of line. AlignPro can align on what you actually
see.

---

## Installing

1. Download `AlignPro-<version>.msi` from [Releases](../../releases)
2. Close PowerPoint
3. Double-click the MSI
4. Start PowerPoint — there will be an **AlignPro** tab on the ribbon

No administrator rights: it installs for you alone, under your own user account.

**To remove it:** Settings → Apps → Installed apps → **AlignPro** → Uninstall. It takes everything
with it, including the trust entry described below.

Windows will warn that the publisher is unknown, because the MSI is not signed by a certificate
authority — see [What this does not give you](#what-this-does-not-give-you).

### What the installer does, and what it asks of you

Read this bit — installing an Office add-in involves a trust decision, and you should know which one.

It copies six files to `%LOCALAPPDATA%\AlignPro` — the add-in and the sample deck — and writes one
registry value under `HKEY_CURRENT_USER` so PowerPoint knows where to find them.

Then it **grants trust**, because it has to. VSTO will not load an add-in unless the machine trusts the
certificate that signed its manifest — without it, PowerPoint sets `LoadBehavior` to 2 and the add-in
silently never appears. AlignPro is signed by a self-signed certificate that your machine has never
heard of.

There were three ways to solve that, and this is the narrowest:

| | What you would be trusting |
|---|---|
| **A VSTO inclusion-list entry** ← what the installer does | This one add-in, at this one path, signed by this one key |
| Importing the certificate | Anything ever signed by that certificate |
| A publicly trusted certificate | Anything the publisher ever signs — and ~$1000/year |

So the installer adds a single entry to [VSTO's inclusion list](https://learn.microsoft.com/en-us/visualstudio/vsto/trusting-office-solutions-by-using-inclusion-lists?view=vs-2022),
holding AlignPro's manifest path and the public key that signed it. **Nothing is added to any
certificate store**, and uninstalling revokes the entry again.

It is still a trust decision, just a precise one. If you would rather not make it, build from source
instead — Visual Studio grants trust to what it builds.

### What this does not give you
<a id="what-this-does-not-give-you"></a>

Windows cannot tell you who published AlignPro, so installing shows an "unknown publisher" warning,
and your browser may object to the download. Only a certificate from a certificate authority fixes that, and for a small tool the economics
do not work: since June 2023 the private key
[must live on hardware](https://knowledge.digicert.com/alerts/code-signing-changes-in-2023), and
[Azure Trusted Signing does not support ClickOnce](https://learn.microsoft.com/en-au/answers/questions/1791756/how-to-sign-clickonce-application-manifest-file-wi)
at all. The source is here to read and build if you would rather verify than trust.

### Requirements

Already present on any machine that runs Office:

- Windows with PowerPoint (desktop; not Microsoft 365 for the web)
- .NET Framework 4.8, which ships with Windows 10 1903 and later
- The VSTO runtime, which ships with Office

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

### Arrange

**Grid** tidies a scatter into even rows and columns, keeping each shape near where it already was.
Sizes are never changed, so it's safe on groups. Leave **Columns** blank for a near-square grid.

---

## Two things worth knowing

**Use AlignPro's Undo button, not Ctrl+Z.** PowerPoint groups changes made by an add-in into a single
undo entry that can cover far more than your last action — in testing, one Ctrl+Z removed four slides.
AlignPro's own Undo and Redo reverse exactly one operation at a time and say what they'll reverse.

**The settings are sticky.** Reference, Measure, Space by and Exact (pt) persist until you change them,
including across slides. If a result looks wrong, check those two dropdowns first — a `Reference` left
on **Anchor** makes Grid lay out inside a single shape's bounds, which looks like the shapes have
collapsed into a corner.

---

## Try it

A ten-slide sample deck is **installed alongside the add-in**, at
`%LOCALAPPDATA%\AlignPro\AlignPro-Sample.pptx`. It is also here in the repository at
[`sample/AlignPro-Sample.pptx`](sample/AlignPro-Sample.pptx).

One slide per capability, each captioned with what to try and what should happen. Slide 2 is the one
to start with — align left with **Measure = Shape frame**, undo, then again with **Visual bounds**,
and watch the rotated shape.

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

# Test, build Release and package the MSI
# (needs: dotnet tool install --global wix --version 5.*)
.\tools\New-Release.ps1 -Version 1.0.0
```

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

