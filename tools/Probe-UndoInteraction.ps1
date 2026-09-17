<#
.SYNOPSIS
    Confirms what Ctrl+Z actually does after AlignPro-style shape changes.

.DESCRIPTION
    Probe 4 of Probe-ShapeGeometry.ps1 showed that one Ctrl+Z after a batch of object-model work
    removed the whole slide, rather than undoing the shape move that came last. The suspected rule is
    that PowerPoint registers STRUCTURAL changes (inserting a slide) on the undo stack but not SHAPE
    GEOMETRY changes - which would mean AlignPro's operations register nothing at all, and a user's
    reflexive Ctrl+Z reaches straight past them to their own previous edit.

    This script separates the two by interleaving a real human edit with a scripted one. It runs in
    two stages with a manual edit in between.

    Stage 1:  creates a scratch presentation with three shapes, then asks you to drag one by hand.
    Stage 2:  moves a DIFFERENT shape from script, the way AlignPro would, then asks you to press
              Ctrl+Z once and report which change came back.

    Never touches a presentation you already had open: stage 1 creates its own and stage 2 refuses to
    run unless that one is still active.

.PARAMETER Stage
    Setup runs stage 1. Align runs stage 2.

.EXAMPLE
    .\Probe-UndoInteraction.ps1 -Stage Setup
    # ...drag ProbeB by hand in PowerPoint...
    .\Probe-UndoInteraction.ps1 -Stage Align
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Setup', 'Align', 'Granularity')]
    [string] $Stage
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$msoTrue           = -1
$msoShapeRectangle = 1
$ppLayoutBlank     = 12
$markerTitle       = 'AlignPro undo probe'

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $msoTrue

function Format-Shape {
    param($Shape)
    '{0,-8} L={1,8:F2} T={2,8:F2}' -f $Shape.Name, $Shape.Left, $Shape.Top
}

if ($Stage -eq 'Setup') {
    $presentation = $ppt.Presentations.Add()
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)

    # A title shape doubles as the marker that stage 2 looks for, so stage 2 can never run against
    # one of your own decks by accident.
    $title = $slide.Shapes.AddTextbox(1, 40, 30, 500, 40)
    $title.Name = 'ProbeMarker'
    $title.TextFrame2.TextRange.Text = $markerTitle

    foreach ($spec in @(
        @{ Name = 'ProbeA'; X = 100 },
        @{ Name = 'ProbeB'; X = 300 },
        @{ Name = 'ProbeC'; X = 500 })) {
        $shape = $slide.Shapes.AddShape($msoShapeRectangle, $spec.X, 250, 120, 80)
        $shape.Name = $spec.Name
        $shape.TextFrame.TextRange.Text = $spec.Name
    }

    Write-Host ''
    Write-Host 'Stage 1 complete. Starting positions:' -ForegroundColor Cyan
    foreach ($name in @('ProbeA', 'ProbeB', 'ProbeC')) {
        Write-Host ('    ' + (Format-Shape -Shape $slide.Shapes.Item($name))) -ForegroundColor DarkGray
    }

    Write-Host ''
    Write-Host '--- Your turn -----------------------------------------------------------------' -ForegroundColor Yellow
    Write-Host 'In PowerPoint, DRAG ProbeB (the middle rectangle) somewhere obviously different -'
    Write-Host 'a long way down or to the side. Use the mouse, not the keyboard.'
    Write-Host ''
    Write-Host 'That puts one known HUMAN edit on the undo stack. Then run:'
    Write-Host '    .\Probe-UndoInteraction.ps1 -Stage Align' -ForegroundColor White
    Write-Host 'Leave the presentation open and active in the meantime.'
    Write-Host '-------------------------------------------------------------------------------' -ForegroundColor Yellow
    return
}

# --- Stage 2 -------------------------------------------------------------------------------------
$presentation = $ppt.ActivePresentation
$slide = $presentation.Slides.Item(1)

$marker = $null
foreach ($shape in $slide.Shapes) {
    if ($shape.Name -eq 'ProbeMarker') { $marker = $shape; break }
}
if ($null -eq $marker) {
    throw "The active presentation is not the undo probe deck (no ProbeMarker shape). Run -Stage Setup first, and leave that presentation active."
}

if ($Stage -eq 'Granularity') {
    # Stage 2 established that a scripted shape move does get its own undo entry when a human edit
    # bounds it. The remaining question is what happens to TWO operations in a row with no user
    # interaction between them: one undo entry for the pair, or one each? That decides whether
    # AlignPro's own UndoManager is load-bearing or merely convenient.
    $a = $slide.Shapes.Item('ProbeA')
    $c = $slide.Shapes.Item('ProbeC')

    $aBefore = Format-Shape -Shape $a
    $cBefore = Format-Shape -Shape $c

    # Two distinct operations, back to back, in one burst - the way two quick ribbon clicks would be.
    $a.Top = $a.Top + 100
    $c.Top = $c.Top + 150

    Write-Host ''
    Write-Host 'Stage 3 complete. Two scripted moves, no interaction between them.' -ForegroundColor Cyan
    Write-Host ('    ProbeA before: ' + $aBefore) -ForegroundColor DarkGray
    Write-Host ('    ProbeA after:  ' + (Format-Shape -Shape $a)) -ForegroundColor DarkGray
    Write-Host ('    ProbeC before: ' + $cBefore) -ForegroundColor DarkGray
    Write-Host ('    ProbeC after:  ' + (Format-Shape -Shape $c)) -ForegroundColor DarkGray

    Write-Host ''
    Write-Host '--- Your turn -----------------------------------------------------------------' -ForegroundColor Yellow
    Write-Host 'Click an empty part of the slide, then press Ctrl+Z ONCE. Which happened?'
    Write-Host ''
    Write-Host '  (a) only ProbeC reverted, ProbeA stayed moved'
    Write-Host '      -> each operation gets its own undo entry. Native Ctrl+Z is well behaved and'
    Write-Host '         our UndoManager is a convenience, not a safety net.'
    Write-Host ''
    Write-Host '  (b) BOTH ProbeA and ProbeC reverted'
    Write-Host '      -> consecutive operations coalesce into one entry. One Ctrl+Z silently undoes'
    Write-Host '         more than the last action, so AlignPro should intercept it.'
    Write-Host ''
    Write-Host '  (c) anything else - say what you saw'
    Write-Host ''
    Write-Host 'Then close this deck WITHOUT saving; the probe deck has nothing worth keeping.'
    Write-Host '-------------------------------------------------------------------------------' -ForegroundColor Yellow
    return
}

$b = $slide.Shapes.Item('ProbeB')
$c = $slide.Shapes.Item('ProbeC')

$bBefore = Format-Shape -Shape $b
$cBefore = Format-Shape -Shape $c

# This is the AlignPro-shaped change: a pure shape geometry write, nothing structural.
$c.Left = $c.Left - 200
$c.Top = $c.Top + 60

Write-Host ''
Write-Host 'Stage 2 complete.' -ForegroundColor Cyan
Write-Host ('    ProbeB, where your drag left it:  ' + $bBefore) -ForegroundColor DarkGray
Write-Host ('    ProbeC before the scripted move:  ' + $cBefore) -ForegroundColor DarkGray
Write-Host ('    ProbeC after the scripted move:   ' + (Format-Shape -Shape $c)) -ForegroundColor DarkGray

Write-Host ''
Write-Host '--- Your turn -----------------------------------------------------------------' -ForegroundColor Yellow
Write-Host 'Click once on an empty part of the slide, then press Ctrl+Z ONCE. Which happened?'
Write-Host ''
Write-Host '  (a) ProbeC jumped back to its pre-script position'
Write-Host '      -> shape geometry changes DO register. Ctrl+Z undoes our work, one property at a time.'
Write-Host ''
Write-Host '  (b) ProbeB jumped back to where it was BEFORE your drag, ProbeC stayed put'
Write-Host '      -> our changes register NOTHING. Ctrl+Z silently reaches past them to your own'
Write-Host '         last edit. This is the suspected answer, and the reason AlignPro has to'
Write-Host '         intercept Ctrl+Z itself.'
Write-Host ''
Write-Host '  (c) the slide or the shapes vanished'
Write-Host '      -> the stack only holds the structural change from stage 1. Worse than (b).'
Write-Host ''
Write-Host '  (d) nothing moved at all'
Write-Host '      -> the stack is empty; Ctrl+Z is inert after automation.'
Write-Host ''
Write-Host 'Record the answer in docs/object-model-findings.md, then close this deck WITHOUT saving.'
Write-Host '-------------------------------------------------------------------------------' -ForegroundColor Yellow
