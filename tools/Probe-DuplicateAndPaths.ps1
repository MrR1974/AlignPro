<#
.SYNOPSIS
    Measures the PowerPoint behaviours that Duplicate and Distribute along curve depend on.

.DESCRIPTION
    The planned spike for steps 9 and 10 (docs/development.md, "Next features"). Four questions,
    numbered on from Probe-ShapeGeometry.ps1's six:

      7.  Does Shape.Duplicate stay inside the undo entry StartNewUndoEntry opened, so that one
          Ctrl+Z removes the copy and nothing before it?
      8.  Where does a duplicate land in the z-order, and at what offset from the original? Does
          ShapeRange.Duplicate hand the copies back in the order of the source range?
      9.  Which coordinate space do Shape.Nodes report for a rotated or flipped freeform - and does
          a straight line have nodes at all, and which way do its flips point?
      10. How do the Arc autoshape's two adjustments map to its start and end angles, and are those
          directions from the centre or the ellipse's parameter?

    Probe 10 is answered by the frame rather than the adjustments: PowerPoint resizes an arc's frame
    whenever its angles change, and the new frame gives away both what the angles mean and what the
    frame is.

    Safe by design, as the other probes: works only in a new presentation, never saves, and only
    quits PowerPoint if it started it.

.PARAMETER KeepOpen
    Leave the scratch presentation open at the end.

.PARAMETER OutFile
    Also write the findings to this path as Markdown, for pasting into object-model-findings.md.

.EXAMPLE
    .\Probe-DuplicateAndPaths.ps1 -OutFile $env:TEMP\duplicate-findings.md
#>
[CmdletBinding()]
param(
    [switch] $KeepOpen,
    [string] $OutFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$msoTrue             = -1
$msoFalse            = 0
$msoShapeRectangle   = 1
$msoShapeArc         = 25
$ppLayoutBlank       = 12
$msoEditingAuto      = 0
$msoEditingCorner    = 1
$msoSegmentLine      = 0
$msoSegmentCurve     = 1

$findings = [System.Collections.Generic.List[object]]::new()

function Add-Finding {
    param(
        [Parameter(Mandatory)] [int]    $Number,
        [Parameter(Mandatory)] [string] $Question,
        [Parameter(Mandatory)] [string] $Answer,
        [string[]] $Evidence = @()
    )
    $findings.Add([pscustomobject]@{ Number = $Number; Question = $Question; Answer = $Answer; Evidence = $Evidence })
    Write-Host ''
    Write-Host ("=== Probe {0}: {1}" -f $Number, $Question) -ForegroundColor Cyan
    Write-Host ("    ANSWER: {0}" -f $Answer) -ForegroundColor Green
    foreach ($line in $Evidence) { Write-Host ("            {0}" -f $line) -ForegroundColor DarkGray }
}

function Format-Shape {
    param($Shape, [string] $Label)
    '{0,-26} L={1,8:F2} T={2,8:F2} W={3,8:F2} H={4,8:F2} Rot={5,6:F1} Z={6}' -f `
        $Label, $Shape.Left, $Shape.Top, $Shape.Width, $Shape.Height, $Shape.Rotation, $Shape.ZOrderPosition
}

function Get-Nodes {
    param($Shape)
    $out = @()
    for ($i = 1; $i -le $Shape.Nodes.Count; $i++) {
        $node = $Shape.Nodes.Item($i)
        # A one-by-two SAFEARRAY whose lower bound is not promised to be zero.
        $p = $node.Points
        $row = $p.GetLowerBound(0); $col = $p.GetLowerBound(1)
        $out += [pscustomobject]@{
            X       = [double]$p.GetValue($row, $col)
            Y       = [double]$p.GetValue($row, $col + 1)
            Segment = [int]$node.SegmentType
            Editing = [int]$node.EditingType
        }
    }
    $out
}

function Format-Nodes {
    param($Nodes)
    ($Nodes | ForEach-Object { '({0:F1},{1:F1}){2}' -f $_.X, $_.Y, $(if ($_.Segment -eq $msoSegmentCurve) { 'c' } else { '' }) }) -join ' '
}

# --- Attach to, or start, PowerPoint -------------------------------------------------------------
$wasAlreadyRunning = [bool] (Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue)
$ppt = New-Object -ComObject PowerPoint.Application
$startedPowerPoint = -not $wasAlreadyRunning
$ppt.Visible = $msoTrue
$presentation = $null

try {
    $presentation = $ppt.Presentations.Add()
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $ppt.ActiveWindow.ViewType = 9   # ppViewNormal
    Write-Host ("Scratch presentation created. Slide {0} x {1} pt" -f `
        $presentation.PageSetup.SlideWidth, $presentation.PageSetup.SlideHeight)

    # =============================================================================================
    # Probe 7 - does Duplicate stay inside the entry StartNewUndoEntry opened?
    # =============================================================================================
    $a = $slide.Shapes.AddShape($msoShapeRectangle, 100, 100, 80, 50)
    $a.Name = 'DupA'

    $ppt.StartNewUndoEntry()
    $a.Left = 200                                     # the operation before: must survive
    $ppt.StartNewUndoEntry()
    $copy = $a.Duplicate().Item(1)                    # the operation under test
    $copy.Name = 'DupA-copy'
    $copy.Left = 400
    $copy.Rotation = 30
    $countBefore = $slide.Shapes.Count

    $ppt.CommandBars.ExecuteMso('Undo')
    Start-Sleep -Milliseconds 500

    $countAfter = $slide.Shapes.Count
    $copyGone = $true
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) { if ($slide.Shapes.Item($i).Name -eq 'DupA-copy') { $copyGone = $false } }
    $aLeft = $slide.Shapes.Item('DupA').Left

    $answer7 = if ($copyGone -and [Math]::Abs($aLeft - 200) -lt 0.5) {
        'YES - one undo removed the copy and its edits, and the move before the boundary held.'
    } elseif ($copyGone) {
        "PARTLY - the copy went, but so did the earlier move (DupA.Left = $aLeft)."
    } else {
        'NO - the copy survived one undo. Duplicate is not inside the entry.'
    }
    Add-Finding -Number 7 `
        -Question 'Does Shape.Duplicate stay inside the entry StartNewUndoEntry opened?' `
        -Answer $answer7 `
        -Evidence @(
            "shapes before undo: $countBefore, after: $countAfter",
            "copy gone: $copyGone",
            ("DupA.Left after undo: {0:F1} (moved to 200 before the boundary)" -f $aLeft),
            'Measured over COM. The ribbon-click harness repeats this through a real click.'
        )

    # =============================================================================================
    # Probe 8 - z-order and offset of a duplicate, and ShapeRange.Duplicate's order
    # =============================================================================================
    for ($i = $slide.Shapes.Count; $i -ge 1; $i--) { $slide.Shapes.Item($i).Delete() }
    $p = $slide.Shapes.AddShape($msoShapeRectangle, 100, 300, 60, 40); $p.Name = 'Back'
    $q = $slide.Shapes.AddShape($msoShapeRectangle, 180, 300, 60, 40); $q.Name = 'Middle'
    $r = $slide.Shapes.AddShape($msoShapeRectangle, 260, 300, 60, 40); $r.Name = 'Front'
    $q.Rotation = 20

    $evidence8 = @(
        (Format-Shape -Shape $p -Label 'Back'),
        (Format-Shape -Shape $q -Label 'Middle'),
        (Format-Shape -Shape $r -Label 'Front')
    )

    $mc = $q.Duplicate().Item(1)
    $evidence8 += Format-Shape -Shape $mc -Label 'Middle.Duplicate()'
    $offsetX = $mc.Left - $q.Left
    $offsetY = $mc.Top - $q.Top
    $onTop = $mc.ZOrderPosition -eq $slide.Shapes.Count
    $keptRotation = [Math]::Abs($mc.Rotation - $q.Rotation) -lt 0.01
    $evidence8 += ("offset from original: ({0:F2}, {1:F2}); copy on top: {2}; rotation copied: {3}" -f $offsetX, $offsetY, $onTop, $keptRotation)

    # Writing an absolute position straight after: does the offset interfere?
    $mc.Left = 500; $mc.Top = 200
    $evidence8 += Format-Shape -Shape $mc -Label 'after Left=500 Top=200'
    $absoluteHeld = ([Math]::Abs($mc.Left - 500) -lt 0.01) -and ([Math]::Abs($mc.Top - 200) -lt 0.01)

    # Range order: select Front, Back, Middle (not z-order) and duplicate the range.
    $range = $slide.Shapes.Range([object[]]@('Front', 'Back', 'Middle'))
    $dupRange = $range.Duplicate()
    $sourceNames = @(); for ($i = 1; $i -le $range.Count; $i++) { $sourceNames += $range.Item($i).Name }
    $copyLefts = @(); for ($i = 1; $i -le $dupRange.Count; $i++) { $copyLefts += ('{0:F0}' -f $dupRange.Item($i).Left) }
    $evidence8 += ("Range order: {0}" -f ($sourceNames -join ', '))
    $evidence8 += ("its Duplicate() lefts, in returned order: {0}  (sources at Front=260, Back=100, Middle=180)" -f ($copyLefts -join ', '))

    $answer8 = if ($onTop -and $absoluteHeld) {
        ("ON TOP, OFFSET ({0:F1}, {1:F1}) - a duplicate lands frontmost; writing an absolute position straight after simply overrides the offset." -f $offsetX, $offsetY)
    } else {
        "UNEXPECTED - see evidence (on top: $onTop, absolute held: $absoluteHeld)."
    }
    Add-Finding -Number 8 `
        -Question 'Where does a duplicate land in the z-order, and does its offset need undoing?' `
        -Answer $answer8 -Evidence $evidence8

    # =============================================================================================
    # Probe 9 - which coordinate space do Nodes report?
    # =============================================================================================
    for ($i = $slide.Shapes.Count; $i -ge 1; $i--) { $slide.Shapes.Item($i).Delete() }

    # An L of two straight segments then a curve, with nothing symmetric about it, so any flip or
    # rotation shows up in the numbers.
    $builder = $slide.Shapes.BuildFreeform($msoEditingCorner, 100, 100)
    $builder.AddNodes($msoSegmentLine, $msoEditingAuto, 300, 100)
    $builder.AddNodes($msoSegmentLine, $msoEditingAuto, 300, 200)
    $builder.AddNodes($msoSegmentCurve, $msoEditingCorner, 320, 260, 380, 260, 400, 200)
    $free = $builder.ConvertToShape()
    $free.Fill.Visible = $msoFalse

    $plain = Get-Nodes -Shape $free
    $evidence9 = @(
        (Format-Shape -Shape $free -Label 'freeform'),
        ('nodes as built:   ' + (Format-Nodes $plain)),
        ('segment types:    ' + (($plain | ForEach-Object { $_.Segment }) -join ',') + '   (1 = curve)')
    )

    $free.Rotation = 90
    $rotated = Get-Nodes -Shape $free
    $evidence9 += Format-Shape -Shape $free -Label 'after Rotation=90'
    $evidence9 += 'nodes rotated:    ' + (Format-Nodes $rotated)
    $sameAfterRotation = $true
    for ($i = 0; $i -lt $plain.Count; $i++) {
        if ([Math]::Abs($plain[$i].X - $rotated[$i].X) -gt 0.05 -or [Math]::Abs($plain[$i].Y - $rotated[$i].Y) -gt 0.05) { $sameAfterRotation = $false }
    }
    $free.Rotation = 0

    $free.Flip(0)   # msoFlipHorizontal
    $flipped = Get-Nodes -Shape $free
    $evidence9 += ('HorizontalFlip = {0}' -f $free.HorizontalFlip)
    $evidence9 += 'nodes flipped H:  ' + (Format-Nodes $flipped)
    $sameAfterFlip = $true
    for ($i = 0; $i -lt $plain.Count; $i++) {
        if ([Math]::Abs($plain[$i].X - $flipped[$i].X) -gt 0.05 -or [Math]::Abs($plain[$i].Y - $flipped[$i].Y) -gt 0.05) { $sameAfterFlip = $false }
    }
    $free.Flip(0)

    # A closed freeform: does the node list repeat its first point?
    $b2 = $slide.Shapes.BuildFreeform($msoEditingCorner, 500, 100)
    $b2.AddNodes($msoSegmentLine, $msoEditingAuto, 600, 100)
    $b2.AddNodes($msoSegmentLine, $msoEditingAuto, 600, 200)
    $b2.AddNodes($msoSegmentLine, $msoEditingAuto, 500, 100)
    $closed = $b2.ConvertToShape()
    $evidence9 += 'closed triangle:  ' + (Format-Nodes (Get-Nodes -Shape $closed))

    # A straight line: nodes, and which flips a line drawn up-and-right carries.
    $line = $slide.Shapes.AddLine(100, 400, 300, 350)
    $evidence9 += Format-Shape -Shape $line -Label 'AddLine(100,400 -> 300,350)'
    $evidence9 += ('line Type={0} HorizontalFlip={1} VerticalFlip={2}' -f $line.Type, $line.HorizontalFlip, $line.VerticalFlip)
    try {
        $lineNodes = @(Get-Nodes -Shape $line)
        $evidence9 += if ($lineNodes.Count -eq 0) { 'line nodes: none - Nodes.Count is 0' } else { 'line nodes: ' + (Format-Nodes $lineNodes) }
    }
    catch { $evidence9 += "line nodes: unavailable ($($_.Exception.Message))" }

    $space = if ($sameAfterRotation -and $sameAfterFlip) { 'UNROTATED, UNFLIPPED' }
             elseif ($sameAfterRotation) { 'UNROTATED BUT FLIPPED' }
             elseif ($sameAfterFlip) { 'ROTATED BUT UNFLIPPED' }
             else { 'ROTATED AND FLIPPED' }
    Add-Finding -Number 9 `
        -Question 'Which coordinate space do Shape.Nodes report for a rotated or flipped freeform?' `
        -Answer ("{0} - nodes unchanged by rotation: {1}; unchanged by a horizontal flip: {2}." -f $space, $sameAfterRotation, $sameAfterFlip) `
        -Evidence $evidence9

    # =============================================================================================
    # Probe 10 - Arc adjustments, read off the frame
    # =============================================================================================
    # An arc's adjustments are just two numbers, but PowerPoint resizes the frame whenever they
    # change, and that frame is what gives the geometry away. A 300x100 arc created with the default
    # quarter is an ellipse of radii 300 and 100 centred on the frame's bottom-left corner, (100, 300).
    # Cut it to 0..45 degrees and the frame's height says which reading is right: 45 degrees as a
    # direction from the centre meets that ellipse 94.87pt below the centre, as the ellipse's
    # parameter 70.71pt below.
    for ($i = $slide.Shapes.Count; $i -ge 1; $i--) { $slide.Shapes.Item($i).Delete() }

    function Format-Arc {
        param($Arc, [string] $Label)
        '{0,-22} L={1,8:F2} T={2,8:F2} W={3,8:F2} H={4,8:F2} rot={5,5:F0} flipH={6,2} adj={7:F2},{8:F2}' -f `
            $Label, $Arc.Left, $Arc.Top, $Arc.Width, $Arc.Height, $Arc.Rotation, $Arc.HorizontalFlip,
            $Arc.Adjustments.Item(1), $Arc.Adjustments.Item(2)
    }

    $arc = $slide.Shapes.AddShape($msoShapeArc, 100, 200, 300, 100)
    $evidence10 = @((Format-Arc $arc 'AddShape 300x100'))

    $arc.Adjustments.Item(1) = 0
    $arc.Adjustments.Item(2) = 45
    $evidence10 += Format-Arc $arc 'adjusted to 0..45'
    $evidence10 += '  direction predicts L=100 T=300 W=300 H=94.87; parameter predicts H=70.71'
    $isDirection = [Math]::Abs($arc.Height - 94.87) -lt 0.05
    $boxHoldsCentre = ([Math]::Abs($arc.Left - 100) -lt 0.05) -and ([Math]::Abs($arc.Top - 300) -lt 0.05)

    $arc.Adjustments.Item(1) = -30
    $arc.Adjustments.Item(2) = 200
    $evidence10 += Format-Arc $arc 'adjusted to -30..200'

    $arc.Adjustments.Item(1) = 0
    $arc.Adjustments.Item(2) = 45
    $arc.Rotation = 30
    $evidence10 += Format-Arc $arc 'rotated 30'
    $rotationMovesFrame = [Math]::Abs($arc.Width - 300) -gt 0.05
    $arc.Rotation = 0
    $arc.Flip(0)
    $evidence10 += Format-Arc $arc 'flipped horizontally'
    $arc.Flip(0)
    $arc.Width = 150
    $evidence10 += Format-Arc $arc 'Width = 150'

    $answer10 = if ($isDirection -and $boxHoldsCentre) {
        'ADJ1 = START, ADJ2 = END, degrees clockwise from three o''clock, as DIRECTIONS from the ellipse''s centre. ' +
        'The frame is NOT the ellipse''s box: it is the box of the arc together with the centre. ' +
        $(if ($rotationMovesFrame) { 'Rotating an arc changes its reported frame, so a rotated arc cannot be rebuilt.' } else { '' })
    } else {
        ('UNEXPECTED - height {0:F2}, box holds centre {1}. Inspect by hand.' -f $arc.Height, $boxHoldsCentre)
    }
    Add-Finding -Number 10 `
        -Question "How do the Arc autoshape's adjustments map to start and end angles?" `
        -Answer $answer10 -Evidence $evidence10
}
finally {
    if ($OutFile -and $findings.Count) {
        $md = [System.Text.StringBuilder]::new()
        [void]$md.AppendLine(('Measured by `tools/Probe-DuplicateAndPaths.ps1` on {0}, PowerPoint {1} build {2}.' -f `
            (Get-Date -Format 'yyyy-MM-dd'), $ppt.Version, $ppt.Build))
        [void]$md.AppendLine()
        foreach ($f in ($findings | Sort-Object Number)) {
            [void]$md.AppendLine(('## {0}. {1}' -f $f.Number, $f.Question))
            [void]$md.AppendLine()
            [void]$md.AppendLine(('**{0}**' -f $f.Answer))
            if ($f.Evidence.Count) {
                [void]$md.AppendLine()
                [void]$md.AppendLine('```text')
                foreach ($line in $f.Evidence) { [void]$md.AppendLine($line) }
                [void]$md.AppendLine('```')
            }
            [void]$md.AppendLine()
        }
        $resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutFile)
        $md.ToString() | Set-Content -Path $resolved -Encoding UTF8
        Write-Host ("Findings written to {0}" -f $resolved) -ForegroundColor Green
    }

    if (-not $KeepOpen) {
        if ($null -ne $presentation) {
            $presentation.Saved = $msoTrue
            $presentation.Close()
        }
        if ($startedPowerPoint) { $ppt.Quit() }
    }
}
