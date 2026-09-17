<#
.SYNOPSIS
    Measures PowerPoint object-model behaviours that the AlignPro geometry engine depends on.

.DESCRIPTION
    Six behaviours decide how AlignPro.Geometry must be designed, and the documentation either
    conflicts or is silent on all of them. This script answers them empirically against the
    installed PowerPoint, in a brand-new presentation, and reports what it found.

    Safe by design:
      - Works exclusively in a new presentation. Never opens, modifies or saves anything of yours.
      - Only quits PowerPoint if it was not already running before the script started.
      - Never saves. The scratch presentation is closed without saving (unless -KeepOpen).

    Probe 4 (undo) cannot be fully automated - PowerPoint's undo stack is only reachable from the
    UI. The script sets up the condition and prints what to check by hand.

.PARAMETER KeepOpen
    Leave the scratch presentation open at the end, for poking at by hand.

.PARAMETER OutFile
    Also write the findings to this path as Markdown.

.EXAMPLE
    .\Probe-ShapeGeometry.ps1 -OutFile ..\docs\object-model-findings.md
#>
[CmdletBinding()]
param(
    [switch] $KeepOpen,
    [string] $OutFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Office enum values, spelled out so the script carries no interop dependency ----------------
$msoTrue                       = -1
$msoFalse                      = 0
$msoShapeRectangle             = 1
$ppLayoutBlank                 = 12
$msoTextOrientationHorizontal  = 1
$msoAutoSizeNone               = 0
$msoAutoSizeShapeToFitText     = 1

$findings = [System.Collections.Generic.List[object]]::new()

function Add-Finding {
    param(
        [Parameter(Mandatory)] [int]    $Number,
        [Parameter(Mandatory)] [string] $Question,
        [Parameter(Mandatory)] [string] $Answer,
        [string[]] $Evidence = @()
    )
    $findings.Add([pscustomobject]@{
        Number   = $Number
        Question = $Question
        Answer   = $Answer
        Evidence = $Evidence
    })
    Write-Host ''
    Write-Host ("=== Probe {0}: {1}" -f $Number, $Question) -ForegroundColor Cyan
    Write-Host ("    ANSWER: {0}" -f $Answer) -ForegroundColor Green
    foreach ($line in $Evidence) { Write-Host ("            {0}" -f $line) -ForegroundColor DarkGray }
}

function Format-Rect {
    param($Shape, [string] $Label)
    '{0,-24} L={1,8:F2}  T={2,8:F2}  W={3,8:F2}  H={4,8:F2}  Rot={5,6:F1}' -f `
        $Label, $Shape.Left, $Shape.Top, $Shape.Width, $Shape.Height, $Shape.Rotation
}

# --- Attach to, or start, PowerPoint -------------------------------------------------------------
# Marshal.GetActiveObject does not exist in .NET Core / PowerShell 7, so detect a running instance
# by process instead. PowerPoint.Application is effectively a singleton: New-Object attaches to the
# running instance when there is one.
$wasAlreadyRunning = [bool] (Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue)
$ppt = New-Object -ComObject PowerPoint.Application
$startedPowerPoint = -not $wasAlreadyRunning
Write-Host $(if ($wasAlreadyRunning) {
    'Attached to the PowerPoint instance already running. It will be left running.'
} else {
    'Started a new PowerPoint instance. It will be closed at the end.'
}) -ForegroundColor Yellow

$ppt.Visible = $msoTrue
$presentation = $null

try {
    $presentation = $ppt.Presentations.Add()
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    Write-Host ("Scratch presentation created. Slide size: {0} x {1} pt" -f `
        $presentation.PageSetup.SlideWidth, $presentation.PageSetup.SlideHeight)

    # =============================================================================================
    # Probe 1 - Do Left/Top/Width/Height describe the unrotated frame or the rotated visual bbox?
    # =============================================================================================
    # This is the one that matters most. If the OM already returns visual bounds, rotation-aware
    # alignment is nearly free. If it returns the unrotated frame, AlignPro must compute visual
    # bounds itself and convert deltas back into frame space.
    $r = $slide.Shapes.AddShape($msoShapeRectangle, 100, 100, 100, 50)
    $before = Format-Rect -Shape $r -Label 'before rotation'
    $centreBeforeX = $r.Left + $r.Width / 2
    $centreBeforeY = $r.Top + $r.Height / 2

    $r.Rotation = 45
    $after = Format-Rect -Shape $r -Label 'after Rotation=45'
    $centreAfterX = $r.Left + $r.Width / 2
    $centreAfterY = $r.Top + $r.Height / 2

    # A 100x50 rect rotated 45 degrees has a visual bounding box of 106.07 x 106.07.
    $theta = 45 * [Math]::PI / 180
    $expectedVisualW = 100 * [Math]::Abs([Math]::Cos($theta)) + 50 * [Math]::Abs([Math]::Sin($theta))
    $expectedVisualH = 100 * [Math]::Abs([Math]::Sin($theta)) + 50 * [Math]::Abs([Math]::Cos($theta))

    $centreHeld = ([Math]::Abs($centreAfterX - $centreBeforeX) -lt 0.01) -and
                  ([Math]::Abs($centreAfterY - $centreBeforeY) -lt 0.01)
    $widthChanged = [Math]::Abs($r.Width - 100) -gt 0.01

    $answer1 = if ($widthChanged) {
        'VISUAL BBOX - the OM reports the rotated bounding box. Rotation-aware bounds is nearly free.'
    } else {
        'UNROTATED FRAME - the OM ignores rotation. AlignPro must compute visual bounds itself.'
    }

    Add-Finding -Number 1 `
        -Question 'Does Shape.Left/Top/Width/Height return the unrotated frame or the rotated visual bbox?' `
        -Answer $answer1 `
        -Evidence @(
            $before
            $after
            ('centre before ({0:F2}, {1:F2})  ->  centre after ({2:F2}, {3:F2})' -f $centreBeforeX, $centreBeforeY, $centreAfterX, $centreAfterY)
            ('a 100x50 rect at 45 deg has a visual bbox of {0:F2} x {1:F2}' -f $expectedVisualW, $expectedVisualH)
            ('centre invariant under rotation: {0}' -f $centreHeld)
        )
    $r.Delete()

    # =============================================================================================
    # Probe 2 - Does Selection.ShapeRange preserve selection order, or fall back to z-order?
    # =============================================================================================
    # "Align to anchor (last selected wins)" depends on this. Shapes are added A, B, C so that
    # z-order is A, B, C, then selected in the order C, A, B. If the range comes back C, A, B the
    # selection order survived; if it comes back A, B, C we only have z-order and need an explicit
    # "Pin anchor" button instead.
    $a = $slide.Shapes.AddShape($msoShapeRectangle,  20, 300, 60, 40); $a.Name = 'ProbeA'
    $b = $slide.Shapes.AddShape($msoShapeRectangle, 120, 300, 60, 40); $b.Name = 'ProbeB'
    $c = $slide.Shapes.AddShape($msoShapeRectangle, 220, 300, 60, 40); $c.Name = 'ProbeC'

    $c.Select($msoTrue)     # replace selection
    $a.Select($msoFalse)    # add to selection
    $b.Select($msoFalse)    # add to selection

    $selectedOrder = @()
    $selRange = $ppt.ActiveWindow.Selection.ShapeRange
    for ($i = 1; $i -le $selRange.Count; $i++) { $selectedOrder += $selRange.Item($i).Name }
    $orderString = $selectedOrder -join ', '

    $answer2 = switch ($orderString) {
        'ProbeC, ProbeA, ProbeB' { 'SELECTION ORDER PRESERVED - "last selected is the anchor" is viable.' }
        'ProbeA, ProbeB, ProbeC' { 'Z-ORDER ONLY - selection order is lost. Ship an explicit "Pin anchor" button.' }
        default                  { "UNEXPECTED order '$orderString' - inspect by hand before relying on either rule." }
    }

    Add-Finding -Number 2 `
        -Question 'Does Selection.ShapeRange preserve selection order?' `
        -Answer $answer2 `
        -Evidence @(
            'added in z-order:  ProbeA, ProbeB, ProbeC'
            'selected in order: ProbeC, ProbeA, ProbeB'
            "ShapeRange returned: $orderString"
            'NOTE: this selects programmatically. Confirm by clicking the three shapes by hand too -'
            'the UI and the OM do not have to agree.'
        )

    # =============================================================================================
    # Probe 3 - Group semantics: does a group behave as one rigid object?
    # =============================================================================================
    # Requirement: a group is always treated as a single object, and its internal spacing must
    # survive alignment. Translation should move children rigidly. The open question is what a
    # RESIZE does - if setting group.Width scales the children and the gaps between them, then
    # "match size to anchor" silently rescales internal spacing, and we must decide whether to
    # allow that or refuse resize verbs on groups.
    $group = $ppt.ActiveWindow.Selection.ShapeRange.Group()
    $group.Name = 'ProbeGroup'

    function Get-ChildGeometry {
        param($Group)
        $out = @()
        for ($i = 1; $i -le $Group.GroupItems.Count; $i++) {
            $item = $Group.GroupItems.Item($i)
            $out += [pscustomobject]@{ Name = $item.Name; Left = $item.Left; Top = $item.Top; Width = $item.Width }
        }
        $out
    }

    $childrenStart = Get-ChildGeometry -Group $group
    # Internal spacing measured left-edge to left-edge, in the group's own child order.
    $gapsStart = @()
    for ($i = 1; $i -lt $childrenStart.Count; $i++) {
        $gapsStart += $childrenStart[$i].Left - $childrenStart[$i - 1].Left
    }

    $groupEvidence = @(Format-Rect -Shape $group -Label 'group before translate')
    foreach ($ch in $childrenStart) {
        $groupEvidence += '    child {0,-8} L={1,8:F2} T={2,8:F2} W={3,8:F2}' -f $ch.Name, $ch.Left, $ch.Top, $ch.Width
    }
    $groupEvidence += ('    left-edge gaps: {0}' -f (($gapsStart | ForEach-Object { '{0:F2}' -f $_ }) -join ', '))

    # --- translate the group -----------------------------------------------------------------
    $group.Left = $group.Left + 60
    $group.Top  = $group.Top  - 30
    $childrenMoved = Get-ChildGeometry -Group $group
    $gapsMoved = @()
    for ($i = 1; $i -lt $childrenMoved.Count; $i++) {
        $gapsMoved += $childrenMoved[$i].Left - $childrenMoved[$i - 1].Left
    }
    $translatePreservesSpacing = $true
    for ($i = 0; $i -lt $gapsStart.Count; $i++) {
        if ([Math]::Abs($gapsMoved[$i] - $gapsStart[$i]) -gt 0.01) { $translatePreservesSpacing = $false }
    }
    $groupEvidence += Format-Rect -Shape $group -Label 'group after translate'
    $groupEvidence += ('    left-edge gaps: {0}  (preserved: {1})' -f `
        (($gapsMoved | ForEach-Object { '{0:F2}' -f $_ }) -join ', '), $translatePreservesSpacing)

    # --- resize the group --------------------------------------------------------------------
    $widthBeforeResize = $group.Width
    $group.Width = $widthBeforeResize * 1.5
    $childrenResized = Get-ChildGeometry -Group $group
    $gapsResized = @()
    for ($i = 1; $i -lt $childrenResized.Count; $i++) {
        $gapsResized += $childrenResized[$i].Left - $childrenResized[$i - 1].Left
    }
    $resizeScalesSpacing = $false
    for ($i = 0; $i -lt $gapsStart.Count; $i++) {
        if ([Math]::Abs($gapsResized[$i] - $gapsStart[$i]) -gt 0.01) { $resizeScalesSpacing = $true }
    }
    $groupEvidence += Format-Rect -Shape $group -Label 'group after 1.5x width'
    foreach ($ch in $childrenResized) {
        $groupEvidence += '    child {0,-8} L={1,8:F2} T={2,8:F2} W={3,8:F2}' -f $ch.Name, $ch.Left, $ch.Top, $ch.Width
    }
    $groupEvidence += ('    left-edge gaps: {0}  (scaled: {1})' -f `
        (($gapsResized | ForEach-Object { '{0:F2}' -f $_ }) -join ', '), $resizeScalesSpacing)

    $answer3 = if ($translatePreservesSpacing -and $resizeScalesSpacing) {
        'RIGID ON TRANSLATE, PROPORTIONAL ON RESIZE - align/distribute are safe for groups as-is. ' +
        'Match-size rescales internal spacing, so it needs an explicit decision.'
    } elseif ($translatePreservesSpacing) {
        'RIGID ON TRANSLATE, resize left spacing untouched - groups are safe for every verb.'
    } else {
        'TRANSLATE DID NOT PRESERVE INTERNAL SPACING - unexpected. Investigate before trusting groups.'
    }

    Add-Finding -Number 3 `
        -Question 'Does a group behave as one rigid object, preserving internal spacing?' `
        -Answer $answer3 `
        -Evidence $groupEvidence
    $group.Delete()

    # =============================================================================================
    # Probe 5 - Are TextRange2 bound properties usable as a text-alignment bounds model?
    # =============================================================================================
    # Run before probe 4 so the undo instructions are the last thing left on screen.
    $tb = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, 300, 100, 200, 80)
    $tb.Name = 'ProbeText'
    $tb.TextFrame2.TextRange.Text = 'Align me'
    $textEvidence = @(Format-Rect -Shape $tb -Label 'textbox frame')

    foreach ($autoSize in @($msoAutoSizeNone, $msoAutoSizeShapeToFitText)) {
        foreach ($wrap in @($msoTrue, $msoFalse)) {
            $tb.TextFrame2.AutoSize = $autoSize
            $tb.TextFrame2.WordWrap = $wrap
            $tr = $tb.TextFrame2.TextRange
            $textEvidence += 'AutoSize={0} WordWrap={1,3}  BoundL={2,8:F2} BoundT={3,8:F2} BoundW={4,8:F2} BoundH={5,8:F2}' -f `
                $autoSize, $wrap, $tr.BoundLeft, $tr.BoundTop, $tr.BoundWidth, $tr.BoundHeight
        }
    }

    $tr = $tb.TextFrame2.TextRange
    $answer5 = if (($tr.BoundWidth -gt 0) -and ($tr.BoundHeight -gt 0)) {
        'USABLE - TextRange2.Bound* returns real geometry. Viable as the TextBounds model.'
    } else {
        'NOT POPULATED - Bound* came back zero. Text-aware alignment needs another mechanism.'
    }

    Add-Finding -Number 5 `
        -Question 'Are TextRange2.BoundLeft/Top/Width/Height reliable?' `
        -Answer $answer5 `
        -Evidence $textEvidence

    # =============================================================================================
    # Probe 6 - Can layout placeholder rects be read, for "align to content area"?
    # =============================================================================================
    $contentLayout = $null
    foreach ($layout in $presentation.SlideMaster.CustomLayouts) {
        if ($layout.Shapes.Placeholders.Count -ge 2) { $contentLayout = $layout; break }
    }

    if ($null -eq $contentLayout) {
        Add-Finding -Number 6 `
            -Question 'Can layout placeholder bounds be read, for "align to content area"?' `
            -Answer 'NO LAYOUT FOUND with 2+ placeholders in the default template. Retry against a real deck.'
    }
    else {
        $phEvidence = @("layout: $($contentLayout.Name)")
        $placeholders = $contentLayout.Shapes.Placeholders
        for ($i = 1; $i -le $placeholders.Count; $i++) {
            $ph = $placeholders.Item($i)
            $phEvidence += Format-Rect -Shape $ph -Label ('  ph[{0}] type={1}' -f $i, $ph.PlaceholderFormat.Type)
        }
        Add-Finding -Number 6 `
            -Question 'Can layout placeholder bounds be read, for "align to content area"?' `
            -Answer 'YES - CustomLayout placeholders expose geometry. PlaceholderBounds reference is viable.' `
            -Evidence $phEvidence
    }

    # =============================================================================================
    # Probe 4 - Does an OM change reach PowerPoint's undo stack? (needs a human)
    # =============================================================================================
    $u = $slide.Shapes.AddShape($msoShapeRectangle, 400, 300, 120, 60)
    $u.Name = 'ProbeUndo'
    $u.TextFrame.TextRange.Text = 'Undo me'
    $undoEvidence = @(Format-Rect -Shape $u -Label 'undo probe at start')
    $u.Left = 500
    $u.Top  = 380
    $undoEvidence += Format-Rect -Shape $u -Label 'moved from script'

    # If ExecuteMso is callable, the "push a neutral command to create an undo entry" trick is
    # available as a fallback. It is never the mechanism - our own UndoManager is.
    try {
        $ppt.CommandBars.ExecuteMso('Bold')
        $undoEvidence += 'CommandBars.ExecuteMso is callable, so the neutral-command trick is available.'
    }
    catch {
        $undoEvidence += "CommandBars.ExecuteMso failed: $($_.Exception.Message)"
    }

    Add-Finding -Number 4 `
        -Question "Does an object-model change reach PowerPoint's undo stack?" `
        -Answer 'NEEDS MANUAL CHECK - see the instructions printed at the end.' `
        -Evidence $undoEvidence
}
finally {
    Write-Host ''
    Write-Host '--- Probe 4 needs you ---------------------------------------------------------' -ForegroundColor Yellow
    Write-Host 'The shape named ProbeUndo was moved from script to (500, 380).'
    Write-Host 'In PowerPoint, click the slide once, then press Ctrl+Z and watch ProbeUndo:'
    Write-Host '  - jumps back to (400, 300)  -> OM changes DO land on the undo stack'
    Write-Host '  - does not move at all      -> they do NOT; our own UndoManager is mandatory'
    Write-Host '  - something else undoes     -> the stack holds unrelated entries; note what'
    Write-Host 'Re-run with -KeepOpen if the presentation has already closed.'
    Write-Host '-------------------------------------------------------------------------------' -ForegroundColor Yellow

    if ($OutFile -and $findings.Count) {
        $md = [System.Text.StringBuilder]::new()
        [void]$md.AppendLine('# PowerPoint object-model findings')
        [void]$md.AppendLine()
        [void]$md.AppendLine('Measured by `tools/Probe-ShapeGeometry.ps1`. These answers drive the design of')
        [void]$md.AppendLine('`AlignPro.Geometry` - do not re-derive them from documentation.')
        [void]$md.AppendLine()
        [void]$md.AppendLine(('- Date: {0}' -f (Get-Date -Format 'yyyy-MM-dd')))
        [void]$md.AppendLine(('- PowerPoint: version {0}, build {1}' -f $ppt.Version, $ppt.Build))
        [void]$md.AppendLine(('- PowerShell: {0}' -f $PSVersionTable.PSVersion))
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
        [void]$md.AppendLine('## Manual undo result')
        [void]$md.AppendLine()
        [void]$md.AppendLine('_Record the Ctrl+Z outcome from probe 4 here._')

        $resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutFile)
        $md.ToString() | Set-Content -Path $resolved -Encoding UTF8
        Write-Host ("Findings written to {0}" -f $resolved) -ForegroundColor Green
    }

    if (-not $KeepOpen) {
        if ($null -ne $presentation) {
            Write-Host 'Closing the scratch presentation without saving.' -ForegroundColor DarkGray
            $presentation.Saved = $msoTrue
            $presentation.Close()
        }
        if ($startedPowerPoint) {
            Write-Host 'Quitting the PowerPoint instance this script started.' -ForegroundColor DarkGray
            $ppt.Quit()
        }
    }
    else {
        Write-Host 'Leaving the scratch presentation open (-KeepOpen). It has never been saved.' -ForegroundColor DarkGray
    }
}
