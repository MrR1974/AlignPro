<#
.SYNOPSIS
    Runs AlignPro end to end inside PowerPoint and asserts the results, with no manual clicking.

.DESCRIPTION
    The add-in exposes AlignProAutomation through Application.COMAddIns.Item("AlignPro.AddIn").Object,
    so a verb can be invoked from script exactly as the ribbon would. Combined with
    CommandBars.ExecuteMso('Undo') standing in for Ctrl+Z, that makes the behaviour that used to need a
    person - especially undo granularity - fully testable here.

    Restarts PowerPoint first so the current build is loaded, then runs each case in its own scratch
    presentation. Nothing is saved.

.EXAMPLE
    .\Test-AlignProEndToEnd.ps1
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$msoTrue           = -1
$msoFalse          = 0
# MsoZOrderCmd, for building the ordering fixture.
$msoSendToBack     = 1
$msoBringForward   = 2
$msoShapeRectangle = 1
$msoShapeArc       = 25
$ppLayoutBlank     = 12
$tolerance         = 0.75

$script:results = [System.Collections.Generic.List[object]]::new()

function Add-Result {
    param([string] $Case, [bool] $Passed, [string] $Detail)
    $script:results.Add([pscustomobject]@{
        Case   = $Case
        Result = if ($Passed) { 'PASS' } else { 'FAIL' }
        Detail = $Detail
    })
}

# --- restart PowerPoint as a user-owned process --------------------------------------------------
function Restart-PowerPoint {
    $running = Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue
    if ($running) {
        try {
            $existing = New-Object -ComObject PowerPoint.Application
            for ($i = 1; $i -le $existing.Presentations.Count; $i++) { $existing.Presentations.Item($i).Saved = $msoTrue }
            $existing.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($existing)
        } catch { }
        $deadline = (Get-Date).AddSeconds(30)
        while ((Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue) -and (Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 500
        }
    }

    $appPaths = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE'
    $exe = $null
    if (Test-Path $appPaths) { try { $exe = (Get-Item $appPaths).GetValue('') } catch { $exe = $null } }
    if (-not $exe -or -not (Test-Path $exe)) {
        $exe = Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\POWERPNT.EXE'
    }
    Start-Process -FilePath $exe | Out-Null

    $deadline = (Get-Date).AddSeconds(60)
    do {
        Start-Sleep -Milliseconds 500
        $proc = Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue | Select-Object -First 1
    } while ((-not $proc -or $proc.MainWindowHandle -eq 0) -and (Get-Date) -lt $deadline)
    Start-Sleep -Seconds 2
}

Write-Host 'Restarting PowerPoint with the current build...' -ForegroundColor DarkGray
Restart-PowerPoint

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $msoTrue

$addin = $null
for ($i = 1; $i -le $ppt.COMAddIns.Count; $i++) {
    $candidate = $ppt.COMAddIns.Item($i)
    if ($candidate.Description -like '*AlignPro*') { $addin = $candidate }
}
if (-not $addin) { throw 'AlignPro is not loaded in PowerPoint.' }
if (-not $addin.Connect) { throw 'AlignPro is registered but not connected.' }

$api = $addin.Object
if (-not $api) { throw 'AlignPro loaded but exposed no automation object (RequestComAddInAutomationService).' }
Write-Host ("Automation surface reached. " + $api.Describe()) -ForegroundColor DarkGray

# --- fixtures ------------------------------------------------------------------------------------
function New-Fixture {
    param([switch] $WithRotation)

    $pres = $ppt.Presentations.Add()
    $slide = $pres.Slides.Add(1, $ppLayoutBlank)

    $specs = @(
        @{ N = 'A'; X = 120; Y = 120; W = 140; H = 70;  R = 0 }
        @{ N = 'B'; X = 300; Y = 220; W = 90;  H = 110; R = 0 }
        @{ N = 'C'; X = 520; Y = 150; W = 160; H = 80;  R = $(if ($WithRotation) { 30 } else { 0 }) }
    )
    foreach ($spec in $specs) {
        $s = $slide.Shapes.AddShape($msoShapeRectangle, $spec.X, $spec.Y, $spec.W, $spec.H)
        $s.Name = $spec.N
        $s.TextFrame.TextRange.Text = $spec.N
        if ($spec.R -ne 0) { $s.Rotation = $spec.R }
    }
    [pscustomobject]@{ Presentation = $pres; Slide = $slide }
}

function Select-All { param($Fixture) $Fixture.Slide.Shapes.Range(@('A', 'B', 'C')).Select($msoTrue) }

function Get-Frames {
    param($Fixture)
    $map = @{}
    foreach ($n in 'A', 'B', 'C') {
        $s = $Fixture.Slide.Shapes.Item($n)
        $map[$n] = [pscustomobject]@{ Left = [double]$s.Left; Top = [double]$s.Top }
    }
    $map
}

function Close-Fixture { param($Fixture) try { $Fixture.Presentation.Saved = $msoTrue; $Fixture.Presentation.Close() } catch { } }

function Test-Same {
    param($Expected, $Actual, [string] $Axis)
    foreach ($n in 'A', 'B', 'C') {
        $e = if ($Axis -eq 'Left') { $Expected[$n].Left } else { $Expected[$n].Top }
        $a = if ($Axis -eq 'Left') { $Actual[$n].Left } else { $Actual[$n].Top }
        if ([Math]::Abs($e - $a) -gt $tolerance) { return $false }
    }
    $true
}

# =================================================================================================
# Case 1: one operation, one undo - and the deck must survive
# =================================================================================================
$f = New-Fixture
try {
    $before = Get-Frames -Fixture $f
    Select-All -Fixture $f
    $status = $api.RunVerb('AlignLeft')
    if ($status) { Add-Result 'Single op runs' $false $status }
    else {
        $after = Get-Frames -Fixture $f
        $moved = -not (Test-Same -Expected $before -Actual $after -Axis 'Left')
        Add-Result 'Single op moves shapes' $moved ("A.left {0:F0} -> {1:F0}" -f $before.A.Left, $after.A.Left)

        $ppt.CommandBars.ExecuteMso('Undo')

        if ($f.Presentation.Slides.Count -eq 0) {
            Add-Result 'Undo keeps the deck' $false 'the whole deck was undone'
        }
        else {
            Add-Result 'Undo keeps the deck' $true 'slides intact'
            $restored = Get-Frames -Fixture $f
            $ok = Test-Same -Expected $before -Actual $restored -Axis 'Left'
            Add-Result 'Undo restores exactly one op' $ok `
                ("A.left back to {0:F0} (was {1:F0})" -f $restored.A.Left, $before.A.Left)
        }
    }
}
finally { Close-Fixture -Fixture $f }

# =================================================================================================
# Case 2: two operations, one undo must reverse only the second
# =================================================================================================
$f = New-Fixture
try {
    Select-All -Fixture $f
    [void]$api.RunVerb('AlignLeft')
    $afterFirst = Get-Frames -Fixture $f

    Select-All -Fixture $f
    [void]$api.RunVerb('AlignTop')
    $afterSecond = Get-Frames -Fixture $f

    $ppt.CommandBars.ExecuteMso('Undo')

    if ($f.Presentation.Slides.Count -eq 0) {
        Add-Result 'Two ops undo separately' $false 'the whole deck was undone'
    }
    else {
        $now = Get-Frames -Fixture $f
        $topsReverted = Test-Same -Expected $afterFirst -Actual $now -Axis 'Top'
        $leftsHeld = Test-Same -Expected $afterFirst -Actual $now -Axis 'Left'
        Add-Result 'Two ops undo separately' ($topsReverted -and $leftsHeld) `
            ("tops reverted={0} lefts held={1}" -f $topsReverted, $leftsHeld)
    }
}
finally { Close-Fixture -Fixture $f }

# =================================================================================================
# Case 3: rotation awareness - the premise of the project
# =================================================================================================
$f = New-Fixture -WithRotation
try {
    # Shape frame: PowerPoint's own behaviour. C is rotated 30 degrees; its frame left is 520, so the
    # leftmost frame is A at 120 and every frame should land on 120.
    [void]$api.SetBoundsModel('ShapeFrame')
    Select-All -Fixture $f
    [void]$api.RunVerb('AlignLeft')
    $frameMode = Get-Frames -Fixture $f
    $allAt120 = ([Math]::Abs($frameMode.A.Left - 120) -le $tolerance) -and
                ([Math]::Abs($frameMode.C.Left - 120) -le $tolerance)
    Add-Result 'Frame mode aligns frames' $allAt120 ("C.left = {0:F1}" -f $frameMode.C.Left)

    # The rotated shape's visual left edge is further left than its frame, so in frame mode it is NOT
    # visually flush. 160x80 at 30 degrees spans 160*cos30 + 80*sin30 = 178.6 wide, centred on the
    # frame centre, so its visual left sits about 9.3pt left of its frame.
    $theta = 30 * [Math]::PI / 180
    $visualWidth = 160 * [Math]::Cos($theta) + 80 * [Math]::Sin($theta)
    $cVisualLeftFrameMode = $frameMode.C.Left + 160 / 2 - $visualWidth / 2
    $misalignment = [Math]::Abs($cVisualLeftFrameMode - $frameMode.A.Left)
    Add-Result 'Frame mode misaligns rotated shape' ($misalignment -gt 2) `
        ("rotated shape sits {0:F1}pt out visually" -f $misalignment)

    # AlignPro's own undo, not ExecuteMso. This is only resetting between two sub-checks, and the
    # native one reaches the coalesced entry that still covers this fixture's own Slides.Add - which
    # deletes the slide and takes the rest of the case with it. Cases 1 and 2 above use ExecuteMso
    # deliberately, because there the coalescing behaviour IS what is under test.
    [void]$api.Undo()

    # Visual bounds: every shape's VISUAL left edge should line up.
    [void]$api.SetBoundsModel('VisualBounds')
    Select-All -Fixture $f
    [void]$api.RunVerb('AlignLeft')
    $visualMode = Get-Frames -Fixture $f

    $cVisualLeft = $visualMode.C.Left + 160 / 2 - $visualWidth / 2
    $flush = [Math]::Abs($cVisualLeft - $visualMode.A.Left) -le $tolerance
    Add-Result 'Visual mode aligns what you see' $flush `
        ("A.left={0:F1} C visual left={1:F1}" -f $visualMode.A.Left, $cVisualLeft)
}
finally { Close-Fixture -Fixture $f }

# =================================================================================================
# Case 5: the match-size margin
# =================================================================================================
$f = New-Fixture
try {
    # A is 140x70, B is 90x110, C is 160x80. C is selected last, so it is the anchor.
    $f.Slide.Shapes.Range(@('A', 'B', 'C')).Select($msoTrue)
    [void]$api.SetSizeMargin('10')
    [void]$api.SetSizeMarginMode('Uniform')
    [void]$api.RunVerb('MatchBoth')

    # A gap of 10 per side takes 20 off each dimension of the 160x80 anchor.
    $a = $f.Slide.Shapes.Item('A')
    $b = $f.Slide.Shapes.Item('B')
    $uniform = ([Math]::Abs($a.Width - 140) -le $tolerance) -and ([Math]::Abs($a.Height - 60) -le $tolerance) -and
               ([Math]::Abs($b.Width - 140) -le $tolerance) -and ([Math]::Abs($b.Height - 60) -le $tolerance)
    Add-Result 'Uniform margin insets every side' $uniform `
        ("A {0:F0}x{1:F0}, B {2:F0}x{3:F0}, expected 140x60" -f $a.Width, $a.Height, $b.Width, $b.Height)

    # AlignPro's own undo, for the same reason as case 3.
    [void]$api.Undo()

    # Cascade: A is two steps back from the anchor, B one step.
    $f.Slide.Shapes.Range(@('A', 'B', 'C')).Select($msoTrue)
    [void]$api.SetSizeMarginMode('Cascade')
    [void]$api.RunVerb('MatchBoth')

    $a = $f.Slide.Shapes.Item('A')
    $b = $f.Slide.Shapes.Item('B')
    $c = $f.Slide.Shapes.Item('C')
    $cascade = ([Math]::Abs($a.Width - 120) -le $tolerance) -and
               ([Math]::Abs($b.Width - 140) -le $tolerance) -and
               ([Math]::Abs($c.Width - 160) -le $tolerance)
    Add-Result 'Cascade tiers along the selection' $cascade `
        ("A {0:F0}, B {1:F0}, anchor C {2:F0}, expected 120/140/160" -f $a.Width, $b.Width, $c.Width)

    [void]$api.SetSizeMarginMode('None')
    [void]$api.SetSizeMargin('')
}
finally { Close-Fixture -Fixture $f }

# =================================================================================================
# Case 6: ordering, in place, and its undo
# =================================================================================================
$f = New-Fixture
try {
    # A fourth shape that is deliberately NOT selected, sitting between B and C in the stack.
    $bar = $f.Slide.Shapes.AddShape($msoShapeRectangle, 400, 100, 40, 300)
    $bar.Name = 'BAR'
    # Created last, so it is frontmost; move it to sit directly behind C.
    $bar.ZOrder($msoSendToBack)
    $bar.ZOrder($msoBringForward)   # now above A
    $bar.ZOrder($msoBringForward)   # now above B, still below C

    $orderBefore = @(1..$f.Slide.Shapes.Count | ForEach-Object { $f.Slide.Shapes.Item($_).Name })
    $barSlotBefore = [array]::IndexOf($orderBefore, 'BAR')

    # Select A, then B, then C. A was clicked first, so A must end on top.
    $f.Slide.Shapes.Item('A').Select($msoTrue)
    $f.Slide.Shapes.Item('B').Select($msoFalse)
    $f.Slide.Shapes.Item('C').Select($msoFalse)

    $status = $api.RunOrder('StackFirstOnTop')
    if ($status) { Add-Result 'Order runs' $false $status }

    $orderAfter = @(1..$f.Slide.Shapes.Count | ForEach-Object { $f.Slide.Shapes.Item($_).Name })

    # Top of the stack is the last entry in the collection.
    Add-Result 'First selected ends on top' ($orderAfter[-1] -eq 'A') `
        ("stack back-to-front: " + ($orderAfter -join ', '))

    # The whole point of ordering in place: BAR must not have changed layer.
    $barSlotAfter = [array]::IndexOf($orderAfter, 'BAR')
    Add-Result 'Unselected shape keeps its layer' ($barSlotBefore -eq $barSlotAfter) `
        ("BAR slot {0} -> {1}" -f $barSlotBefore, $barSlotAfter)

    # Undo must put the original stacking back exactly.
    [void]$api.Undo()
    $orderUndone = @(1..$f.Slide.Shapes.Count | ForEach-Object { $f.Slide.Shapes.Item($_).Name })
    Add-Result 'Undo restores the stacking order' `
        (($orderUndone -join ',') -eq ($orderBefore -join ',')) `
        ("{0}  ->  {1}" -f ($orderBefore -join ','), ($orderUndone -join ','))

    # And redo must put it back again.
    [void]$api.Redo()
    $orderRedone = @(1..$f.Slide.Shapes.Count | ForEach-Object { $f.Slide.Shapes.Item($_).Name })
    Add-Result 'Redo reapplies the stacking order' `
        (($orderRedone -join ',') -eq ($orderAfter -join ',')) `
        ("{0}  ->  {1}" -f ($orderAfter -join ','), ($orderRedone -join ','))

    # A single shape has no order to speak of, so it should be refused rather than silently doing
    # nothing.
    $f.Slide.Shapes.Item('A').Select($msoTrue)
    $refusal = $api.RunOrder('StackFirstOnTop')
    Add-Result 'Order refuses a single shape' ($refusal -ne '') ("said: " + $refusal)
}
finally { Close-Fixture -Fixture $f }

# --- helpers for the cases below, which each need a selection in a specific order ------------------
function New-EmptyFixture {
    $pres = $ppt.Presentations.Add()
    $slide = $pres.Slides.Add(1, $ppLayoutBlank)
    [pscustomobject]@{ Presentation = $pres; Slide = $slide }
}

function Add-Rect {
    param($Fixture, [string] $Name, [double] $X, [double] $Y, [double] $W, [double] $H, [double] $R = 0, [int] $Type = $msoShapeRectangle)
    $s = $Fixture.Slide.Shapes.AddShape($Type, $X, $Y, $W, $H)
    $s.Name = $Name
    if ($R -ne 0) { $s.Rotation = $R }
    $s
}

# Selection order is the input for anchors and curves, and Range(...).Select() promises nothing about
# it, so shapes are added one at a time.
function Select-InOrder {
    param($Fixture, [string[]] $Names)
    $first = $true
    foreach ($n in $Names) {
        $Fixture.Slide.Shapes.Item($n).Select($(if ($first) { $msoTrue } else { $msoFalse }))
        $first = $false
    }
}

function Get-Centre {
    param($Fixture, [string] $Name)
    $s = $Fixture.Slide.Shapes.Item($Name)
    [pscustomobject]@{ X = $s.Left + $s.Width / 2; Y = $s.Top + $s.Height / 2; R = [double]$s.Rotation }
}

function Test-Near { param([double] $A, [double] $B, [double] $Within = $tolerance) [Math]::Abs($A - $B) -le $Within }

# Settings persist in the add-in between cases, so each case below pins what it relies on.
[void]$api.SetReference('SelectionBounds')
[void]$api.SetBoundsModel('ShapeFrame')
[void]$api.SetExactSpacing('')

# =================================================================================================
# Case 7: Grow - the direction supplies the sign, and a negative margin is refused
# =================================================================================================
$f = New-Fixture
try {
    $refused = $api.SetSizeMargin('-10')
    Add-Result 'A negative margin is refused' ($refused -like '*Grow*') ("said: " + $refused)

    [void]$api.SetSizeMargin('10')
    [void]$api.SetSizeDirection('Grow')
    [void]$api.SetSizeMarginMode('Cascade')
    Select-InOrder -Fixture $f -Names @('A', 'B', 'C')
    $status = $api.RunVerb('MatchBoth')

    # Anchor C is 160x80. Grow with Cascade: A is two steps out, B one.
    $a = $f.Slide.Shapes.Item('A'); $b = $f.Slide.Shapes.Item('B')
    $ok = (Test-Near $a.Width 200) -and (Test-Near $a.Height 120) -and (Test-Near $b.Width 180)
    Add-Result 'Grow with Cascade makes the first selected largest' $ok `
        ("A {0:F0}x{1:F0}, B {2:F0}, expected 200x120 and 180 {3}" -f $a.Width, $a.Height, $b.Width, $status)
}
finally {
    [void]$api.SetSizeDirection('Shrink')
    [void]$api.SetSizeMarginMode('None')
    [void]$api.SetSizeMargin('')
    Close-Fixture -Fixture $f
}

# =================================================================================================
# Case 8: Match rotation, and undo puts the angles back
# =================================================================================================
$f = New-EmptyFixture
try {
    [void](Add-Rect $f 'R1' 100 100 80 50 10)
    [void](Add-Rect $f 'R2' 250 100 60 60)
    [void](Add-Rect $f 'Anchor' 450 100 120 60 35)
    $before = Get-Centre $f 'R1'

    Select-InOrder -Fixture $f -Names @('R1', 'R2', 'Anchor')
    $status = $api.RunVerb('MatchRotation')

    $r1 = Get-Centre $f 'R1'; $r2 = Get-Centre $f 'R2'
    Add-Result 'Match rotation takes the anchor angle' ((Test-Near $r1.R 35 0.1) -and (Test-Near $r2.R 35 0.1)) `
        ("R1={0:F1} R2={1:F1}, expected 35 {2}" -f $r1.R, $r2.R, $status)
    Add-Result 'Match rotation keeps centres' ((Test-Near $r1.X $before.X) -and (Test-Near $r1.Y $before.Y)) `
        ("R1 centre ({0:F1},{1:F1}) was ({2:F1},{3:F1})" -f $r1.X, $r1.Y, $before.X, $before.Y)

    $ppt.CommandBars.ExecuteMso('Undo')
    $undone = Get-Centre $f 'R1'
    Add-Result 'Ctrl+Z restores the angle' (Test-Near $undone.R 10 0.1) ("R1 back to {0:F1}, was 10" -f $undone.R)
}
finally { Close-Fixture -Fixture $f }

# =================================================================================================
# Case 9: Duplicate, and its undo and redo
# =================================================================================================
$f = New-EmptyFixture
try {
    [void](Add-Rect $f 'Orig' 100 200 60 40)
    [void](Add-Rect $f 'Partner' 100 260 60 20)
    $countBefore = $f.Slide.Shapes.Count

    [void]$api.SetDuplicate(80, 0, 0, 3, 'OwnCentre')
    [void]$api.SetRotateShapes($true)
    Select-InOrder -Fixture $f -Names @('Orig', 'Partner')
    $status = $api.RunDuplicate()

    $countAfter = $f.Slide.Shapes.Count
    Add-Result 'Duplicate makes every copy' ($countAfter -eq $countBefore + 6) `
        ("{0} shapes, expected {1} {2}" -f $countAfter, ($countBefore + 6), $status)

    # The last copy of Orig: three steps of 80 to the right, and the pair keeps its layout.
    $lefts = @(); for ($i = 1; $i -le $countAfter; $i++) { $lefts += [Math]::Round($f.Slide.Shapes.Item($i).Left) }
    $ok = ($lefts -contains 180) -and ($lefts -contains 260) -and ($lefts -contains 340)
    Add-Result 'Copies step by the offset' $ok ("lefts: " + (($lefts | Sort-Object -Unique) -join ', '))

    $selected = $ppt.ActiveWindow.Selection.ShapeRange.Count
    Add-Result 'Originals and copies are left selected' ($selected -eq 8) "$selected selected"

    [void]$api.Undo()
    Add-Result 'Undo removes every copy' ($f.Slide.Shapes.Count -eq $countBefore) "$($f.Slide.Shapes.Count) shapes"

    [void]$api.Redo()
    Add-Result 'Redo makes them again' ($f.Slide.Shapes.Count -eq $countBefore + 6) "$($f.Slide.Shapes.Count) shapes"

    # The copies redo made have new ids; undo must find those, not the ids from the first run.
    [void]$api.Undo()
    Add-Result 'Undo after redo removes the new copies' ($f.Slide.Shapes.Count -eq $countBefore) "$($f.Slide.Shapes.Count) shapes"

    # Native undo: duplicate again, then Ctrl+Z must take away the copies and nothing else.
    Select-InOrder -Fixture $f -Names @('Orig', 'Partner')
    [void]$api.RunDuplicate()
    $ppt.CommandBars.ExecuteMso('Undo')
    $nativeOk = ($f.Presentation.Slides.Count -eq 1) -and ($f.Slide.Shapes.Count -eq $countBefore)
    Add-Result 'Ctrl+Z removes only the copies' $nativeOk `
        ("{0} slides, {1} shapes (expected 1 and {2})" -f $f.Presentation.Slides.Count, $f.Slide.Shapes.Count, $countBefore)
}
finally { Close-Fixture -Fixture $f }

# =================================================================================================
# Case 10: Duplicate round the slide centre - the radial array
# =================================================================================================
$f = New-EmptyFixture
try {
    # 20x20, centred 100pt above the slide centre (480, 270).
    [void](Add-Rect $f 'Spoke' 470 160 20 20)
    [void]$api.SetDuplicate(0, 0, 90, 3, 'SlideCentre')
    [void]$api.SetRotateShapes($true)
    Select-InOrder -Fixture $f -Names @('Spoke')
    $status = $api.RunDuplicate()

    $found = 0
    foreach ($spot in @(@{ X = 580; Y = 270; R = 90 }, @{ X = 480; Y = 370; R = 180 }, @{ X = 380; Y = 270; R = 270 })) {
        for ($i = 1; $i -le $f.Slide.Shapes.Count; $i++) {
            $s = $f.Slide.Shapes.Item($i)
            if ((Test-Near ($s.Left + 10) $spot.X) -and (Test-Near ($s.Top + 10) $spot.Y) -and (Test-Near $s.Rotation $spot.R 0.1)) { $found++ }
        }
    }
    Add-Result 'Duplicate turns copies round the slide centre' ($found -eq 3) "$found of 3 copies at 3, 6 and 9 o'clock $status"
}
finally { Close-Fixture -Fixture $f }

# =================================================================================================
# Case 11: Distribute along a circle, an arc and a path
# =================================================================================================
$f = New-EmptyFixture
try {
    $msoShapeOval = 9
    foreach ($n in 1..4) { [void](Add-Rect $f "Dot$n" (60 * $n) 450 20 20) }
    [void](Add-Rect $f 'Ring' 380 170 200 200 0 $msoShapeOval)   # centred on (480, 270), radius 100

    [void]$api.SetExactSpacing('')
    [void]$api.SetRotateShapes($true)
    Select-InOrder -Fixture $f -Names @('Dot1', 'Dot2', 'Dot3', 'Dot4', 'Ring')
    $status = $api.RunDistributeCurve()

    $expected = @(@{ X = 480; Y = 170; R = 0 }, @{ X = 580; Y = 270; R = 90 }, @{ X = 480; Y = 370; R = 180 }, @{ X = 380; Y = 270; R = 270 })
    $ok = 0
    for ($i = 0; $i -lt 4; $i++) {
        $c = Get-Centre $f "Dot$($i + 1)"
        if ((Test-Near $c.X $expected[$i].X) -and (Test-Near $c.Y $expected[$i].Y) -and (Test-Near $c.R $expected[$i].R 0.6)) { $ok++ }
    }
    Add-Result 'Circle: four shapes at twelve, three, six and nine' ($ok -eq 4) "$ok of 4 placed and turned $status"

    [void]$api.Undo()
    $back = Get-Centre $f 'Dot1'
    Add-Result 'Circle: undo puts them back' ((Test-Near $back.X 70) -and (Test-Near $back.Y 460) -and (Test-Near $back.R 0 0.1)) `
        ("Dot1 at ({0:F1},{1:F1}) rot {2:F1}" -f $back.X, $back.Y, $back.R)

    # Arc: PowerPoint's default quarter, from twelve o'clock round to three. Its frame is the box of
    # the arc and its centre (probe 10), so the centre is the frame's bottom-left corner and the
    # quarter runs from the top-left corner to the bottom-right.
    $f.Slide.Shapes.Item('Ring').Delete()
    $arc = Add-Rect $f 'Bow' 380 170 200 200 0 $msoShapeArc
    $adj1 = $arc.Adjustments.Item(1); $adj2 = $arc.Adjustments.Item(2)
    Select-InOrder -Fixture $f -Names @('Dot1', 'Dot2', 'Dot3', 'Bow')
    $status = $api.RunDistributeCurve()
    $first = Get-Centre $f 'Dot1'; $last = Get-Centre $f 'Dot3'
    $ok = (Test-Near $first.X 380) -and (Test-Near $first.Y 170) -and (Test-Near $last.X 580) -and (Test-Near $last.Y 370)
    Add-Result 'Arc: both ends included' $ok `
        ("adj=({0},{1}); first ({2:F1},{3:F1}) last ({4:F1},{5:F1}) {6}" -f $adj1, $adj2, $first.X, $first.Y, $last.X, $last.Y, $status)

    # Flipped, the frame mirrors about the ellipse's centre (probe 10): it now spans 180..380, and
    # the quarter runs from twelve o'clock round to nine.
    $arc.Flip(0)
    Select-InOrder -Fixture $f -Names @('Dot1', 'Dot2', 'Dot3', 'Bow')
    $status = $api.RunDistributeCurve()
    $first = Get-Centre $f 'Dot1'; $last = Get-Centre $f 'Dot3'
    $ok = (Test-Near $first.X 380) -and (Test-Near $first.Y 170) -and (Test-Near $last.X 180) -and (Test-Near $last.Y 370)
    Add-Result 'Arc: a flipped arc is followed as drawn' $ok `
        ("frame L={0:F1}; first ({1:F1},{2:F1}) last ({3:F1},{4:F1}) {5}" -f $arc.Left, $first.X, $first.Y, $last.X, $last.Y, $status)

    $arc.Rotation = 30
    $refusal = $api.RunDistributeCurve()
    Add-Result 'Arc: a rotated arc is refused' ($refusal -like '*rotated arc*') ("said: " + $refusal)

    # Path: a freeform L, 300pt long.
    $builder = $f.Slide.Shapes.BuildFreeform(1, 100, 100)
    $builder.AddNodes(0, 0, 300, 100)
    $builder.AddNodes(0, 0, 300, 200)
    $path = $builder.ConvertToShape()
    $path.Name = 'Track'
    Select-InOrder -Fixture $f -Names @('Dot1', 'Dot2', 'Dot3', 'Track')
    $status = $api.RunDistributeCurve()
    $mid = Get-Centre $f 'Dot2'; $end = Get-Centre $f 'Dot3'
    $ok = (Test-Near $mid.X 250) -and (Test-Near $mid.Y 100) -and (Test-Near $end.X 300) -and (Test-Near $end.Y 200) -and (Test-Near $end.R 90 0.1)
    Add-Result 'Path: placed by distance, turned at the corner' $ok `
        ("middle ({0:F1},{1:F1}) end ({2:F1},{3:F1}) rot {4:F1} {5}" -f $mid.X, $mid.Y, $end.X, $end.Y, $end.R, $status)

    # Something that is not a curve must be refused by name, not approximated.
    Select-InOrder -Fixture $f -Names @('Dot1', 'Dot4')
    $refusal = $api.RunDistributeCurve()
    Add-Result 'A rectangle is not taken for a curve' ($refusal -like '*oval*') ("said: " + $refusal)
}
finally { Close-Fixture -Fixture $f }

# =================================================================================================
Write-Host ''
$script:results | Format-Table -AutoSize
$failed = @($script:results | Where-Object { $_.Result -eq 'FAIL' }).Count
if ($failed -eq 0) {
    Write-Host ("All {0} checks passed." -f $script:results.Count) -ForegroundColor Green
}
else {
    Write-Host ("{0} of {1} checks FAILED." -f $failed, $script:results.Count) -ForegroundColor Red
}
Write-Host ("Add-in log: " + $api.LogPath) -ForegroundColor DarkGray
Write-Host ''
