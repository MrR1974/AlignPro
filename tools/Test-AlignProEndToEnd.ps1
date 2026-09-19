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
