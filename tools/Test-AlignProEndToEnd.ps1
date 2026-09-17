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

    $ppt.CommandBars.ExecuteMso('Undo')

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
