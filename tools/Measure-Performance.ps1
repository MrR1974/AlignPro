<#
.SYNOPSIS
    Times AlignPro's operations on a busy slide, so a performance change has a number before and after.

.DESCRIPTION
    Builds one slide in a scratch presentation - a 10 x 10 grid of text boxes nudged off true, and
    five groups of four - then times, through the add-in's automation surface:

      - Align left on two shapes, which is mostly the fixed cost every command pays
      - Align left on every shape on the slide
      - Stack on two shapes, on that 120-shape slide
      - Duplicate: one shape twenty times, and ten shapes five times each
      - Tidy with nothing selected, over the whole slide

    Each is undone after timing, so every run starts from the same slide, and each is taken as the
    best of three: the fastest run is the one least disturbed by everything else on the machine.
    The times include the cross-process call from this script, which is the same for every row.

    Restarts PowerPoint first so the current build is loaded, as the test harnesses do. Nothing is
    saved.

.PARAMETER Runs
    How many times to time each operation. The best is reported.

.EXAMPLE
    .\Measure-Performance.ps1
#>
[CmdletBinding()]
param(
    [int] $Runs = 3
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$msoTrue = -1
$msoFalse = 0
$msoShapeRectangle = 1
$ppLayoutBlank = 12

if (Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue) {
    throw 'PowerPoint is running. Close it first: this script restarts it to load the current build.'
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $msoTrue
$api = $ppt.COMAddIns.Item('AlignPro.AddIn').Object
if (-not $api) { throw 'AlignPro is not loaded.' }

$pres = $ppt.Presentations.Add()
try {
    $slide = $pres.Slides.Add(1, $ppLayoutBlank)

    # The grid: text in every box, so text bounds cost what they cost on a real slide.
    $names = @()
    for ($i = 0; $i -lt 100; $i++) {
        $x = 20 + ($i % 10) * 90 + (($i * 7) % 5 - 2) * 0.4
        $y = 20 + [Math]::Floor($i / 10) * 45 + (($i * 3) % 5 - 2) * 0.4
        $s = $slide.Shapes.AddShape($msoShapeRectangle, $x, $y, 70, 30)
        $s.Name = "Box$i"
        $s.TextFrame.TextRange.Text = "Box $i"
        $names += $s.Name
    }
    for ($g = 0; $g -lt 5; $g++) {
        $members = @()
        for ($k = 0; $k -lt 4; $k++) {
            $s = $slide.Shapes.AddShape($msoShapeRectangle, 20 + $g * 180 + $k * 40, 480, 30, 30)
            $s.Name = "G$g-$k"
            $members += $s.Name
        }
        $slide.Shapes.Range($members).Group().Name = "Group$g"
    }
    Write-Host ("Slide built: {0} top-level shapes." -f $slide.Shapes.Count) -ForegroundColor DarkGray

    [void]$api.SetReference('SelectionBounds')
    [void]$api.SetBoundsModel('ShapeFrame')
    [void]$api.SetTidyTolerance('')

    function Select-Names { param([string[]] $Names)
        $first = $true
        foreach ($n in $Names) { $slide.Shapes.Item($n).Select($(if ($first) { $msoTrue } else { $msoFalse })); $first = $false }
    }

    # Times one operation: set up, run, undo - $Runs times - and keeps the best.
    function Measure-Op {
        param([string] $Name, [scriptblock] $Setup, [scriptblock] $Run, [int] $Undos = 1)
        $best = [double]::MaxValue
        $said = ''
        for ($r = 0; $r -lt $Runs; $r++) {
            & $Setup
            $watch = [System.Diagnostics.Stopwatch]::StartNew()
            $said = & $Run
            $watch.Stop()
            $best = [Math]::Min($best, $watch.Elapsed.TotalMilliseconds)
            for ($u = 0; $u -lt $Undos; $u++) { $ppt.CommandBars.ExecuteMso('Undo') }
        }
        [pscustomobject]@{ Operation = $Name; 'Best (ms)' = [Math]::Round($best, 0); Said = $said }
    }

    $results = @(
        Measure-Op 'Align left, 2 shapes' { Select-Names @('Box11', 'Box22') } { $api.RunVerb('AlignLeft') }
        Measure-Op 'Align left, all 100 boxes' { Select-Names $names } { $api.RunVerb('AlignLeft') }
        Measure-Op 'Stack, 2 shapes of 105' { Select-Names @('Box5', 'Box50') } { $api.RunOrder('StackFirstOnTop') }
        Measure-Op 'Duplicate 1 shape x20' { [void]$api.SetDuplicate(3, 3, 0, 20, 'OwnCentre'); Select-Names @('Box0') } { $api.RunDuplicate() }
        Measure-Op 'Duplicate 10 shapes x5' { [void]$api.SetDuplicate(3, 3, 0, 5, 'OwnCentre'); Select-Names ($names[0..9]) } { $api.RunDuplicate() }
        Measure-Op 'Tidy, whole slide' { $ppt.ActiveWindow.Selection.Unselect() } { $api.RunTidy() }
    )

    $results | Format-Table -AutoSize -Wrap
}
finally {
    $pres.Saved = $msoTrue
    $pres.Close()
    $ppt.Quit()
}
