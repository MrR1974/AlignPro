<#
.SYNOPSIS
    Builds a scratch deck for exercising the AlignPro ribbon by hand.

.DESCRIPTION
    One slide per behaviour worth checking, each deliberately untidy so the effect of a command is
    obvious. The deck is never saved.

    Slide 1  Align          scattered rectangles, one rotated - the case native PowerPoint gets wrong
    Slide 2  Distribute     five rectangles of differing widths, so the four spacing modes disagree
    Slide 3  Groups         a group plus loose shapes, for group rigidity and resize refusal
    Slide 4  Text bounds    text boxes whose frames line up but whose text does not

.EXAMPLE
    .\New-TestDeck.ps1
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$msoTrue                      = -1
$msoShapeRectangle            = 1
$ppLayoutBlank                = 12
$msoTextOrientationHorizontal = 1

<#
    Start PowerPoint as an ordinary process rather than letting New-Object create it.

    An Office application created through COM automation is owned by the automation client, and
    exits when the last reference to it is released - which happens the moment this script ends,
    taking the deck with it. Starting the executable first and only then attaching leaves PowerPoint
    user-owned, so it stays put after the script exits.
#>
$wasRunning = [bool] (Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue)
if (-not $wasRunning) {
    # GetValue('') reads a key's default value; Get-ItemProperty would need a property name that
    # does not exist under Set-StrictMode.
    $appPaths = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE'
    $exe = $null
    if (Test-Path $appPaths) {
        try { $exe = (Get-Item $appPaths).GetValue('') } catch { $exe = $null }
    }
    if (-not $exe -or -not (Test-Path $exe)) {
        $exe = Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\POWERPNT.EXE'
    }
    if (-not (Test-Path $exe)) { throw "Could not locate POWERPNT.EXE (looked at '$exe')." }

    Write-Host 'Starting PowerPoint...' -ForegroundColor DarkGray
    Start-Process -FilePath $exe | Out-Null

    # Wait for a real window before attaching, so New-Object joins this instance instead of
    # spinning up a second one that nobody owns.
    $deadline = (Get-Date).AddSeconds(60)
    do {
        Start-Sleep -Milliseconds 500
        $proc = Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue | Select-Object -First 1
    } while ((-not $proc -or $proc.MainWindowHandle -eq 0) -and (Get-Date) -lt $deadline)

    if (-not $proc -or $proc.MainWindowHandle -eq 0) { throw 'PowerPoint did not finish starting.' }
    Start-Sleep -Seconds 1
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $msoTrue
$presentation = $ppt.Presentations.Add()

function Add-Slide {
    param([string] $Title)
    $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, $ppLayoutBlank)
    $caption = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, 30, 20, 700, 30)
    $caption.Name = 'Caption'
    $caption.TextFrame2.TextRange.Text = $Title
    $caption.TextFrame2.TextRange.Font.Size = 20
    $slide
}

function Add-Box {
    param($Slide, [string] $Name, [double] $X, [double] $Y, [double] $W, [double] $H, [double] $Rotation = 0)
    $shape = $Slide.Shapes.AddShape($msoShapeRectangle, $X, $Y, $W, $H)
    $shape.Name = $Name
    $shape.TextFrame.TextRange.Text = $Name
    if ($Rotation -ne 0) { $shape.Rotation = $Rotation }
    $shape
}

# --- Slide 1: alignment, including the rotated case ---------------------------------------------
$s1 = Add-Slide -Title 'Align - select all, then try Left with Measure = Shape frame vs Visual bounds'
Add-Box -Slide $s1 -Name 'A' -X 120 -Y 120 -W 140 -H 70              | Out-Null
Add-Box -Slide $s1 -Name 'B' -X 300 -Y 220 -W 90  -H 110             | Out-Null
Add-Box -Slide $s1 -Name 'C rotated' -X 520 -Y 150 -W 160 -H 80 -Rotation 30 | Out-Null
Add-Box -Slide $s1 -Name 'D' -X 700 -Y 320 -W 120 -H 60              | Out-Null

# --- Slide 2: distribution with unequal widths ---------------------------------------------------
$s2 = Add-Slide -Title 'Distribute - widths differ, so the four Space by modes give four answers'
Add-Box -Slide $s2 -Name 'P' -X 60  -Y 250 -W 200 -H 80 | Out-Null
Add-Box -Slide $s2 -Name 'Q' -X 320 -Y 250 -W 60  -H 80 | Out-Null
Add-Box -Slide $s2 -Name 'R' -X 430 -Y 250 -W 120 -H 80 | Out-Null
Add-Box -Slide $s2 -Name 'S' -X 640 -Y 250 -W 40  -H 80 | Out-Null
Add-Box -Slide $s2 -Name 'T' -X 780 -Y 250 -W 150 -H 80 | Out-Null

# --- Slide 3: groups -----------------------------------------------------------------------------
$s3 = Add-Slide -Title 'Groups - align moves the group rigidly; Match size should refuse it'
$g1 = Add-Box -Slide $s3 -Name 'G1' -X 100 -Y 180 -W 80 -H 60
$g2 = Add-Box -Slide $s3 -Name 'G2' -X 220 -Y 180 -W 80 -H 60
$g3 = Add-Box -Slide $s3 -Name 'G3' -X 340 -Y 180 -W 80 -H 60
$range = $s3.Shapes.Range(@($g1.Name, $g2.Name, $g3.Name))
$group = $range.Group()
$group.Name = 'TheGroup'
Add-Box -Slide $s3 -Name 'Loose1' -X 560 -Y 300 -W 140 -H 90 | Out-Null
Add-Box -Slide $s3 -Name 'Loose2' -X 760 -Y 380 -W 100 -H 50 | Out-Null

# --- Slide 4: text bounds ------------------------------------------------------------------------
$s4 = Add-Slide -Title 'Text bounds - frames start level at x=200; the text does not'
foreach ($spec in @(
    @{ Name = 'T1'; Y = 150; Text = 'Short' },
    @{ Name = 'T2'; Y = 230; Text = 'A longer caption' },
    @{ Name = 'T3'; Y = 310; Text = 'Middle' })) {
    $tb = $s4.Shapes.AddTextbox($msoTextOrientationHorizontal, 200, $spec.Y, 300, 40)
    $tb.Name = $spec.Name
    $tb.TextFrame2.TextRange.Text = $spec.Text
    $tb.TextFrame2.TextRange.Font.Size = 18
    # Different horizontal alignment moves the text within an identical frame.
    $tb.TextFrame2.TextRange.ParagraphFormat.Alignment = @(1, 2, 3)[[array]::IndexOf(@('T1', 'T2', 'T3'), $spec.Name)]
}

$ppt.ActiveWindow.View.GotoSlide(1)

Write-Host ''
Write-Host 'Test deck ready - four slides, never saved.' -ForegroundColor Green
Write-Host 'The AlignPro tab should be on the ribbon. Suggested first check:' -ForegroundColor Cyan
Write-Host '  1. Slide 1: select all four shapes (Ctrl+A).'
Write-Host '  2. Set Measure = Shape frame, click Align > Left. The rotated shape looks wrong.'
Write-Host '  3. AlignPro > Undo, set Measure = Visual bounds, click Align > Left again.'
Write-Host '     Now every shape is visually flush - that is the whole point of the project.'
Write-Host ''
Write-Host 'WARNING about Ctrl+Z on this deck specifically:' -ForegroundColor Yellow
Write-Host '  This script built the deck through the object model, which PowerPoint holds open as a'
Write-Host '  single undo entry. AlignPro only claims Undo once it has an operation of its own to'
Write-Host '  reverse, so pressing Ctrl+Z BEFORE your first AlignPro command falls through to'
Write-Host '  PowerPoint and removes every slide. Do an AlignPro command first, then Ctrl+Z is ours.'
Write-Host '  A deck you authored by hand does not have this problem - your own last edit closes the'
Write-Host '  group.' -ForegroundColor Yellow
Write-Host ''
Write-Host 'Close without saving when done.' -ForegroundColor DarkGray
