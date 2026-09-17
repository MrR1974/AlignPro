<#
.SYNOPSIS
    Builds a saved sample presentation for trying AlignPro by hand.

.DESCRIPTION
    One slide per capability, each deliberately untidy and captioned with what to try and what should
    happen. Unlike New-TestDeck.ps1 this deck is saved to a real file, so it can be reopened and reused.

    The deck is built, saved, closed and then reopened. That last step matters: building it through the
    object model leaves one large undo entry open, and reopening the saved file starts with a clean undo
    history so Ctrl+Z behaves normally from the first click.

.PARAMETER Path
    Where to write the .pptx. Defaults to sample\AlignPro-Sample.pptx beside the project.

    Deliberately NOT your Documents folder: that is redirected to OneDrive, where PowerPoint turns
    AutoSave on and silently writes every experiment straight back into the file. A fixture that
    rewrites itself as you poke at it is no fixture at all. A local path keeps AutoSave off, so the
    deck only changes if you explicitly save it.

.PARAMETER Force
    Overwrite an existing file.

.EXAMPLE
    .\New-SampleDeck.ps1
#>
[CmdletBinding()]
param(
    [string] $Path,
    [switch] $Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$msoTrue                      = -1
$msoFalse                     = 0
$msoShapeRectangle            = 1
$msoShapeRoundedRectangle     = 5
$msoShapeOval                 = 9
$ppLayoutBlank                = 12
$msoTextOrientationHorizontal = 1
$ppSaveAsOpenXMLPresentation  = 24
$msoAlignLeft                 = 1
$msoAlignCenter               = 2
$msoAlignRight                = 3

if (-not $Path) {
    $sampleDir = Join-Path (Split-Path $PSScriptRoot -Parent) 'sample'
    if (-not (Test-Path $sampleDir)) { New-Item -ItemType Directory -Path $sampleDir | Out-Null }
    $Path = Join-Path $sampleDir 'AlignPro-Sample.pptx'
}
$Path = [System.IO.Path]::GetFullPath($Path)

$parent = Split-Path $Path -Parent
if ($parent -and -not (Test-Path $parent)) { New-Item -ItemType Directory -Path $parent | Out-Null }

if ((Test-Path $Path) -and -not $Force) {
    throw "'$Path' already exists. Use -Force to replace it."
}

function Rgb { param([int] $R, [int] $G, [int] $B) $R + ($G * 256) + ($B * 65536) }

$blue   = Rgb 31 119 180
$orange = Rgb 255 127 14
$green  = Rgb 44 160 44
$red    = Rgb 214 39 40
$purple = Rgb 148 103 189
$grey   = Rgb 127 127 127
$ink    = Rgb 40 40 40

# --- start PowerPoint as its own process so it survives this script ------------------------------
$wasRunning = [bool] (Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue)
if (-not $wasRunning) {
    $appPaths = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE'
    $exe = $null
    if (Test-Path $appPaths) { try { $exe = (Get-Item $appPaths).GetValue('') } catch { $exe = $null } }
    if (-not $exe -or -not (Test-Path $exe)) {
        $exe = Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\POWERPNT.EXE'
    }
    if (-not (Test-Path $exe)) { throw "Could not locate POWERPNT.EXE (looked at '$exe')." }

    Start-Process -FilePath $exe | Out-Null
    $deadline = (Get-Date).AddSeconds(60)
    do {
        Start-Sleep -Milliseconds 500
        $proc = Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue | Select-Object -First 1
    } while ((-not $proc -or $proc.MainWindowHandle -eq 0) -and (Get-Date) -lt $deadline)
    Start-Sleep -Seconds 1
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $msoTrue
$pres = $ppt.Presentations.Add()

# --- helpers --------------------------------------------------------------------------------------
function Add-Slide {
    param([string] $Title, [string] $Try)

    $slide = $pres.Slides.Add($pres.Slides.Count + 1, $ppLayoutBlank)

    $heading = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, 40, 24, 880, 34)
    $heading.Name = 'Heading'
    $heading.TextFrame2.TextRange.Text = $Title
    $heading.TextFrame2.TextRange.Font.Size = 24
    $heading.TextFrame2.TextRange.Font.Bold = $msoTrue
    $heading.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $ink

    $note = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, 40, 60, 880, 30)
    $note.Name = 'Instructions'
    $note.TextFrame2.TextRange.Text = $Try
    $note.TextFrame2.TextRange.Font.Size = 13
    $note.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $grey

    $slide
}

function Add-Box {
    param(
        $Slide, [string] $Name, [double] $X, [double] $Y, [double] $W, [double] $H,
        [int] $Colour = $blue, [double] $Rotation = 0, [int] $Shape = $msoShapeRectangle,
        [string] $Label
    )

    $s = $Slide.Shapes.AddShape($Shape, $X, $Y, $W, $H)
    $s.Name = $Name
    $s.Fill.ForeColor.RGB = $Colour
    $s.Line.Visible = $msoFalse
    $s.TextFrame2.TextRange.Text = if ($PSBoundParameters.ContainsKey('Label')) { $Label } else { $Name }
    $s.TextFrame2.TextRange.Font.Size = 14
    $s.TextFrame2.TextRange.Font.Bold = $msoTrue
    $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = (Rgb 255 255 255)
    if ($Rotation -ne 0) { $s.Rotation = $Rotation }
    $s
}

# =================================================================================================
# 1. Overview
# =================================================================================================
$s = Add-Slide -Title 'AlignPro sample deck' `
               -Try  'Each slide says what to try. Nothing here is precious - edit it freely and save over it if you like.'
$body = $s.Shapes.AddTextbox($msoTextOrientationHorizontal, 60, 130, 840, 340)
$body.TextFrame2.TextRange.Text = @'
The AlignPro tab is on the ribbon.

Two dropdowns carry the state the six align buttons are crossed with:

    Reference  -  what to align against: the anchor shape, the selection, the slide, the
                  slide margins, or the layout's content placeholder.

    Measure    -  which rectangle to align: the shape frame (what PowerPoint itself uses,
                  which ignores rotation), the visual bounds (what you actually see), or
                  the text inside.

Undo is PowerPoint's own - Ctrl+Z, the ribbon button and the Quick Access Toolbar all
reverse exactly one AlignPro command.

Slide 2 is the one to try first. It is the reason this add-in exists.
'@
$body.TextFrame2.TextRange.Font.Size = 15
$body.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $ink

# =================================================================================================
# 2. Rotation - the headline case
# =================================================================================================
$s = Add-Slide -Title 'Rotation: the case PowerPoint gets wrong' `
               -Try  'Select all four. Align > Left with Measure = Shape frame, undo, then again with Measure = Visual bounds.'
Add-Box -Slide $s -Name 'Plain1'  -X 140 -Y 150 -W 150 -H 80  -Colour $blue   | Out-Null
Add-Box -Slide $s -Name 'Tall'    -X 330 -Y 260 -W 90  -H 150 -Colour $green  | Out-Null
Add-Box -Slide $s -Name 'Rotated' -X 520 -Y 160 -W 180 -H 90  -Colour $orange -Rotation 30 -Label 'Rotated 30' | Out-Null
Add-Box -Slide $s -Name 'Plain2'  -X 740 -Y 330 -W 130 -H 70  -Colour $purple | Out-Null
$hint = $s.Shapes.AddTextbox($msoTextOrientationHorizontal, 40, 470, 880, 40)
$hint.TextFrame2.TextRange.Text = 'In frame mode the rotated shape looks wrong, because PowerPoint aligns an unrotated rectangle it never draws. Visual bounds fixes it.'
$hint.TextFrame2.TextRange.Font.Size = 12
$hint.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $grey

# =================================================================================================
# 3. Align to an anchor
# =================================================================================================
$s = Add-Slide -Title 'Align to an anchor' `
               -Try  'Set Reference = Anchor. Select the three small shapes, then Ctrl+click BIG ANCHOR last, and Align > Left.'
Add-Box -Slide $s -Name 'Small1' -X 120 -Y 170 -W 110 -H 60 -Colour $blue  | Out-Null
Add-Box -Slide $s -Name 'Small2' -X 250 -Y 260 -W 110 -H 60 -Colour $green | Out-Null
Add-Box -Slide $s -Name 'Small3' -X 180 -Y 350 -W 110 -H 60 -Colour $purple | Out-Null
Add-Box -Slide $s -Name 'Anchor' -X 560 -Y 200 -W 280 -H 200 -Colour $red -Label 'BIG ANCHOR' | Out-Null

# =================================================================================================
# 4. Distribute - the four modes
# =================================================================================================
$s = Add-Slide -Title 'Distribute: four modes, four answers' `
               -Try  'Select all five. Try Distribute > Horizontally with each Space by setting. Widths differ, so all four differ.'
Add-Box -Slide $s -Name 'W200' -X 50  -Y 250 -W 200 -H 90 -Colour $blue   -Label '200' | Out-Null
Add-Box -Slide $s -Name 'W60'  -X 300 -Y 250 -W 60  -H 90 -Colour $orange -Label '60'  | Out-Null
Add-Box -Slide $s -Name 'W120' -X 420 -Y 250 -W 120 -H 90 -Colour $green  -Label '120' | Out-Null
Add-Box -Slide $s -Name 'W40'  -X 620 -Y 250 -W 40  -H 90 -Colour $red    -Label '40'  | Out-Null
Add-Box -Slide $s -Name 'W150' -X 760 -Y 250 -W 150 -H 90 -Colour $purple -Label '150' | Out-Null
$hint = $s.Shapes.AddTextbox($msoTextOrientationHorizontal, 40, 380, 880, 60)
$hint.TextFrame2.TextRange.Text = 'Leading edges lines up the left edges at an even pitch. Centres evens the centres. Trailing edges evens the right edges. Space between evens the visible gaps - the only one of the four that does.'
$hint.TextFrame2.TextRange.Font.Size = 12
$hint.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $grey

# =================================================================================================
# 5. Exact spacing
# =================================================================================================
$s = Add-Slide -Title 'Exact spacing' `
               -Try  'Select all four, put 12 in the Exact (pt) box, Space by = Space between, then Distribute > Horizontally.'
for ($i = 0; $i -lt 4; $i++) {
    Add-Box -Slide $s -Name ("Gap$i") -X (80 + $i * 190) -Y 240 -W 120 -H 100 -Colour $blue -Label ('#' + ($i + 1)) | Out-Null
}
$hint = $s.Shapes.AddTextbox($msoTextOrientationHorizontal, 40, 380, 880, 40)
$hint.TextFrame2.TextRange.Text = 'Clear the box again to go back to evening out whatever space is already there. In the edge and centre modes the number is a pitch, not a gap.'
$hint.TextFrame2.TextRange.Font.Size = 12
$hint.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $grey

# =================================================================================================
# 6. Match size
# =================================================================================================
$s = Add-Slide -Title 'Match size to the anchor' `
               -Try  'Select the three odd shapes, then Ctrl+click TARGET last. Match size > Both. Try the From centre toggle too.'
Add-Box -Slide $s -Name 'Odd1' -X 110 -Y 180 -W 80  -H 130 -Colour $blue   | Out-Null
Add-Box -Slide $s -Name 'Odd2' -X 240 -Y 220 -W 160 -H 60  -Colour $green  | Out-Null
Add-Box -Slide $s -Name 'Odd3' -X 130 -Y 350 -W 110 -H 90  -Colour $purple | Out-Null
Add-Box -Slide $s -Name 'Target' -X 600 -Y 230 -W 200 -H 120 -Colour $red -Label 'TARGET' | Out-Null

# =================================================================================================
# 7. Groups
# =================================================================================================
$s = Add-Slide -Title 'Groups are one object' `
               -Try  'Align the group with the loose shapes - its internal spacing must not change. Then try Match size on it.'
$g1 = Add-Box -Slide $s -Name 'G1' -X 110 -Y 200 -W 90 -H 70 -Colour $blue
$g2 = Add-Box -Slide $s -Name 'G2' -X 230 -Y 200 -W 90 -H 70 -Colour $blue
$g3 = Add-Box -Slide $s -Name 'G3' -X 350 -Y 200 -W 90 -H 70 -Colour $blue
$group = $s.Shapes.Range(@('G1', 'G2', 'G3')).Group()
$group.Name = 'TheGroup'
Add-Box -Slide $s -Name 'Loose1' -X 600 -Y 300 -W 150 -H 90 -Colour $orange | Out-Null
Add-Box -Slide $s -Name 'Loose2' -X 790 -Y 390 -W 110 -H 60 -Colour $green  | Out-Null
$hint = $s.Shapes.AddTextbox($msoTextOrientationHorizontal, 40, 460, 880, 50)
$hint.TextFrame2.TextRange.Text = 'Match size refuses groups on purpose: scaling a group scales the gaps between its children, which silently distorts a diagram. AlignPro tells you rather than doing it.'
$hint.TextFrame2.TextRange.Font.Size = 12
$hint.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $grey

# =================================================================================================
# 8. Text bounds
# =================================================================================================
$s = Add-Slide -Title 'Align the text, not the box' `
               -Try  'All three frames already start at the same x. Select them and Align > Left with Measure = Text bounds.'
$alignments = @($msoAlignLeft, $msoAlignCenter, $msoAlignRight)
$labels = @('left-aligned text', 'centred text', 'right-aligned text')
for ($i = 0; $i -lt 3; $i++) {
    $tb = $s.Shapes.AddShape($msoShapeRoundedRectangle, 250, (170 + $i * 90), 420, 60)
    $tb.Name = "Text$i"
    $tb.Fill.ForeColor.RGB = (Rgb 235 235 240)
    $tb.Line.ForeColor.RGB = $grey
    $tb.TextFrame2.TextRange.Text = $labels[$i]
    $tb.TextFrame2.TextRange.Font.Size = 16
    $tb.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $ink
    $tb.TextFrame2.TextRange.ParagraphFormat.Alignment = $alignments[$i]
}
$hint = $s.Shapes.AddTextbox($msoTextOrientationHorizontal, 40, 450, 880, 50)
$hint.TextFrame2.TextRange.Text = 'The frames are identical, so frame mode does nothing. Text bounds moves each box so the text itself lines up.'
$hint.TextFrame2.TextRange.Font.Size = 12
$hint.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $grey

# =================================================================================================
# 9. Grid
# =================================================================================================
$s = Add-Slide -Title 'Tidy a scatter into a grid' `
               -Try  'Select all nine circles. Leave Columns blank for a near-square grid, or type 3. Then Arrange > Grid.'
$positions = @(
    @{ X = 120; Y = 150 }, @{ X = 430; Y = 190 }, @{ X = 760; Y = 140 },
    @{ X = 190; Y = 300 }, @{ X = 520; Y = 260 }, @{ X = 690; Y = 330 },
    @{ X = 110; Y = 420 }, @{ X = 400; Y = 400 }, @{ X = 800; Y = 430 }
)
for ($i = 0; $i -lt $positions.Count; $i++) {
    Add-Box -Slide $s -Name ("Dot$i") -X $positions[$i].X -Y $positions[$i].Y -W 80 -H 80 `
            -Colour $blue -Shape $msoShapeOval -Label ($i + 1) | Out-Null
}
$hint = $s.Shapes.AddTextbox($msoTextOrientationHorizontal, 40, 500, 880, 30)
$hint.TextFrame2.TextRange.Text = 'Grid only repositions - sizes are left alone, so it is safe on groups too. Set Reference = Slide to spread across the whole slide.'
$hint.TextFrame2.TextRange.Font.Size = 12
$hint.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $grey

# =================================================================================================
# 10. Slide and margin references
# =================================================================================================
$s = Add-Slide -Title 'Align to the slide, or to its margins' `
               -Try  'Select one shape. Reference = Slide then Align > Centre. Then Reference = Slide margins with Margin = 36.'
Add-Box -Slide $s -Name 'Wanderer' -X 150 -Y 200 -W 220 -H 120 -Colour $orange -Label 'Move me' | Out-Null
Add-Box -Slide $s -Name 'Friend'   -X 620 -Y 340 -W 160 -H 90  -Colour $green   | Out-Null
$hint = $s.Shapes.AddTextbox($msoTextOrientationHorizontal, 40, 470, 880, 50)
$hint.TextFrame2.TextRange.Text = 'Aligning a single shape to the selection is meaningless, so AlignPro refuses it and says why. Against the slide it is perfectly meaningful, so it works.'
$hint.TextFrame2.TextRange.Font.Size = 12
$hint.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $grey

# --- save, close, reopen --------------------------------------------------------------------------
if (Test-Path $Path) { Remove-Item $Path -Force }
$pres.SaveAs($Path, $ppSaveAsOpenXMLPresentation)
$pres.Close()

# Reopening starts with an empty undo history, so Ctrl+Z behaves from the very first click rather
# than sitting on top of one huge entry for everything this script just did.
$reopened = $ppt.Presentations.Open($Path)
$ppt.ActiveWindow.View.GotoSlide(1)

Write-Host ''
Write-Host ("Sample deck saved and reopened: {0}" -f $Path) -ForegroundColor Green
Write-Host ("{0} slides. Slide 2 is the one to try first." -f $reopened.Slides.Count) -ForegroundColor Cyan
Write-Host 'Undo history is clean, so Ctrl+Z is safe from the first click.' -ForegroundColor DarkGray
Write-Host ''
