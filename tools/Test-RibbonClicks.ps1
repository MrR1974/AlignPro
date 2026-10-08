<#
.SYNOPSIS
    Tests AlignPro by clicking its ribbon buttons for real, through UI Automation.

.DESCRIPTION
    Test-AlignProEndToEnd.ps1 drives the add-in through its automation surface, over cross-process COM.
    That is fast and precise, but it is NOT the context a ribbon click runs in, and the difference is
    not academic: a bug shipped where Align Left silently did nothing when clicked, while every check
    in that harness passed. PowerPoint defers some operations while a command is executing, and an
    inbound COM call has no command in flight.

    This harness closes that gap. Setup and assertions still go through COM - selecting shapes, reading
    geometry - but the operation under test is invoked by genuinely clicking the ribbon button, so
    Office dispatches it exactly as it would for a person.

    What it cannot do is replace the other harness. Clicking is slow and depends on UI structure, so
    this covers the handful of paths where the calling context matters, and the geometry stays covered
    by the unit tests and the COM harness.

    Deliberately avoids operations that raise a message box: a modal dialog blocks PowerPoint's UI
    thread and would hang the run. Those paths are covered through the automation surface instead.

.EXAMPLE
    .\Test-RibbonClicks.ps1
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

$A  = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]

$tolerance = 0.75
$results = [System.Collections.Generic.List[object]]::new()

function Add-Result {
    param([string] $Case, [bool] $Passed, [string] $Detail)
    $results.Add([pscustomobject]@{ Case = $Case; Result = $(if ($Passed) { 'PASS' } else { 'FAIL' }); Detail = $Detail })
}

# --- PowerPoint, restarted so the current build is loaded and the deck is pristine ----------------
function Restart-PowerPointWithSample {
    try {
        $existing = New-Object -ComObject PowerPoint.Application
        for ($i = $existing.Presentations.Count; $i -ge 1; $i--) { $existing.Presentations.Item($i).Saved = -1 }
        $existing.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($existing)
    } catch { }
    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 500 }

    & (Join-Path $PSScriptRoot 'New-SampleDeck.ps1') -Force | Out-Null
    Start-Sleep -Seconds 3
}

function Get-PowerPointWindow {
    $proc = Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $proc) { throw 'PowerPoint is not running.' }
    $window = $A::RootElement.FindFirst($TS::Children,
        (New-Object System.Windows.Automation.PropertyCondition($A::ProcessIdProperty, $proc.Id)))
    if (-not $window) { throw 'Could not reach the PowerPoint window through UI Automation.' }
    $window
}

<#
    Maximises the PowerPoint window.

    Not cosmetic. PowerPoint started over COM opens at roughly 500x400, and at that width Office
    collapses every ribbon group into a single drop-down - so the individual buttons this harness
    clicks are not in the UI Automation tree at all, and every case fails with "no button matching
    'Left' is visible", which reads like the add-in failed to load. Whether the window happens to be
    big enough is otherwise down to whatever size PowerPoint last remembered.
#>
function Expand-PowerPointWindow {
    param($Window)
    try {
        $pattern = $Window.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
        $pattern.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Maximized)
        Start-Sleep -Milliseconds 800
    }
    catch {
        Write-Host "  could not maximise the window: $($_.Exception.Message)" -ForegroundColor DarkYellow
    }

    $width = $Window.Current.BoundingRectangle.Width
    if ($width -lt 1100) {
        throw ("The PowerPoint window is only {0:F0}px wide. The ribbon collapses below about 1100px, " +
               "which hides the buttons this harness clicks." -f $width)
    }
}

function Select-AlignProTab {
    param($Window)
    $tab = $Window.FindFirst($TS::Descendants, (New-Object System.Windows.Automation.AndCondition(
        (New-Object System.Windows.Automation.PropertyCondition($A::ControlTypeProperty, $CT::TabItem)),
        (New-Object System.Windows.Automation.PropertyCondition($A::NameProperty, 'AlignPro')))))
    if (-not $tab) { throw 'The AlignPro ribbon tab is not present - is the add-in loaded?' }
    $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
    Start-Sleep -Milliseconds 600
}

<#
    Clicks a ribbon button by label pattern. Elements are looked up fresh every time, because the
    ribbon rebuilds them and a cached reference goes stale.

    Matching is by pattern and prefers an enabled control, because a label can also belong to one of
    PowerPoint's own buttons elsewhere in the window.
#>
function Invoke-RibbonButton {
    param($Window, [string] $Pattern)

    $buttons = $Window.FindAll($TS::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition($A::ControlTypeProperty, $CT::Button)))

    $matched = @()
    foreach ($b in $buttons) { if ($b.Current.Name -like $Pattern) { $matched += $b } }
    if ($matched.Count -eq 0) { throw "No ribbon button matching '$Pattern' is visible. Is the AlignPro tab selected?" }

    $button = $matched | Where-Object { $_.Current.IsEnabled } | Select-Object -First 1
    if (-not $button) { throw "Every button matching '$Pattern' is disabled ($($matched.Count) found)." }

    $button.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
    Start-Sleep -Milliseconds 700
    $button.Current.Name
}

<#
    A modal dialog would block PowerPoint's UI thread and hang everything after it, so every check
    that might raise one calls this, and it is dismissed either way so the rest of the run continues.
    Returns the dialog's title and message, or $null when there was none.

    Two things about AlignPro's message boxes make this harder than it looks, and both once hung the
    run. The box is not a direct child of PowerPoint's window but sits deeper, so it is searched for
    among all descendants, by the standard dialog class. And its OK button reports itself to UI
    Automation as a pane, not a button, and offers no Invoke - so it is found by class and name, and
    when it cannot be invoked the dialog is brought forward and Enter pressed instead.
#>
function Get-BlockingDialog {
    param($Window)
    Start-Sleep -Milliseconds 300
    $dialog = $Window.FindFirst($TS::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition($A::ClassNameProperty, '#32770')))
    if (-not $dialog) { return $null }

    $title = $dialog.Current.Name
    $message = @()
    $statics = $dialog.FindAll($TS::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition($A::ClassNameProperty, 'Static')))
    foreach ($st in $statics) { if ($st.Current.Name) { $message += $st.Current.Name } }

    $ok = $dialog.FindFirst($TS::Descendants, (New-Object System.Windows.Automation.AndCondition(
        (New-Object System.Windows.Automation.PropertyCondition($A::ClassNameProperty, 'Button')),
        (New-Object System.Windows.Automation.PropertyCondition($A::NameProperty, 'OK')))))
    $invoke = $null
    if ($ok -and $ok.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$invoke)) {
        $invoke.Invoke()
    }
    else {
        $shell = New-Object -ComObject WScript.Shell
        [void]$shell.AppActivate($title)
        Start-Sleep -Milliseconds 300
        $shell.SendKeys('{ENTER}')
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
    }
    Start-Sleep -Milliseconds 500
    ($title + ': ' + ($message -join ' '))
}

<#
    Presses Ctrl+Z as a person would. AlignPro has no undo of its own: this is the only undo there is.

    This is the case UndoBoundary exists for. PowerPoint brackets each of its own commands in an undo
    entry, but object-model writes accumulate into one open entry, so without a boundary a single
    Ctrl+Z reverses every AlignPro operation since the last native edit - and any automation before
    them. StartNewUndoEntry closes that entry, so each Ctrl+Z reverses exactly one AlignPro operation.
#>
function Send-NativeUndo {
    param($Window)

    # AppActivate rather than the element's SetFocus: the top-level PowerPoint window reports
    # "Target element cannot receive focus" through UI Automation, because focus belongs to the
    # editing surface inside it.
    $proc = Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $proc) { throw 'PowerPoint is not running.' }

    $shell = New-Object -ComObject WScript.Shell
    if (-not $shell.AppActivate($proc.Id)) {
        Start-Sleep -Milliseconds 500
        [void]$shell.AppActivate($proc.Id)
    }
    Start-Sleep -Milliseconds 500

    # Esc first, and it is load-bearing. Invoking a ribbon button through UI Automation leaves
    # keyboard focus on that button, and Ctrl+Z sent in that state never reaches the document - it
    # silently does nothing, which reads exactly like a broken undo. Esc returns focus to the slide.
    $shell.SendKeys('{ESC}')
    Start-Sleep -Milliseconds 400
    $shell.SendKeys('^z')
    Start-Sleep -Milliseconds 1400
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
}

# --- COM side: setup and assertions ---------------------------------------------------------------
Write-Host ''
Write-Host 'Restarting PowerPoint with a pristine sample deck...' -ForegroundColor DarkGray
Restart-PowerPointWithSample

$ppt = New-Object -ComObject PowerPoint.Application
$addIn = $null
for ($i = 1; $i -le $ppt.COMAddIns.Count; $i++) {
    $candidate = $ppt.COMAddIns.Item($i)
    if ($candidate.Description -like '*AlignPro*') { $addIn = $candidate }
}
if (-not $addIn -or -not $addIn.Connect) { throw 'AlignPro is not loaded in PowerPoint.' }
$api = $addIn.Object

# Settings persist across slides, so pin them rather than inheriting whatever was last used by hand.
[void]$api.SetReference('SelectionBounds')
[void]$api.SetBoundsModel('ShapeFrame')
[void]$api.SetExactSpacing('')
[void]$api.SetGridColumns('')
[void]$api.SetSizeMargin('')
[void]$api.SetSizeMarginMode('None')

$window = Get-PowerPointWindow
Expand-PowerPointWindow -Window $window
Select-AlignProTab -Window $window
Write-Host "Driving the ribbon of: $($window.Current.Name)" -ForegroundColor DarkGray

$presentation = $ppt.ActivePresentation

function Select-Shapes {
    param([int] $Slide, [string[]] $Names)
    $ppt.ActiveWindow.View.GotoSlide($Slide)
    Start-Sleep -Milliseconds 250
    $presentation.Slides.Item($Slide).Shapes.Range($Names).Select(-1)
    Start-Sleep -Milliseconds 250
}

function Get-Lefts { param([int] $Slide, [string[]] $Names)
    $map = @{}; foreach ($n in $Names) { $map[$n] = [double]$presentation.Slides.Item($Slide).Shapes.Item($n).Left }; $map }

function Get-Tops { param([int] $Slide, [string[]] $Names)
    $map = @{}; foreach ($n in $Names) { $map[$n] = [double]$presentation.Slides.Item($Slide).Shapes.Item($n).Top }; $map }

# =================================================================================================
# The regression this harness exists for: does clicking actually do the work?
# =================================================================================================
$slide2 = @('Plain1', 'Tall', 'Rotated', 'Plain2')
$before = Get-Lefts -Slide 2 -Names $slide2
Select-Shapes -Slide 2 -Names $slide2
$clicked = Invoke-RibbonButton -Window $window -Pattern 'Left'

$dialog = Get-BlockingDialog -Window $window
if ($dialog) { Add-Result 'No unexpected dialog' $false $dialog } else { Add-Result 'No unexpected dialog' $true '' }

$afterLeft = Get-Lefts -Slide 2 -Names $slide2
$moved = @($slide2 | Where-Object { [Math]::Abs($before[$_] - $afterLeft[$_]) -gt $tolerance }).Count
Add-Result 'Clicking Align Left moves shapes' ($moved -gt 0) "$moved of $($slide2.Count) shapes moved"

$target = ($afterLeft.Values | Measure-Object -Minimum).Minimum
$aligned = @($slide2 | Where-Object { [Math]::Abs($afterLeft[$_] - $target) -le $tolerance }).Count
Add-Result 'They all land on one edge' ($aligned -eq $slide2.Count) "$aligned of $($slide2.Count) at left=$('{0:F0}' -f $target)"

# =================================================================================================
# A second click must apply on top of the first, not replace or undo it
# =================================================================================================
$topsBefore = Get-Tops -Slide 2 -Names $slide2
Select-Shapes -Slide 2 -Names $slide2
$clicked = Invoke-RibbonButton -Window $window -Pattern 'Top'

$afterTop = Get-Tops -Slide 2 -Names $slide2
$topsMoved = @($slide2 | Where-Object { [Math]::Abs($topsBefore[$_] - $afterTop[$_]) -gt $tolerance }).Count
$leftsHeld = @($slide2 | Where-Object { [Math]::Abs((Get-Lefts -Slide 2 -Names $slide2)[$_] - $afterLeft[$_]) -le $tolerance }).Count
Add-Result 'A second click stacks on the first' (($topsMoved -gt 0) -and ($leftsHeld -eq $slide2.Count)) `
    "$topsMoved moved vertically, $leftsHeld kept their new left"

# =================================================================================================
# Ctrl+Z, pressed - must reverse one AlignPro operation at a time
# =================================================================================================
Send-NativeUndo -Window $window
$dialog = Get-BlockingDialog -Window $window
if ($dialog) { Add-Result 'Ctrl+Z raises no dialog' $false $dialog } else { Add-Result 'Ctrl+Z raises no dialog' $true '' }

$ctrlZTops = Get-Tops -Slide 2 -Names $slide2
$ctrlZLefts = Get-Lefts -Slide 2 -Names $slide2
$topsBack = @($slide2 | Where-Object { [Math]::Abs($ctrlZTops[$_] - $topsBefore[$_]) -le $tolerance }).Count
$leftsSurvived = @($slide2 | Where-Object { [Math]::Abs($ctrlZLefts[$_] - $afterLeft[$_]) -le $tolerance }).Count
Add-Result 'Ctrl+Z reverses only the last operation' (($topsBack -eq $slide2.Count) -and ($leftsSurvived -eq $slide2.Count)) `
    "$topsBack of $($slide2.Count) tops restored, $leftsSurvived kept their align-left"

Send-NativeUndo -Window $window
$secondLefts = Get-Lefts -Slide 2 -Names $slide2
$leftsReverted = @($slide2 | Where-Object { [Math]::Abs($secondLefts[$_] - $before[$_]) -le $tolerance }).Count
Add-Result 'A second Ctrl+Z reverses the first operation' ($leftsReverted -eq $slide2.Count) `
    "$leftsReverted of $($slide2.Count) back where they started"

# The deck itself must survive. Before StartNewUndoEntry, the scripted deck build shared one undo
# entry with the AlignPro operations, and a single Ctrl+Z emptied the presentation.
$slideCount = $presentation.Slides.Count
Add-Result 'Ctrl+Z leaves the deck intact' ($slideCount -ge 17) "$slideCount slides still present"

# =================================================================================================
# Distribute and Grid, clicked - the other two verbs with their own code paths
# =================================================================================================
$slide4 = @('W200', 'W60', 'W120', 'W40', 'W150')
$distBefore = Get-Lefts -Slide 4 -Names $slide4
Select-Shapes -Slide 4 -Names $slide4
$clicked = Invoke-RibbonButton -Window $window -Pattern 'Horizontally'
$distAfter = Get-Lefts -Slide 4 -Names $slide4
$distMoved = @($slide4 | Where-Object { [Math]::Abs($distBefore[$_] - $distAfter[$_]) -gt $tolerance }).Count
Add-Result 'Clicking Distribute moves shapes' ($distMoved -gt 0) "$distMoved of $($slide4.Count) shapes moved"

# Slide 13 in the sample deck: the Grid slide.
$gridShapes = 0..8 | ForEach-Object { "Dot$_" }
Select-Shapes -Slide 13 -Names $gridShapes
$clicked = Invoke-RibbonButton -Window $window -Pattern 'Grid'
$gridLefts = Get-Lefts -Slide 13 -Names $gridShapes
$columns = @($gridLefts.Values | ForEach-Object { [Math]::Round($_, 0) } | Sort-Object -Unique).Count
Add-Result 'Clicking Grid forms three columns' ($columns -eq 3) "$columns distinct columns"

# =================================================================================================
# Order, clicked - the path where a deferred call would be invisible from the COM harness
# =================================================================================================
# Selection ORDER is the whole input here, and Range(...).Select() says nothing about order, so the
# shapes are clicked one at a time exactly as a person would.
function Select-ShapesInOrder {
    param([int] $Slide, [string[]] $Names)
    $ppt.ActiveWindow.View.GotoSlide($Slide)
    Start-Sleep -Milliseconds 250
    $shapes = $presentation.Slides.Item($Slide).Shapes
    $first = $true
    foreach ($n in $Names) {
        # Replace on the first, extend thereafter.
        $shapes.Item($n).Select($(if ($first) { -1 } else { 0 }))
        $first = $false
        Start-Sleep -Milliseconds 120
    }
    Start-Sleep -Milliseconds 250
}

function Get-ZOrder { param([int] $Slide, [string] $Name)
    [int]$presentation.Slides.Item($Slide).Shapes.Item($Name).ZOrderPosition }

$orderSlide = 7
$barABefore = Get-ZOrder -Slide $orderSlide -Name 'BarA'
$barBBefore = Get-ZOrder -Slide $orderSlide -Name 'BarB'

Select-ShapesInOrder -Slide $orderSlide -Names @('Card1', 'Card2', 'Card3')
$clicked = Invoke-RibbonButton -Window $window -Pattern 'Stack'

$dialog = Get-BlockingDialog -Window $window
if ($dialog) { Add-Result 'Order raises no dialog' $false $dialog } else { Add-Result 'Order raises no dialog' $true '' }

$z1 = Get-ZOrder -Slide $orderSlide -Name 'Card1'
$z2 = Get-ZOrder -Slide $orderSlide -Name 'Card2'
$z3 = Get-ZOrder -Slide $orderSlide -Name 'Card3'

# Card1 was clicked first, so it must end above the other two. This is the assertion that would fail
# if Shape.ZOrder were deferred the way ExecuteMso is inside a ribbon callback.
Add-Result 'Clicking Stack puts the first selected on top' (($z1 -gt $z2) -and ($z2 -gt $z3)) `
    "Card1=$z1 Card2=$z2 Card3=$z3 (higher is nearer the front)"

$barAAfter = Get-ZOrder -Slide $orderSlide -Name 'BarA'
$barBAfter = Get-ZOrder -Slide $orderSlide -Name 'BarB'
Add-Result 'Unselected shapes keep their layer' `
    (($barABefore -eq $barAAfter) -and ($barBBefore -eq $barBAfter)) `
    "BarA $barABefore->$barAAfter, BarB $barBBefore->$barBAfter"

$clicked = Invoke-RibbonButton -Window $window -Pattern 'Reverse'
$r1 = Get-ZOrder -Slide $orderSlide -Name 'Card1'
$r3 = Get-ZOrder -Slide $orderSlide -Name 'Card3'
Add-Result 'Clicking Reverse flips the stack' ($r3 -gt $r1) "Card1=$r1 Card3=$r3"

Send-NativeUndo -Window $window
$u1 = Get-ZOrder -Slide $orderSlide -Name 'Card1'
$u3 = Get-ZOrder -Slide $orderSlide -Name 'Card3'
Add-Result 'Ctrl+Z restores the previous stacking' ($u1 -gt $u3) "Card1=$u1 Card3=$u3"

# =================================================================================================
# The 1.3 verbs, clicked. Settings go through COM - dropdowns and edit boxes are not what is under
# test - and the button that runs the verb is clicked for real.
# =================================================================================================
function Get-Shape { param([int] $Slide, [string] $Name) $presentation.Slides.Item($Slide).Shapes.Item($Name) }

# Slide 10: Grow with Cascade. CORE is 120x80; 1st is two steps of 12 out, 2nd one step.
[void]$api.SetSizeMargin('12')
[void]$api.SetSizeDirection('Grow')
[void]$api.SetSizeMarginMode('Cascade')
Select-ShapesInOrder -Slide 10 -Names @('Grow1', 'Grow2', 'Grow3', 'Core')
$clicked = Invoke-RibbonButton -Window $window -Pattern 'Both'
$g1 = Get-Shape 10 'Grow1'; $g3 = Get-Shape 10 'Grow3'
Add-Result 'Clicking Both with Grow grows the first selected most' `
    (([Math]::Abs($g1.Width - 192) -le $tolerance) -and ([Math]::Abs($g3.Width - 144) -le $tolerance)) `
    ("1st {0:F0} wide, 3rd {1:F0}, expected 192 and 144" -f $g1.Width, $g3.Width)
[void]$api.SetSizeDirection('Shrink')
[void]$api.SetSizeMarginMode('None')
[void]$api.SetSizeMargin('')

# Slide 15: Match rotation, then Ctrl+Z.
Select-ShapesInOrder -Slide 15 -Names @('Tilt1', 'Tilt2', 'Tilt3', 'Level')
$clicked = Invoke-RibbonButton -Window $window -Pattern 'Rotation'
$angles = @('Tilt1', 'Tilt2', 'Tilt3') | ForEach-Object { [double](Get-Shape 15 $_).Rotation }
Add-Result 'Clicking Rotation turns every shape to the anchor' `
    (@($angles | Where-Object { [Math]::Abs($_ - 20) -le 0.1 }).Count -eq 3) ("angles: " + ($angles -join ', '))
Send-NativeUndo -Window $window
Add-Result 'Ctrl+Z restores the angles' ([Math]::Abs((Get-Shape 15 'Tilt1').Rotation - 12) -le 0.1) `
    ("Tilt1 back to {0:F1}, was 12" -f (Get-Shape 15 'Tilt1').Rotation)

# Slide 16: Duplicate round the slide centre, then Ctrl+Z.
$dupSlide = $presentation.Slides.Item(16)
$countBefore = $dupSlide.Shapes.Count
[void]$api.SetDuplicate(0, 0, 30, 11, 'SlideCentre')
[void]$api.SetRotateShapes($true)
Select-ShapesInOrder -Slide 16 -Names @('Seed')
$clicked = Invoke-RibbonButton -Window $window -Pattern 'Duplicate'
$dialog = Get-BlockingDialog -Window $window
if ($dialog) { Add-Result 'Duplicate raises no dialog' $false $dialog }
Add-Result 'Clicking Duplicate makes eleven copies' ($dupSlide.Shapes.Count -eq $countBefore + 11) `
    ("{0} shapes, expected {1}" -f $dupSlide.Shapes.Count, ($countBefore + 11))
$selected = $ppt.ActiveWindow.Selection.ShapeRange.Count
Add-Result 'The original and copies are left selected' ($selected -eq 12) "$selected selected"
Send-NativeUndo -Window $window
Add-Result 'Ctrl+Z removes every copy and nothing else' `
    (($dupSlide.Shapes.Count -eq $countBefore) -and ($presentation.Slides.Count -ge 17)) `
    ("{0} shapes, {1} slides" -f $dupSlide.Shapes.Count, $presentation.Slides.Count)

# Slide 17: round the ring. RING is 240x240 at (70, 120), so twelve o'clock is (190, 120).
Select-ShapesInOrder -Slide 17 -Names @('Bead1', 'Bead2', 'Bead3', 'Bead4', 'Ring')
$clicked = Invoke-RibbonButton -Window $window -Pattern 'Along curve'
$b1 = Get-Shape 17 'Bead1'; $b2 = Get-Shape 17 'Bead2'
$ok = ([Math]::Abs($b1.Left + 15 - 190) -le $tolerance) -and ([Math]::Abs($b1.Top + 15 - 120) -le $tolerance) -and
      ([Math]::Abs($b2.Left + 15 - 310) -le $tolerance) -and ([Math]::Abs($b2.Rotation - 90) -le 0.6)
Add-Result 'Clicking Along curve rings the beads' $ok `
    ("Bead1 centre ({0:F1},{1:F1}), Bead2 ({2:F1},{3:F1}) rot {4:F1}" -f ($b1.Left + 15), ($b1.Top + 15), ($b2.Left + 15), ($b2.Top + 15), $b2.Rotation)
Send-NativeUndo -Window $window
Add-Result 'Ctrl+Z takes the beads back' ([Math]::Abs((Get-Shape 17 'Bead2').Rotation) -le 0.1) `
    ("Bead2 rotation {0:F1}" -f (Get-Shape 17 'Bead2').Rotation)

# Slide 18: shapes picked inside a group, then a real click on Left.
$panel = $presentation.Slides.Item(18).Shapes.Item('Panel')
$items = @('Item1', 'Item2', 'Item3')
[void]$api.SetReference('SelectionBounds')
[void]$api.SetBoundsModel('ShapeFrame')
$ppt.ActiveWindow.View.GotoSlide(18)
Start-Sleep -Milliseconds 250
$first = $true
foreach ($n in $items) {
    $panel.GroupItems.Item($n).Select($(if ($first) { -1 } else { 0 }))
    $first = $false
    Start-Sleep -Milliseconds 120
}
$clicked = Invoke-RibbonButton -Window $window -Pattern 'Left'
$dialog = Get-BlockingDialog -Window $window
if ($dialog) { Add-Result 'Inside a group raises no dialog' $false $dialog }
$lefts = $items | ForEach-Object { [Math]::Round($panel.GroupItems.Item($_).Left, 1) }
$cardLeft = $panel.GroupItems.Item('Card').Left
$ok = (@($lefts | Sort-Object -Unique).Count -eq 1) -and ($lefts[0] -eq 260) -and ([Math]::Abs($cardLeft - 220) -le $tolerance) -and
      ($presentation.Slides.Item(18).Shapes.Count -eq 4)
Add-Result 'Clicking Left inside a group aligns only the picked shapes' $ok `
    ("lefts {0}; card left {1:F0}; group intact" -f ($lefts -join ', '), $cardLeft)
Send-NativeUndo -Window $window
Add-Result 'Ctrl+Z puts them back' ([Math]::Abs($panel.GroupItems.Item('Item2').Left - 330) -le $tolerance) `
    ("Item2 left {0:F1}, was 330" -f $panel.GroupItems.Item('Item2').Left)

# Stack inside the group. Item1 was made before Item3, so it starts below it; picked first, Stack
# must lift it on top. The card behind them, not picked, keeps its layer.
$cardZ = $panel.GroupItems.Item('Card').ZOrderPosition
$first = $true
foreach ($n in @('Item1', 'Item3')) {
    $panel.GroupItems.Item($n).Select($(if ($first) { -1 } else { 0 }))
    $first = $false
    Start-Sleep -Milliseconds 120
}
$clicked = Invoke-RibbonButton -Window $window -Pattern 'Stack'
$dialog = Get-BlockingDialog -Window $window
if ($dialog) { Add-Result 'Stack inside a group raises no dialog' $false $dialog }
$z1 = $panel.GroupItems.Item('Item1').ZOrderPosition
$z3 = $panel.GroupItems.Item('Item3').ZOrderPosition
Add-Result 'Clicking Stack inside a group restacks within it' (($z1 -gt $z3) -and ($panel.GroupItems.Item('Card').ZOrderPosition -eq $cardZ)) `
    ("Item1 z={0} Item3 z={1}; card z {2} -> {3}" -f $z1, $z3, $cardZ, $panel.GroupItems.Item('Card').ZOrderPosition)
Send-NativeUndo -Window $window

# Duplicate inside the group: the copy lands in the group, and the group stays one shape on the slide.
# The selection is read rather than GroupItems, which PowerShell reads stale after a copy (probe 23).
[void]$api.SetDuplicate(0, 70, 0, 1, 'OwnCentre')
$panel.GroupItems.Item('Item1').Select(-1)
Start-Sleep -Milliseconds 150
$clicked = Invoke-RibbonButton -Window $window -Pattern 'Duplicate'
$dialog = Get-BlockingDialog -Window $window
if ($dialog) { Add-Result 'Duplicate inside a group raises no dialog' $false $dialog }
$sel = $ppt.ActiveWindow.Selection
$picked = 0; $parent = ''
if ($sel.HasChildShapeRange) { $picked = $sel.ChildShapeRange.Count; $parent = $sel.ChildShapeRange.Item(2).ParentGroup.Name }
Add-Result 'Clicking Duplicate inside a group copies into the group' (($picked -eq 2) -and ($parent -eq 'Panel') -and ($presentation.Slides.Item(18).Shapes.Count -eq 4)) `
    ("{0} picked, copy's group '{1}', {2} top-level shapes" -f $picked, $parent, $presentation.Slides.Item(18).Shapes.Count)
Send-NativeUndo -Window $window

# Slide 19: inside a group turned 15 degrees, Left must follow the group's edge.
$tiltPanel = $presentation.Slides.Item(19).Shapes.Item('TiltPanel')
$tiltItems = @('TiltItem1', 'TiltItem2', 'TiltItem3')
$ppt.ActiveWindow.View.GotoSlide(19)
Start-Sleep -Milliseconds 250
$first = $true
foreach ($n in $tiltItems) {
    $tiltPanel.GroupItems.Item($n).Select($(if ($first) { -1 } else { 0 }))
    $first = $false
    Start-Sleep -Milliseconds 120
}
$clicked = Invoke-RibbonButton -Window $window -Pattern 'Left'
$dialog = Get-BlockingDialog -Window $window
if ($dialog) { Add-Result 'Inside a rotated group raises no dialog' $false $dialog }
$t = 15 * [Math]::PI / 180
$along = $tiltItems | ForEach-Object {
    $c = $tiltPanel.GroupItems.Item($_)
    [Math]::Round((($c.Left + $c.Width / 2) * [Math]::Cos($t) + ($c.Top + $c.Height / 2) * [Math]::Sin($t)) - $c.Width / 2, 1)
}
Add-Result 'Clicking Left inside a rotated group follows its edge' (@($along | Sort-Object -Unique).Count -eq 1) `
    ("along-group lefts {0}" -f ($along -join ', '))
Send-NativeUndo -Window $window

# Slide 20: Tidy, clicked. It always reports, so a dialog is expected here - and it must be Tidy's.
$tidySlide = $presentation.Slides.Item(20)
$cells = 1..9 | ForEach-Object { "Cell$_" }
$steps = 1..4 | ForEach-Object { "Step$_" }
[void]$api.SetBoundsModel('ShapeFrame')
[void]$api.SetTidyTolerance('')
$cell2Before = $tidySlide.Shapes.Item('Cell2').Left
Select-Shapes -Slide 20 -Names ($cells + $steps + @('StagA', 'StagB', 'Link'))
$clicked = Invoke-RibbonButton -Window $window -Pattern 'Tidy'
$dialog = Get-BlockingDialog -Window $window
Add-Result 'Clicking Tidy reports what it did' ($dialog -like 'AlignPro - Tidy: Tidied*') ("dialog: " + $dialog)

function Get-TidyShape { param([string] $Name) $tidySlide.Shapes.Item($Name) }
$columnsAligned = 0
for ($c = 0; $c -lt 3; $c++) {
    $lefts = @(); for ($r = 0; $r -lt 3; $r++) { $lefts += [Math]::Round((Get-TidyShape "Cell$($r * 3 + $c + 1)").Left, 2) }
    if (@($lefts | Sort-Object -Unique).Count -eq 1) { $columnsAligned++ }
}
$rowsAligned = 0
for ($r = 0; $r -lt 3; $r++) {
    $tops = @(); for ($c = 0; $c -lt 3; $c++) { $tops += [Math]::Round((Get-TidyShape "Cell$($r * 3 + $c + 1)").Top, 2) }
    if (@($tops | Sort-Object -Unique).Count -eq 1) { $rowsAligned++ }
}
$gaps = @(); for ($i = 1; $i -le 3; $i++) { $gaps += [Math]::Round((Get-TidyShape "Step$($i + 1)").Left - (Get-TidyShape "Step$i").Left - 50, 2) }
$staggerKept = ([Math]::Abs((Get-TidyShape 'StagA').Left - 520) -le 0.01) -and ([Math]::Abs((Get-TidyShape 'StagB').Left - 526) -le 0.01)
Add-Result 'Tidy squares the grid, evens the row, keeps the stagger' `
    (($columnsAligned -eq 3) -and ($rowsAligned -eq 3) -and (@($gaps | Sort-Object -Unique).Count -eq 1) -and $staggerKept) `
    ("{0}/3 columns, {1}/3 rows, row gaps {2}, stagger kept {3}" -f $columnsAligned, $rowsAligned, ($gaps -join '/'), $staggerKept)

Send-NativeUndo -Window $window
Add-Result 'Ctrl+Z undoes the whole tidy' ([Math]::Abs($tidySlide.Shapes.Item('Cell2').Left - $cell2Before) -le $tolerance) `
    ("Cell2 left {0:F2}, was {1:F2}" -f $tidySlide.Shapes.Item('Cell2').Left, $cell2Before)

# =================================================================================================
Write-Host ''
$results | Format-Table -AutoSize
$failed = @($results | Where-Object { $_.Result -eq 'FAIL' }).Count
if ($failed -eq 0) {
    Write-Host ("All {0} ribbon-click checks passed." -f $results.Count) -ForegroundColor Green
}
else {
    Write-Host ("{0} of {1} ribbon-click checks FAILED." -f $failed, $results.Count) -ForegroundColor Red
}
Write-Host ''
exit $failed
