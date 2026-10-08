<#
.SYNOPSIS
    Measures the PowerPoint behaviours that working on shapes inside a group depends on.

.DESCRIPTION
    The spike for "tools inside a group" (docs/development.md). Numbered on from
    Probe-DuplicateAndPaths.ps1's 7 to 10:

      11. What does the selection report when shapes inside a group are selected - in what order,
          and can shapes from two groups, or a grouped and a loose shape, be selected together?
      12. In an unrotated group, are a child's Left/Top/Width/Height slide coordinates, are they
          writable, and what happens to the group's frame and the other children?
      13. Does one Ctrl+Z reverse child writes made after StartNewUndoEntry, group frame included?
      14. In a rotated group, what does a child report, and where does a write to it really go?
          Does moving one child shift its siblings on the slide?
      15. In a rotated group, are a child's TextRange2.Bound* values in slide space?
      16. How do nested groups appear in GroupItems and ParentGroup, and in the selection?
      17. Where does Duplicate put the copy of a child?
      18. How does ZOrder behave on a child, and is GroupItems indexed in z-order?
      19. Inside a nested group, does ZOrder on a leaf restack it within the inner group only, and
          does anything in the object model say which inner group a leaf belongs to?
      20. Probe 17 again, properly: after Duplicate on a child, is the copy really in the group? Every
          view of it is read fresh, the group's true contents are found by ungrouping a copy of it,
          and the alternatives - duplicating the child range, ungroup and regroup - are measured.
      21. Does GroupItems ever catch up with a copy made inside the group - after selecting it,
          a new undo entry, a nudge of the group, a save, a reopen? And while it has not, can the
          copy be reached through the selection instead?
      22. Does PowerPoint's own Duplicate command (Ctrl+D) inside a group make the same copy that
          GroupItems cannot see - so decks people already have may hold such shapes?
      23. Does deselecting everything and then selecting the group again make GroupItems list the
          copy - by script, via another shape, and from the keyboard as a person would?

    Where a shape really sits is measured, not trusted: the group is duplicated and the duplicate
    ungrouped, which bakes the group's rotation into each shape, and the loose shapes are read.

    Safe by design, as the other probes: works only in a new presentation, never saves, and only
    quits PowerPoint if it started it.

.PARAMETER KeepOpen
    Leave the scratch presentation open at the end.

.PARAMETER OutFile
    Also write the findings to this path as Markdown, for pasting into object-model-findings.md.

.EXAMPLE
    .\Probe-Groups.ps1 -OutFile $env:TEMP\group-findings.md
#>
[CmdletBinding()]
param(
    [switch] $KeepOpen,
    [string] $OutFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$msoTrue           = -1
$msoFalse          = 0
$msoShapeRectangle = 1
$msoBringToFront   = 0
$ppLayoutBlank     = 12

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
    param($Shape, [string] $Label = $Shape.Name)
    '{0,-12} L={1,8:F2} T={2,8:F2} W={3,7:F2} H={4,7:F2} Rot={5,6:F1} centre=({6:F2},{7:F2})' -f `
        $Label, $Shape.Left, $Shape.Top, $Shape.Width, $Shape.Height, $Shape.Rotation,
        ($Shape.Left + $Shape.Width / 2), ($Shape.Top + $Shape.Height / 2)
}

function New-Rect {
    param($Slide, [string] $Name, [double] $X, [double] $Y, [double] $W = 60, [double] $H = 40, [string] $Text)
    $s = $Slide.Shapes.AddShape($msoShapeRectangle, $X, $Y, $W, $H)
    $s.Name = $Name
    if ($Text) { $s.TextFrame.TextRange.Text = $Text }
    $s
}

function New-Group {
    param($Slide, [string] $Name, [string[]] $Members)
    $g = $Slide.Shapes.Range($Members).Group()
    $g.Name = $Name
    $g
}

# Where each child of a group really is: duplicate the group in place, ungroup the copy, read the
# loose shapes, delete them. Returns name -> @{ CX; CY; Rot; W; H }.
function Get-Truth {
    param($Group)
    $dup = $Group.Duplicate().Item(1)
    $dup.Left = $Group.Left
    $dup.Top = $Group.Top
    $loose = $dup.Ungroup()
    $truth = @{}
    $names = @()
    for ($i = 1; $i -le $loose.Count; $i++) { $names += $loose.Item($i).Name }
    # Copies take new names, so match by position in GroupItems order.
    for ($i = 1; $i -le $loose.Count; $i++) {
        $s = $loose.Item($i)
        $original = $Group.GroupItems.Item($i).Name
        $truth[$original] = @{ CX = $s.Left + $s.Width / 2; CY = $s.Top + $s.Height / 2; Rot = $s.Rotation; W = $s.Width; H = $s.Height }
    }
    $loose.Delete()
    $truth
}

function Format-Truth {
    param($Truth, [string] $Name)
    $t = $Truth[$Name]
    '{0,-12} TRUE centre=({1:F2},{2:F2}) Rot={3,6:F1} W={4:F2} H={5:F2}' -f $Name, $t.CX, $t.CY, $t.Rot, $t.W, $t.H
}

function Get-SelectionReport {
    param($App)
    $sel = $App.ActiveWindow.Selection
    $lines = @("Selection.Type={0}" -f $sel.Type)
    try {
        $range = $sel.ShapeRange
        $names = @(); for ($i = 1; $i -le $range.Count; $i++) { $names += $range.Item($i).Name }
        $lines += "ShapeRange: " + ($names -join ', ')
    } catch { $lines += "ShapeRange: (error) $($_.Exception.Message)" }
    try {
        $lines += "HasChildShapeRange=" + $sel.HasChildShapeRange
        if ($sel.HasChildShapeRange) {
            $child = $sel.ChildShapeRange
            $names = @(); for ($i = 1; $i -le $child.Count; $i++) { $names += $child.Item($i).Name }
            $lines += "ChildShapeRange: " + ($names -join ', ')
        }
    } catch { $lines += "ChildShapeRange: (error) $($_.Exception.Message)" }
    $lines
}

function Invoke-Probe {
    param([int] $Number, [string] $Question, [scriptblock] $Body)
    try { & $Body }
    catch { Add-Finding $Number $Question ('PROBE FAILED: ' + $_.Exception.Message) @($_.ScriptStackTrace) }
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

    [void](New-Rect $slide 'K1' 100 100)
    [void](New-Rect $slide 'K2' 220 140)
    [void](New-Rect $slide 'K3' 340 120)
    $g1 = New-Group $slide 'G1' @('K1', 'K2', 'K3')
    [void](New-Rect $slide 'M1' 520 100)
    [void](New-Rect $slide 'M2' 620 160)
    $g2 = New-Group $slide 'G2' @('M1', 'M2')
    $loose = New-Rect $slide 'T' 100 420

    # =============================================================================================
    Invoke-Probe 11 'What does the selection report for shapes inside a group?' {
        $ev = @()
        $g1.GroupItems.Item('K3').Select($msoTrue)
        $g1.GroupItems.Item('K1').Select($msoFalse)
        $g1.GroupItems.Item('K2').Select($msoFalse)
        $ev += '-- selected K3, then K1, then K2 (all in G1):'
        $ev += Get-SelectionReport $ppt
        $orderKept = ($ev -join ' ') -match 'ChildShapeRange: K3, K1, K2'

        $ev += '-- then added M1 (in G2):'
        try { $g2.GroupItems.Item('M1').Select($msoFalse); $ev += Get-SelectionReport $ppt }
        catch { $ev += "Select failed: $($_.Exception.Message)" }

        $g1.GroupItems.Item('K1').Select($msoTrue)
        $ev += '-- selected K1, then added loose T:'
        try { $loose.Select($msoFalse); $ev += Get-SelectionReport $ppt }
        catch { $ev += "Select failed: $($_.Exception.Message)" }

        # Control: if this works, the view was active and the failures above are refusals.
        $ev += '-- control: selected loose T alone, then added G2 as a whole:'
        try { $loose.Select($msoTrue); $g2.Select($msoFalse); $ev += Get-SelectionReport $ppt }
        catch { $ev += "Select failed: $($_.Exception.Message)" }

        $g1.Select($msoTrue)
        $ev += '-- selected G1 itself:'
        $ev += Get-SelectionReport $ppt

        Add-Finding 11 'What does the selection report for shapes inside a group?' `
            ("Script selection order kept in ChildShapeRange: {0}. Cross-group and mixed cases below." -f $orderKept) $ev
    }

    # =============================================================================================
    Invoke-Probe 12 'Unrotated group: are child frames slide coordinates, and writable?' {
        $ev = @()
        $ev += Format-Shape $g1 'G1 before'
        foreach ($n in 'K1', 'K2', 'K3') { $ev += Format-Shape $g1.GroupItems.Item($n) }
        $k = $g1.GroupItems.Item('K2')
        $ev += ("K2.Id={0} ParentGroup={1} Type={2}" -f $k.Id, $k.ParentGroup.Name, $k.Type)
        $inSlideShapes = $false
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) { if ($slide.Shapes.Item($i).Id -eq $k.Id) { $inSlideShapes = $true } }
        $ev += "K2 found among slide.Shapes by Id: $inSlideShapes"

        $k.Left = 100
        $ev += '-- K2.Left = 100 (inside the group frame):'
        $ev += Format-Shape $g1 'G1'
        foreach ($n in 'K1', 'K2', 'K3') { $ev += Format-Shape $g1.GroupItems.Item($n) }

        $g1.GroupItems.Item('K3').Left = 40
        $ev += '-- K3.Left = 40 (outside the group frame):'
        $ev += Format-Shape $g1 'G1'
        foreach ($n in 'K1', 'K2', 'K3') { $ev += Format-Shape $g1.GroupItems.Item($n) }

        $g1.GroupItems.Item('K1').Width = 120
        $g1.GroupItems.Item('K1').Rotation = 15
        $ev += '-- K1.Width = 120, K1.Rotation = 15:'
        $ev += Format-Shape $g1 'G1'
        foreach ($n in 'K1', 'K2', 'K3') { $ev += Format-Shape $g1.GroupItems.Item($n) }
        $truth = Get-Truth $g1
        foreach ($n in 'K1', 'K2', 'K3') { $ev += Format-Truth $truth $n }

        Add-Finding 12 'Unrotated group: are child frames slide coordinates, and writable?' 'See evidence.' $ev
    }

    # =============================================================================================
    Invoke-Probe 13 'Does one Ctrl+Z reverse child writes, group frame included?' {
        $ev = @()
        $before = Format-Shape $g1 'G1'
        $k1 = $g1.GroupItems.Item('K1'); $k2 = $g1.GroupItems.Item('K2')
        $k1Left = $k1.Left; $k2Top = $k2.Top
        $ppt.StartNewUndoEntry()
        $k1.Left = [single]($k1Left + 200)
        $k2.Top = [single]($k2Top + 150)
        $ev += 'after:  ' + (Format-Shape $g1 'G1')
        $ppt.CommandBars.ExecuteMso('Undo')
        Start-Sleep -Milliseconds 500
        $k1 = $g1.GroupItems.Item('K1'); $k2 = $g1.GroupItems.Item('K2')
        $ev += 'before: ' + $before
        $ev += 'undone: ' + (Format-Shape $g1 'G1')
        $ok = ([Math]::Abs($k1.Left - $k1Left) -lt 0.5) -and ([Math]::Abs($k2.Top - $k2Top) -lt 0.5) -and ($presentation.Slides.Count -eq 1)
        Add-Finding 13 'Does one Ctrl+Z reverse child writes, group frame included?' `
            ("Both child writes reversed by one undo, slide intact: {0}" -f $ok) $ev
    }

    # =============================================================================================
    Invoke-Probe 14 'Rotated group: what does a child report, and where do writes go?' {
        $ev = @()
        [void](New-Rect $slide 'R1' 100 260 60 40 'one')
        [void](New-Rect $slide 'R2' 220 260 60 40 'two')
        [void](New-Rect $slide 'R3' 340 300 60 40 'three')
        $r = New-Group $slide 'R' @('R1', 'R2', 'R3')
        $r.Rotation = 30
        $ev += Format-Shape $r 'R (30deg)'
        foreach ($n in 'R1', 'R2', 'R3') { $ev += Format-Shape $r.GroupItems.Item($n) }
        $truth = Get-Truth $r
        foreach ($n in 'R1', 'R2', 'R3') { $ev += Format-Truth $truth $n }

        $ev += '-- R1.Left += 20 (inside the frame):'
        $r1 = $r.GroupItems.Item('R1'); $r1.Left = [single]($r1.Left + 20)
        $ev += Format-Shape $r 'R'
        foreach ($n in 'R1', 'R2', 'R3') { $ev += Format-Shape $r.GroupItems.Item($n) }
        $truth = Get-Truth $r
        foreach ($n in 'R1', 'R2', 'R3') { $ev += Format-Truth $truth $n }

        $ev += '-- R3.Top += 100 (outside the frame, so the group frame grows):'
        $r3 = $r.GroupItems.Item('R3'); $r3.Top = [single]($r3.Top + 100)
        $ev += Format-Shape $r 'R'
        foreach ($n in 'R1', 'R2', 'R3') { $ev += Format-Shape $r.GroupItems.Item($n) }
        $truth = Get-Truth $r
        foreach ($n in 'R1', 'R2', 'R3') { $ev += Format-Truth $truth $n }

        $ev += '-- R2.Rotation = 10:'
        $r.GroupItems.Item('R2').Rotation = 10
        $ev += Format-Shape $r.GroupItems.Item('R2')
        $truth = Get-Truth $r
        $ev += Format-Truth $truth 'R2'

        Add-Finding 14 'Rotated group: what does a child report, and where do writes go?' 'See evidence.' $ev
    }

    # =============================================================================================
    Invoke-Probe 15 'Rotated group: are a child''s text bounds in slide space?' {
        $ev = @()
        $r = $slide.Shapes.Item('R')
        $r2 = $r.GroupItems.Item('R2')
        $t = $r2.TextFrame2.TextRange
        $ev += ('R2 text bounds L={0:F2} T={1:F2} W={2:F2} H={3:F2} centre=({4:F2},{5:F2})' -f `
            $t.BoundLeft, $t.BoundTop, $t.BoundWidth, $t.BoundHeight, ($t.BoundLeft + $t.BoundWidth / 2), ($t.BoundTop + $t.BoundHeight / 2))
        $ev += Format-Shape $r2
        $truth = Get-Truth $r
        $ev += Format-Truth $truth 'R2'
        Add-Finding 15 'Rotated group: are a child''s text bounds in slide space?' 'Compare the text centre with the TRUE centre.' $ev
    }

    # =============================================================================================
    Invoke-Probe 16 'How do nested groups appear?' {
        $ev = @()
        [void](New-Rect $slide 'P1' 100 470 50 30)
        [void](New-Rect $slide 'P2' 200 470 50 30)
        [void](New-Rect $slide 'P3' 300 470 50 30)
        $inner = New-Group $slide 'Inner' @('P2', 'P3')
        $outer = New-Group $slide 'Outer' @('P1', 'Inner')
        $names = @(); for ($i = 1; $i -le $outer.GroupItems.Count; $i++) { $names += $outer.GroupItems.Item($i).Name }
        $ev += "Outer.GroupItems: " + ($names -join ', ')
        $p2 = $outer.GroupItems.Item('P2')
        $ev += "P2.ParentGroup = " + $p2.ParentGroup.Name
        $ev += "Inner still exists as a shape name: " + ([bool]($names -contains 'Inner'))

        $outer.GroupItems.Item('P1').Select($msoTrue)
        $p2.Select($msoFalse)
        $ev += '-- selected P1, then P2:'
        $ev += Get-SelectionReport $ppt
        Add-Finding 16 'How do nested groups appear?' 'See evidence.' $ev
    }

    # =============================================================================================
    Invoke-Probe 17 'Where does Duplicate put the copy of a child?' {
        $ev = @()
        $countBefore = $slide.Shapes.Count
        $itemsBefore = $g1.GroupItems.Count
        $copy = $g1.GroupItems.Item('K2').Duplicate().Item(1)
        $fresh = $slide.Shapes.Item('G1')
        $names = @(); for ($i = 1; $i -le $fresh.GroupItems.Count; $i++) { $names += $fresh.GroupItems.Item($i).Name }
        $ev += ("slide.Shapes {0} -> {1}; G1.GroupItems {2} -> {3} ({4})" -f $countBefore, $slide.Shapes.Count, $itemsBefore, $fresh.GroupItems.Count, ($names -join ', '))
        $ev += ("copy.Id={0} K2.Id={1}" -f $copy.Id, $g1.GroupItems.Item('K2').Id)
        $parent = try { $copy.ParentGroup.Name } catch { '(none - top level)' }
        $ev += "copy '{0}' ParentGroup: {1}" -f $copy.Name, $parent
        $ev += Format-Shape $copy 'copy'
        $ev += Format-Shape $g1.GroupItems.Item('K2') 'K2'
        $copy.Delete()
        Add-Finding 17 'Where does Duplicate put the copy of a child?' 'See evidence.' $ev
    }

    # =============================================================================================
    Invoke-Probe 18 'How does ZOrder behave on a child?' {
        $ev = @()
        function Get-Items { $o = @(); for ($i = 1; $i -le $g1.GroupItems.Count; $i++) { $s = $g1.GroupItems.Item($i); $o += ('{0}(z{1})' -f $s.Name, $s.ZOrderPosition) }; $o -join ', ' }
        $top = @(); for ($i = 1; $i -le $slide.Shapes.Count; $i++) { $top += $slide.Shapes.Item($i).Name }
        $ev += "slide.Shapes before: " + ($top -join ', ')
        $ev += "G1.GroupItems before: " + (Get-Items)
        $g1.GroupItems.Item('K1').ZOrder($msoBringToFront)
        $ev += "-- K1.ZOrder(BringToFront):"
        $ev += "G1.GroupItems after:  " + (Get-Items)
        $top = @(); for ($i = 1; $i -le $slide.Shapes.Count; $i++) { $top += $slide.Shapes.Item($i).Name }
        $ev += "slide.Shapes after:  " + ($top -join ', ')
        Add-Finding 18 'How does ZOrder behave on a child?' 'See evidence.' $ev
    }

    # =============================================================================================
    Invoke-Probe 19 'Inside a nested group, how far does ZOrder move a leaf?' {
        $ev = @()
        $outer = $slide.Shapes.Item('Outer')
        function Get-Stack { $o = @(); for ($i = 1; $i -le $outer.GroupItems.Count; $i++) { $s = $outer.GroupItems.Item($i); $o += ('{0}(z{1})' -f $s.Name, $s.ZOrderPosition) }; $o -join ', ' }
        $ev += "before:                 " + (Get-Stack)
        $outer.GroupItems.Item('P1').ZOrder($msoBringToFront)
        $ev += "P1 (outer level) front: " + (Get-Stack)
        $outer.GroupItems.Item('P2').ZOrder($msoBringToFront)
        $ev += "P2 (in Inner) front:    " + (Get-Stack)
        $p2 = $outer.GroupItems.Item('P2')
        $ev += ("P2.ParentGroup={0}; P2.Child={1}" -f $p2.ParentGroup.Name, $p2.Child)
        $p1 = $outer.GroupItems.Item('P1')
        $z1 = $p1.ZOrderPosition; $z2 = $p2.ZOrderPosition
        $answer = if ($z2 -gt $z1) { 'P2 rose above P1: ZOrder crossed out of the inner group.' } else { 'P2 stayed below P1: ZOrder stayed within the inner group.' }
        Add-Finding 19 'Inside a nested group, how far does ZOrder move a leaf?' $answer $ev
    }

    # =============================================================================================
    Invoke-Probe 20 'Is the copy of a child really in the group?' {
        $ev = @()
        [void](New-Rect $slide 'D1' 100 300 60 40)
        [void](New-Rect $slide 'D2' 200 300 60 40)
        [void](New-Rect $slide 'D3' 300 300 60 40)
        $d = New-Group $slide 'D' @('D1', 'D2', 'D3')
        $groupId = $d.Id

        # Every view of the slide, read fresh: top-level shapes, and each group's items by id.
        function Get-Census {
            $lines = @()
            $top = @(); for ($i = 1; $i -le $slide.Shapes.Count; $i++) { $top += $slide.Shapes.Item($i).Name }
            $lines += ("slide.Shapes ({0}): {1}" -f $slide.Shapes.Count, ($top -join ', '))
            $g = $slide.Shapes.Item('D')
            $items = @(); for ($i = 1; $i -le $g.GroupItems.Count; $i++) { $c = $g.GroupItems.Item($i); $items += ('{0}#{1}(z{2})' -f $c.Name, $c.Id, $c.ZOrderPosition) }
            $lines += ("D.GroupItems ({0}): {1}" -f $g.GroupItems.Count, ($items -join ', '))
            $lines += ("D frame: L={0:F1} T={1:F1} W={2:F1} H={3:F1}" -f $g.Left, $g.Top, $g.Width, $g.Height)
            $lines
        }
        # What the group really holds: copy it, ungroup the copy, count and name the loose shapes.
        function Get-TrueContents {
            $g = $slide.Shapes.Item('D')
            $dup = $g.Duplicate().Item(1)
            $loose = $dup.Ungroup()
            $out = @(); for ($i = 1; $i -le $loose.Count; $i++) { $s = $loose.Item($i); $out += ('{0}@({1:F0},{2:F0})' -f $s.Name, $s.Left, $s.Top) }
            $loose.Delete()
            "true contents ({0}): {1}" -f $out.Count, ($out -join ', ')
        }
        function Find-Anywhere { param([int] $Id)
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.Id -eq $Id) { return "top-level #$i" }
                if ($s.Type -eq 6) { for ($j = 1; $j -le $s.GroupItems.Count; $j++) { if ($s.GroupItems.Item($j).Id -eq $Id) { return "in group '$($s.Name)' item $j" } } }
            }
            'nowhere'
        }

        $ev += '-- before:'
        $ev += Get-Census

        $ppt.StartNewUndoEntry()
        $copy = $slide.Shapes.Item('D').GroupItems.Item('D2').Duplicate().Item(1)
        $copyId = $copy.Id
        $ev += '-- after D2.Duplicate():'
        $ev += ("copy: name={0} id={1} Child={2} ParentGroup={3}#{4} z={5} at ({6:F0},{7:F0})" -f `
            $copy.Name, $copyId, $copy.Child, $copy.ParentGroup.Name, $copy.ParentGroup.Id, $copy.ZOrderPosition, $copy.Left, $copy.Top)
        $ev += ("group D id still {0}: {1}" -f $groupId, ($slide.Shapes.Item('D').Id -eq $groupId))
        $ev += Get-Census
        $ev += ("copy found by walking the slide: " + (Find-Anywhere $copyId))
        $ev += Get-TrueContents

        $copy.Left = 500
        $ev += '-- copy.Left = 500:'
        $ev += ("copy now at ({0:F0},{1:F0})" -f $copy.Left, $copy.Top)
        $ev += Get-Census
        $ev += Get-TrueContents

        $ppt.CommandBars.ExecuteMso('Undo')
        Start-Sleep -Milliseconds 500
        $ev += '-- one Ctrl+Z:'
        $ev += Get-Census
        $ev += ("copy found by walking the slide: " + (Find-Anywhere $copyId))
        $ev += Get-TrueContents

        # The child range a selection inside the group gives, duplicated as one.
        $g = $slide.Shapes.Item('D')
        $g.GroupItems.Item('D1').Select($msoTrue)
        $g.GroupItems.Item('D3').Select($msoFalse)
        $ev += '-- Selection.ChildShapeRange (D1, D3).Duplicate():'
        try {
            $copies = $ppt.ActiveWindow.Selection.ChildShapeRange.Duplicate()
            $names = @(); for ($i = 1; $i -le $copies.Count; $i++) { $c = $copies.Item($i); $names += ('{0}#{1} parent={2}' -f $c.Name, $c.Id, $(try { $c.ParentGroup.Name } catch { '(none)' })) }
            $ev += ("copies: " + ($names -join '; '))
            $ev += Get-Census
            $ev += Get-TrueContents
        } catch { $ev += "failed: $($_.Exception.Message)" }

        # Ungroup and regroup: does the group come back with its name, its id, and room for more?
        $ev += '-- Ungroup, add a loose copy of D1, Regroup:'
        $g = $slide.Shapes.Item('D')
        $members = $g.Ungroup()
        $loose = $slide.Shapes.Item('D1').Duplicate().Item(1)
        $loose.Name = 'D1-copy'
        try {
            $again = $members.Regroup()
            $ev += ("Regroup: name={0} id={1} (was #{2}) items={3}" -f $again.Name, $again.Id, $groupId, $again.GroupItems.Count)
        } catch { $ev += "Regroup failed: $($_.Exception.Message)" }
        $all = @(); for ($i = 1; $i -le $slide.Shapes.Count; $i++) { $all += $slide.Shapes.Item($i).Name }
        $ev += "slide.Shapes: " + ($all -join ', ')

        Add-Finding 20 'Is the copy of a child really in the group?' 'See evidence.' $ev
    }

    # =============================================================================================
    Invoke-Probe 21 'Does GroupItems ever catch up with a copy?' {
        $ev = @()
        [void](New-Rect $slide 'E1' 100 420 60 40)
        [void](New-Rect $slide 'E2' 200 420 60 40)
        $e = New-Group $slide 'E' @('E1', 'E2')
        $copy = $slide.Shapes.Item('E').GroupItems.Item('E2').Duplicate().Item(1)
        $copyId = $copy.Id
        $copy.Name = 'E2-copy'

        function Get-Count { param([string] $When)
            $g = $slide.Shapes.Item('E')
            $ids = @(); for ($i = 1; $i -le $g.GroupItems.Count; $i++) { $ids += $g.GroupItems.Item($i).Id }
            '{0,-36} GroupItems={1} copy listed={2}' -f $When, $g.GroupItems.Count, ($ids -contains $copyId)
        }
        $ev += Get-Count 'straight after Duplicate'

        $copy.Select($msoTrue)
        $ev += Get-Count 'after copy.Select()'
        $sel = $ppt.ActiveWindow.Selection
        if ($sel.HasChildShapeRange) {
            $c = $sel.ChildShapeRange.Item(1)
            $ev += ('selection child range: {0}#{1}; writable: ' -f $c.Name, $c.Id) + $(try { $c.Left = [single]($c.Left + 1); 'yes' } catch { 'no' })
        } else { $ev += 'selection has no child range' }

        $ppt.StartNewUndoEntry()
        $ev += Get-Count 'after StartNewUndoEntry'

        $slide.Shapes.Item('E').Select($msoTrue)
        $ev += Get-Count 'after selecting the group'

        $g = $slide.Shapes.Item('E'); $g.Left = [single]($g.Left + 1); $g.Left = [single]($g.Left - 1)
        $ev += Get-Count 'after nudging the group'

        $g = $slide.Shapes.Item('E')
        $dup = $g.Duplicate().Item(1)
        $loose = $dup.Ungroup()
        $names = @(); for ($i = 1; $i -le $loose.Count; $i++) { $names += $loose.Item($i).Name }
        $loose.Delete()
        $ev += 'live E, ungrouped copy holds:   ' + ($names -join ', ')

        $path = Join-Path $env:TEMP ("probe-groups-{0}.pptx" -f [guid]::NewGuid().ToString('N'))
        # A copy, so the scratch deck stays untitled and the file reopened below is a different one.
        $presentation.SaveCopyAs($path)
        $ev += Get-Count 'after SaveCopyAs'

        $other = $ppt.Presentations.Open($path, $msoFalse, $msoFalse, $msoFalse)   # no window; edited only in memory
        $og = $other.Slides.Item(1).Shapes.Item('E')
        $ids = @(); for ($i = 1; $i -le $og.GroupItems.Count; $i++) { $ids += $og.GroupItems.Item($i).Name }
        $ev += ('{0,-36} GroupItems={1} names={2}' -f 'reopened from disk', $og.GroupItems.Count, ($ids -join ', '))
        $oslide = $other.Slides.Item(1)
        $census = @()
        for ($i = 1; $i -le $oslide.Shapes.Count; $i++) {
            $sh = $oslide.Shapes.Item($i)
            $entry = $sh.Name
            if ($sh.Type -eq 6) { $kids = @(); for ($j = 1; $j -le $sh.GroupItems.Count; $j++) { $kids += $sh.GroupItems.Item($j).Name }; $entry += '[' + ($kids -join ',') + ']' }
            $census += $entry
        }
        $ev += 'reopened slide: ' + ($census -join ', ')
        $dup = $og.Duplicate().Item(1)
        $loose = $dup.Ungroup()
        $names = @(); for ($i = 1; $i -le $loose.Count; $i++) { $names += $loose.Item($i).Name }
        $ev += 'reopened E, ungrouped copy holds: ' + ($names -join ', ')
        $other.Close()
        Remove-Item $path -ErrorAction SilentlyContinue

        Add-Finding 21 'Does GroupItems ever catch up with a copy?' 'See evidence.' $ev
    }

    # =============================================================================================
    Invoke-Probe 22 'Does PowerPoint''s own Duplicate make copies GroupItems cannot see?' {
        $ev = @()
        [void](New-Rect $slide 'F1' 600 420 60 40)
        [void](New-Rect $slide 'F2' 700 420 60 40)
        $fg = New-Group $slide 'F' @('F1', 'F2')
        $fg.GroupItems.Item('F2').Select($msoTrue)
        # No ExecuteMso name for it was found ('DuplicateSelection', 'Duplicate', 'ObjectDuplicate' all
        # fail), so press Ctrl+D as a person would, with the window in front and the slide focused.
        $ran = $null
        $proc = Get-Process -Name 'POWERPNT' | Select-Object -First 1
        $shell = New-Object -ComObject WScript.Shell
        if ($shell.AppActivate($proc.Id)) {
            Start-Sleep -Milliseconds 600
            $fg.GroupItems.Item('F2').Select($msoTrue)
            Start-Sleep -Milliseconds 300
            $shell.SendKeys('^d')
            $ran = 'Ctrl+D'
        } else { $ev += 'could not bring PowerPoint to the front' }
        Start-Sleep -Milliseconds 800
        if ($ran) {
            $g = $slide.Shapes.Item('F')
            $ev += ("pressed {0}; GroupItems={1}; slide.Shapes={2}" -f $ran, $g.GroupItems.Count, $slide.Shapes.Count)
            $sel = $ppt.ActiveWindow.Selection
            if ($sel.HasChildShapeRange) { $c = $sel.ChildShapeRange.Item(1); $ev += ('now selected: {0}#{1} Child={2}' -f $c.Name, $c.Id, $c.Child) }
            $dup = $g.Duplicate().Item(1)
            $loose = $dup.Ungroup()
            $names = @(); for ($i = 1; $i -le $loose.Count; $i++) { $names += $loose.Item($i).Name }
            $loose.Delete()
            $ev += 'F, ungrouped copy holds: ' + ($names -join ', ')
            $answer = if ($names.Count -le 2) { 'Ctrl+D did not reach the slide - no copy was made.' } elseif ($names.Count -gt $g.GroupItems.Count) { "Yes: PowerPoint's own Ctrl+D makes a member GroupItems does not list." } else { "No: the copy PowerPoint's own Ctrl+D makes is listed." }
        } else { $answer = 'Could not press Ctrl+D in PowerPoint; check by hand.' }
        Add-Finding 22 'Does PowerPoint''s own Duplicate make copies GroupItems cannot see?' $answer $ev
    }

    # =============================================================================================
    Invoke-Probe 23 'Does deselecting and reselecting the group refresh GroupItems?' {
        $ev = @()
        [void](New-Rect $slide 'H1' 600 100 60 40)
        [void](New-Rect $slide 'H2' 700 100 60 40)
        $h = New-Group $slide 'H' @('H1', 'H2')
        $copy = $h.GroupItems.Item('H2').Duplicate().Item(1)
        $copyId = $copy.Id

        function Get-Listed { param([string] $When)
            $g = $slide.Shapes.Item('H')
            $ids = @(); for ($i = 1; $i -le $g.GroupItems.Count; $i++) { $ids += $g.GroupItems.Item($i).Id }
            '{0,-52} GroupItems={1} copy listed={2}' -f $When, $g.GroupItems.Count, ($ids -contains $copyId)
        }
        $ev += Get-Listed 'straight after Duplicate'

        $ppt.ActiveWindow.Selection.Unselect()
        $ev += Get-Listed 'script: Selection.Unselect()'
        $slide.Shapes.Item('H').Select($msoTrue)
        $ev += Get-Listed 'script: then select the group'

        $slide.Shapes.Item('T').Select($msoTrue)
        $slide.Shapes.Item('H').Select($msoTrue)
        $ev += Get-Listed 'script: select another shape, then the group'

        # From the keyboard: Esc twice clears any selection, Tab walks the slide's shapes.
        $proc = Get-Process -Name 'POWERPNT' | Select-Object -First 1
        $shell = New-Object -ComObject WScript.Shell
        if ($shell.AppActivate($proc.Id)) {
            Start-Sleep -Milliseconds 600
            $shell.SendKeys('{ESC}'); Start-Sleep -Milliseconds 300
            $shell.SendKeys('{ESC}'); Start-Sleep -Milliseconds 300
            $ev += ('keyboard: after Esc, Esc - selection type {0}' -f $ppt.ActiveWindow.Selection.Type)
            $ev += Get-Listed 'keyboard: Esc, Esc'
            $reached = $false
            for ($i = 1; $i -le $slide.Shapes.Count + 2 -and -not $reached; $i++) {
                $shell.SendKeys('{TAB}'); Start-Sleep -Milliseconds 300
                $sel = $ppt.ActiveWindow.Selection
                if ($sel.Type -eq 2 -and $sel.ShapeRange.Count -eq 1 -and $sel.ShapeRange.Item(1).Name -eq 'H') { $reached = $true }
            }
            $ev += Get-Listed ('keyboard: Tab to the group (reached: {0})' -f $reached)
        } else { $ev += 'could not bring PowerPoint to the front' }

        $listed = ($ev -join ' ') -match 'copy listed=True'
        $answer = if ($listed) { 'Yes: at least one of these made GroupItems list the copy - see which.' } else { 'No: none of these made GroupItems list the copy.' }
        Add-Finding 23 'Does deselecting and reselecting the group refresh GroupItems?' $answer $ev
    }
}
finally {
    if ($OutFile -and $findings.Count) {
        $md = [System.Text.StringBuilder]::new()
        [void]$md.AppendLine(('Measured by `tools/Probe-Groups.ps1` on {0}, PowerPoint {1} build {2}.' -f `
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
