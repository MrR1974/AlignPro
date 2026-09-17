<#
.SYNOPSIS
    Hunts for a command that closes PowerPoint's undo coalescing group while leaving the document
    visually unchanged.

.DESCRIPTION
    Established by earlier runs of this probe and by the add-in's own log:

      - Object-model changes coalesce into one undo entry, and ExecuteMso('Undo') is a faithful
        stand-in for Ctrl+Z, so all of this is measurable from script.
      - A selection change does NOT close the group. Only a command that modifies the document does:
        Bold closes it, while Copy, SelectAll, ViewGridlines and ViewRuler - all non-modifying - do not.
      - PowerPoint ignores customUI <commands> repurposing entirely, so Undo cannot be intercepted
        that way.

    That leaves one attractive option: if AlignPro closes the group itself immediately before applying
    its changes, PowerPoint's own undo entry then contains exactly one AlignPro operation and native
    Ctrl+Z becomes correct - no keyboard hook needed. The closer has to be invisible, which is what
    this script looks for.

    Each case: build a fixture, move A, run the candidate, move C, undo once. A surviving at 320 with
    C back at 200 means the group closed. Formatting is read afterwards to catch side effects.

.EXAMPLE
    .\Probe-UndoGrouping.ps1
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$msoTrue           = -1
$msoFalse          = 0
$msoShapeRectangle = 1
$ppLayoutBlank     = 12

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $msoTrue

function New-Fixture {
    $pres = $ppt.Presentations.Add()
    $slide = $pres.Slides.Add(1, $ppLayoutBlank)
    foreach ($spec in @(@{ N = 'A'; X = 100 }, @{ N = 'B'; X = 300 }, @{ N = 'C'; X = 500 })) {
        $s = $slide.Shapes.AddShape($msoShapeRectangle, $spec.X, 200, 120, 80)
        $s.Name = $spec.N
        $s.TextFrame.TextRange.Text = $spec.N
    }
    # 'Blank' deliberately carries no text, to test whether a text command on an empty shape still
    # registers as a modification.
    $blank = $slide.Shapes.AddShape($msoShapeRectangle, 100, 380, 120, 80)
    $blank.Name = 'Blank'
    [pscustomobject]@{ Presentation = $pres; Slide = $slide }
}

function Invoke-Mso {
    param([string] $Id)
    $ppt.CommandBars.ExecuteMso($Id)
}

function Test-Closer {
    param([string] $Label, [scriptblock] $Action)

    $f = New-Fixture
    $row = [ordered]@{ Candidate = $Label; Ran = 'ok'; ClosedGroup = ''; VisibleChange = '' }

    try {
        $f.Slide.Shapes.Item('A').Top = 320

        try { & $Action $f }
        catch {
            $row.Ran = 'FAILED'
            $row.ClosedGroup = '-'
            $row.VisibleChange = ($_.Exception.Message -replace "`r?`n", ' ')
            return [pscustomobject]$row
        }

        # Any formatting left behind on the shapes the candidate touched?
        $changes = @()
        foreach ($name in 'B', 'Blank') {
            try {
                $font = $f.Slide.Shapes.Item($name).TextFrame2.TextRange.Font
                if ($font.Bold -eq $msoTrue) { $changes += "$name bold" }
                if ($font.Italic -eq $msoTrue) { $changes += "$name italic" }
            }
            catch { }
        }
        $row.VisibleChange = if ($changes.Count) { $changes -join ', ' } else { 'none' }

        $f.Slide.Shapes.Item('C').Top = 400
        Invoke-Mso -Id 'Undo'

        if ($f.Presentation.Slides.Count -eq 0) {
            $row.ClosedGroup = 'NO - deck emptied'
        }
        else {
            $a = [int] $f.Slide.Shapes.Item('A').Top
            $c = [int] $f.Slide.Shapes.Item('C').Top
            $row.ClosedGroup = if ($a -ge 319 -and $c -le 201) { 'YES' } else { "no (A=$a C=$c)" }
        }
    }
    finally {
        try { $f.Presentation.Saved = $msoTrue; $f.Presentation.Close() } catch { }
    }

    [pscustomobject]$row
}

$cases = @(
    @{
        Label  = 'Bold once (known closer)'
        Action = { param($f) $f.Slide.Shapes.Item('B').Select($msoTrue); Invoke-Mso 'Bold' }
    }
    @{
        Label  = 'Bold twice (self-cancelling)'
        Action = { param($f) $f.Slide.Shapes.Item('B').Select($msoTrue); Invoke-Mso 'Bold'; Invoke-Mso 'Bold' }
    }
    @{
        Label  = 'Bold on a shape with no text'
        Action = { param($f) $f.Slide.Shapes.Item('Blank').Select($msoTrue); Invoke-Mso 'Bold' }
    }
    @{
        Label  = 'Italic twice on no-text shape'
        Action = { param($f) $f.Slide.Shapes.Item('Blank').Select($msoTrue); Invoke-Mso 'Italic'; Invoke-Mso 'Italic' }
    }
    @{
        # The promising one. A real undo restores the previous formatting exactly, so unlike a double
        # toggle it is lossless even when the selection's formatting is mixed. The question is whether
        # the group boundary survives undoing the command that created it.
        Label  = 'Italic then ExecuteMso(Undo)'
        Action = { param($f) $f.Slide.Shapes.Item('B').Select($msoTrue); Invoke-Mso 'Italic'; Invoke-Mso 'Undo' }
    }
    @{
        Label  = 'Bold on no-text shape then Undo'
        Action = { param($f) $f.Slide.Shapes.Item('Blank').Select($msoTrue); Invoke-Mso 'Bold'; Invoke-Mso 'Undo' }
    }
)

Write-Host ''
Write-Host 'Looking for an invisible group closer...' -ForegroundColor Cyan
Write-Host ''
$results = foreach ($case in $cases) { Test-Closer -Label $case.Label -Action $case.Action }
$results | Format-Table -AutoSize

Write-Host 'Wanted: ClosedGroup=YES with VisibleChange=none. That lets AlignPro close the group before' -ForegroundColor DarkGray
Write-Host 'applying, so PowerPoint own undo covers exactly one AlignPro operation and no keyboard hook' -ForegroundColor DarkGray
Write-Host 'is needed. If nothing qualifies, a hook is the only remaining protection for Ctrl+Z.' -ForegroundColor DarkGray
Write-Host ''
