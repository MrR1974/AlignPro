<#
.SYNOPSIS
    Removes AlignPro from PowerPoint for the current user.

.DESCRIPTION
    Undoes exactly what Install-AlignPro.ps1 did: the registry key PowerPoint loads the add-in from,
    the trust entry, the Add/Remove Programs listing, and the installed files. Nothing else on the
    machine was touched, so nothing else needs cleaning up.

    Your presentations are never touched.

    A copy of this script is installed alongside the add-in, and Settings > Apps > Installed apps >
    AlignPro > Uninstall runs that copy.

.PARAMETER InstallPath
    Where the files were installed. Defaults to the location the installer uses, but the registry is
    consulted first so a custom location is found automatically.

.PARAMETER KeepFiles
    Unregister the add-in but leave the files on disk.

.PARAMETER KeepLog
    Leave the diagnostic log behind. Useful when removing AlignPro to report a problem with it.

.PARAMETER FromArp
    Set when Windows launched this from Add/Remove Programs. Holds the window open at the end so the
    result is readable, rather than letting the console close the instant the script finishes.

.EXAMPLE
    .\Uninstall-AlignPro.ps1
#>
[CmdletBinding()]
param(
    [string] $InstallPath,
    [switch] $KeepFiles,
    [switch] $KeepLog,
    [switch] $FromArp
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$registryKey = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\AlignPro.AddIn'
$uninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\AlignPro'

function Finish {
    param([int] $Code = 0)
    if ($FromArp) {
        Write-Host 'Press any key to close this window.' -ForegroundColor DarkGray
        [void]$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
    exit $Code
}

Write-Host ''
Write-Host 'Removing AlignPro' -ForegroundColor Cyan
Write-Host ''

# Prefer the path PowerPoint is actually loading from over the default, so a custom install location
# gets cleaned up rather than left behind.
if (-not $InstallPath -and (Test-Path $registryKey)) {
    $manifest = (Get-ItemProperty $registryKey).Manifest
    if ($manifest) {
        $path = $manifest -replace '^file:///', '' -replace '\|vstolocal$', '' -replace '/', '\'
        $InstallPath = Split-Path $path -Parent
    }
}
if (-not $InstallPath) { $InstallPath = Join-Path $env:LOCALAPPDATA 'AlignPro' }

if (Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue) {
    Write-Host 'PowerPoint is running. Close it and run this again, so the files are not locked.' -ForegroundColor Red
    Write-Host ''
    Finish 1
}

if (Test-Path $registryKey) {
    Remove-Item -Path $registryKey -Recurse -Force
    Write-Host '  unregistered from PowerPoint'
}
else {
    Write-Host '  was not registered' -ForegroundColor DarkGray
}

# Revoke the trust the installer granted. Leaving it behind would outlive the add-in it was for -
# a standing grant for something no longer on the machine.
$inclusionRoot = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
$revoked = 0
Get-ChildItem $inclusionRoot -ErrorAction SilentlyContinue | ForEach-Object {
    $url = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).Url
    if ($url -and $url -like '*AlignPro.AddIn.vsto') {
        Remove-Item -Path $_.PSPath -Recurse -Force
        $revoked++
    }
}
Write-Host $(if ($revoked -gt 0) { "  revoked trust ($revoked inclusion-list entries)" } else { '  no trust entries to revoke' })

if (Test-Path $uninstallKey) {
    Remove-Item -Path $uninstallKey -Recurse -Force
    Write-Host '  delisted from Add/Remove Programs'
}

if ($KeepFiles) {
    Write-Host "  files left in place at $InstallPath" -ForegroundColor DarkGray
}
elseif (Test-Path $InstallPath) {
    # Delete only the files the installer put there. A user who pointed the installer at a folder of
    # their own should not lose anything else in it.
    $ours = @(
        'AlignPro.AddIn.dll'
        'AlignPro.AddIn.dll.manifest'
        'AlignPro.AddIn.vsto'
        'AlignPro.Geometry.dll'
        'Microsoft.Office.Tools.Common.v4.0.Utilities.dll'
        'AlignPro-Sample.pptx'
    )

    # The add-in writes its diagnostic log into this same folder, so without counting it as ours the
    # folder would always survive an uninstall - looking like the uninstaller had failed.
    if (-not $KeepLog) { $ours += 'alignpro.log' }

    # This script lives in that folder too and is deleted last, because it is the one currently
    # running. PowerShell does not hold it open, so this normally succeeds; if the filesystem
    # disagrees, the folder simply survives with one file in it and the message below says so.
    $ours += 'Uninstall-AlignPro.ps1'

    $stubborn = @()
    foreach ($file in $ours) {
        $full = Join-Path $InstallPath $file
        if (Test-Path $full) {
            try { Remove-Item -LiteralPath $full -Force }
            catch { $stubborn += $file }
        }
    }

    $remaining = @(Get-ChildItem -LiteralPath $InstallPath -Force -ErrorAction SilentlyContinue)
    if ($remaining.Count -eq 0) {
        [System.IO.Directory]::Delete($InstallPath, $false)
        Write-Host "  removed $InstallPath"
    }
    elseif ($stubborn.Count -gt 0 -and $remaining.Count -eq $stubborn.Count) {
        Write-Host "  removed AlignPro's files; $InstallPath holds only this script, which Windows will tidy"
    }
    else {
        Write-Host "  removed AlignPro's files; left $InstallPath because it holds other things"
    }
}
else {
    Write-Host "  nothing to remove at $InstallPath" -ForegroundColor DarkGray
}

Write-Host ''
Write-Host 'Done. AlignPro will not load next time PowerPoint starts.' -ForegroundColor Green
Write-Host ''
Finish 0
