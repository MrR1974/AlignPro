<#
.SYNOPSIS
    Removes AlignPro from PowerPoint for the current user.

.DESCRIPTION
    Undoes exactly what Install-AlignPro.ps1 did: deletes the registry key PowerPoint loads the add-in
    from, and removes the installed files. Nothing else on the machine was touched, so nothing else
    needs cleaning up.

    Your presentations are never touched.

.PARAMETER InstallPath
    Where the files were installed. Defaults to the location the installer uses, but the registry is
    consulted first so a custom location is found automatically.

.PARAMETER KeepFiles
    Unregister the add-in but leave the files on disk.

.PARAMETER KeepLog
    Leave the diagnostic log behind. Useful when removing AlignPro to report a problem with it.

.EXAMPLE
    .\Uninstall-AlignPro.ps1
#>
[CmdletBinding()]
param(
    [string] $InstallPath,
    [switch] $KeepFiles,
    [switch] $KeepLog
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$registryKey = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\AlignPro.AddIn'

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
    exit 1
}

if (Test-Path $registryKey) {
    Remove-Item -Path $registryKey -Recurse -Force
    Write-Host '  unregistered from PowerPoint'
}
else {
    Write-Host '  was not registered' -ForegroundColor DarkGray
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
    )

    # The add-in writes its diagnostic log into this same folder, so without counting it as ours the
    # folder would always survive an uninstall - looking like the uninstaller had failed.
    if (-not $KeepLog) { $ours += 'alignpro.log' }
    foreach ($file in $ours) {
        $full = Join-Path $InstallPath $file
        if (Test-Path $full) { Remove-Item -LiteralPath $full -Force }
    }

    $remaining = @(Get-ChildItem -LiteralPath $InstallPath -Force -ErrorAction SilentlyContinue)
    if ($remaining.Count -eq 0) {
        [System.IO.Directory]::Delete($InstallPath, $false)
        Write-Host "  removed $InstallPath"
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
