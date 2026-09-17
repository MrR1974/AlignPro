<#
.SYNOPSIS
    Installs AlignPro into PowerPoint for the current user.

.DESCRIPTION
    Copies the add-in files to a stable location and registers them with PowerPoint. That is the whole
    installation: some files and one registry value under HKEY_CURRENT_USER.

    No administrator rights, no Visual Studio and no compiler are needed.

    Trust IS needed, however. VSTO will not load an add-in unless the machine trusts the certificate
    that signed its manifest; without it PowerPoint sets LoadBehavior to 2 and the add-in silently
    never appears. The "|vstolocal" suffix used below controls where the add-in is loaded FROM - it
    does not exempt it from that check. This script does not grant trust, and cannot: that comes from
    the signing certificate being one the machine already trusts.

    Everything it needs is already present on a machine that runs Office: the .NET Framework ships with
    Windows, and the VSTO runtime ships with Office. The script checks both and says so plainly if
    either is missing.

    Safe to re-run: installing over an existing copy upgrades it.

.PARAMETER InstallPath
    Where to put the files. Defaults to AlignPro under your local application data.

.PARAMETER Source
    Folder holding the add-in files. Defaults to the folder this script is in, which is where they sit
    when you extract the release zip.

.PARAMETER Force
    Install even if PowerPoint is running. The copy will fail if PowerPoint has the files open, so this
    is only useful when upgrading a copy PowerPoint has not loaded.

.EXAMPLE
    .\Install-AlignPro.ps1

.EXAMPLE
    .\Install-AlignPro.ps1 -InstallPath 'D:\Tools\AlignPro'
#>
[CmdletBinding()]
param(
    [string] $InstallPath = (Join-Path $env:LOCALAPPDATA 'AlignPro'),
    [string] $Source = $PSScriptRoot,
    [switch] $Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$addInName = 'AlignPro.AddIn'
$registryKey = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$addInName"

# The complete set. Anything missing means an incomplete download rather than a broken machine.
$required = @(
    'AlignPro.AddIn.dll'
    'AlignPro.AddIn.dll.manifest'
    'AlignPro.AddIn.vsto'
    'AlignPro.Geometry.dll'
    'Microsoft.Office.Tools.Common.v4.0.Utilities.dll'
)

function Fail {
    param([string] $Message, [string] $Remedy)
    Write-Host ''
    Write-Host $Message -ForegroundColor Red
    if ($Remedy) { Write-Host $Remedy -ForegroundColor Yellow }
    Write-Host ''
    exit 1
}

Write-Host ''
Write-Host 'Installing AlignPro' -ForegroundColor Cyan
Write-Host ''

# --- what we are installing ----------------------------------------------------------------------
$missing = $required | Where-Object { -not (Test-Path (Join-Path $Source $_)) }
if ($missing) {
    Fail "These files are missing from '$Source':`n  $($missing -join "`n  ")" `
         'Extract the whole release zip and run this script from inside it.'
}

# --- prerequisites -------------------------------------------------------------------------------
# .NET Framework 4.8 is release 528040 or higher, and ships with Windows 10 1903 onwards.
$ndp = 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'
$release = if (Test-Path $ndp) { (Get-ItemProperty $ndp).Release } else { 0 }
if ($release -lt 528040) {
    Fail 'The .NET Framework 4.8 is required and was not found.' `
         'Install it from https://dotnet.microsoft.com/download/dotnet-framework/net48 and run this again.'
}
Write-Host "  .NET Framework 4.8            present"

# The VSTO runtime registers itself in the 32-bit view of the registry even on 64-bit Windows, so both
# have to be checked - looking only at the native view reports it missing on most machines.
$vstoKeys = @(
    'HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R'
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup\v4R'
)
$vsto = $vstoKeys | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $vsto) {
    Fail 'The Visual Studio Tools for Office runtime is required and was not found.' `
         "It normally ships with Office. Install it from https://www.microsoft.com/download/details.aspx?id=48217 and run this again."
}
Write-Host "  VSTO runtime                  $((Get-ItemProperty $vsto).Version)"

if (-not (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE')) {
    Fail 'PowerPoint does not appear to be installed.' 'AlignPro is a PowerPoint add-in and has nothing to attach to.'
}
Write-Host "  PowerPoint                    present"

# --- PowerPoint must not be holding the files ----------------------------------------------------
if ((Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue) -and -not $Force) {
    Fail 'PowerPoint is running.' 'Close PowerPoint and run this again, so the files are not locked.'
}

# --- copy ----------------------------------------------------------------------------------------
if (-not (Test-Path $InstallPath)) { New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null }
foreach ($file in $required) {
    Copy-Item (Join-Path $Source $file) (Join-Path $InstallPath $file) -Force
}
Write-Host ''
Write-Host "  installed to                  $InstallPath"

# Windows marks anything that arrived from the internet, and that mark survives both unzipping and
# copying. The .NET loader refuses to load a marked assembly, PowerPoint gives up and sets
# LoadBehavior to 2, and the add-in simply never appears - with nothing to show why. Clearing it here
# is the difference between the download working and silently doing nothing.
$blocked = 0
foreach ($file in $required) {
    $full = Join-Path $InstallPath $file
    if (Get-Item $full -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue) {
        Unblock-File -LiteralPath $full
        $blocked++
    }
}
if ($blocked -gt 0) {
    Write-Host "  unblocked                     $blocked files marked as downloaded from the internet"
}

# --- register ------------------------------------------------------------------------------------
# HKCU, so no administrator rights. LoadBehavior 3 means "load at startup".
if (-not (Test-Path $registryKey)) { New-Item -Path $registryKey -Force | Out-Null }
$manifest = 'file:///' + ((Join-Path $InstallPath 'AlignPro.AddIn.vsto') -replace '\\', '/') + '|vstolocal'
Set-ItemProperty -Path $registryKey -Name 'FriendlyName' -Value 'AlignPro'
Set-ItemProperty -Path $registryKey -Name 'Description'  -Value 'Align, distribute, size and tidy shapes in PowerPoint'
# Always force 3. If a previous attempt failed - blocked files being the usual reason - PowerPoint
# will have set this to 2, meaning "do not try again", and would ignore the add-in forever otherwise.
Set-ItemProperty -Path $registryKey -Name 'LoadBehavior' -Value 3 -Type DWord
Set-ItemProperty -Path $registryKey -Name 'Manifest'     -Value $manifest

Write-Host "  registered for                $env:USERNAME (no admin rights used)"
Write-Host ''
Write-Host 'Done. Start PowerPoint and look for the AlignPro tab.' -ForegroundColor Green
Write-Host ''
Write-Host 'If the tab does not appear, check File > Options > Add-ins > Disabled Items.' -ForegroundColor DarkGray
Write-Host 'To remove it again, run Uninstall-AlignPro.ps1.' -ForegroundColor DarkGray
Write-Host ''
