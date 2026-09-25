<#
.SYNOPSIS
    Installs AlignPro into PowerPoint for the current user.

.DESCRIPTION
    Copies the add-in files to a stable location and registers them with PowerPoint. That is the whole
    installation: some files and a handful of registry values under HKEY_CURRENT_USER.

    No administrator rights, no Visual Studio and no compiler are needed.

    Trust IS needed, and this script grants it - deliberately, and as narrowly as possible.

    VSTO will not load an add-in unless it trusts the certificate that signed its manifest; without
    that, PowerPoint sets LoadBehavior to 2 and the add-in silently never appears. AlignPro is signed
    by a self-signed certificate that no other machine has heard of.

    Rather than ask you to add that certificate to your trust stores - which would trust anything it
    ever signs - this adds one entry to VSTO's inclusion list: trust for THIS add-in, at THIS path,
    signed by THIS key, and nothing else. No certificate store is touched, and the uninstaller removes
    the entry again.

    It is still a trust decision, just a precise one. If you would rather not make it, build from
    source instead: Visual Studio grants trust to what it builds.

    Everything it needs is already present on a machine that runs Office: the .NET Framework ships with
    Windows, and the VSTO runtime ships with Office. The script checks both and says so plainly if
    either is missing.

    Normally reached through install.ps1, which downloads the release and calls this. Running it by
    hand from an extracted release zip does exactly the same thing.

    Safe to re-run: installing over an existing copy upgrades it.

.PARAMETER InstallPath
    Where to put the files. Defaults to AlignPro under your local application data.

.PARAMETER Source
    Folder holding the add-in files. Defaults to the folder this script is in, which is where they sit
    when you extract the release zip.

.PARAMETER Version
    Version being installed, recorded in Add/Remove Programs. Optional.

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
    [string] $Version,
    [switch] $Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$addInName = 'AlignPro.AddIn'
$registryKey = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$addInName"
$uninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\AlignPro'
$script:vstoWarning = $false

# The complete set. Anything missing means an incomplete download rather than a broken machine.
$required = @(
    'AlignPro.AddIn.dll'
    'AlignPro.AddIn.dll.manifest'
    'AlignPro.AddIn.vsto'
    'AlignPro.Geometry.dll'
    'Microsoft.Office.Tools.Common.v4.0.Utilities.dll'
)

# Carried along when the package has them, which the release zip always does. Absent when installing
# straight out of a build folder, which is a developer's case and not worth failing over.
$extras = @(
    'Uninstall-AlignPro.ps1'
    'AlignPro-Sample.pptx'
)

# Run by path (the .cmd, or by hand) this is a script file; handed over by install.ps1 it is a script
# block inside someone's interactive session, where `exit` closes their PowerShell window before they
# can read why. Captured here because inside Fail $MyInvocation describes the function instead.
$runAsFile = $MyInvocation.MyCommand.CommandType -eq 'ExternalScript'

function Fail {
    param([string] $Message, [string] $Remedy)
    Write-Host ''
    Write-Host $Message -ForegroundColor Red
    if ($Remedy) { Write-Host $Remedy -ForegroundColor Yellow }
    Write-Host ''
    if ($runAsFile) { exit 1 }

    # Already said above, so the caller should stop without saying it again.
    $reported = [System.Exception]::new($Message)
    $reported.Data['AlignPro.Reported'] = $true
    throw $reported
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
# Every check below reads HKLM in BOTH registry views explicitly, rather than through an HKLM: path.
# A path is interpreted relative to the bitness of whatever PowerShell the user happened to launch:
# 64-bit PowerShell sees the native view, 32-bit PowerShell is redirected into WOW6432Node and cannot
# see the native one at all. Since Office, the VSTO runtime and Windows do not agree on which view
# they register in, a path-based check passes or fails depending on which shell was opened - which is
# exactly the kind of false negative that blocked a working machine and sent someone to a dead link.
function Test-RegistryKey {
    param([string] $SubKey, [string] $ValueName)
    foreach ($view in @('Registry64', 'Registry32')) {
        try {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
            try {
                $key = $base.OpenSubKey($SubKey)
                if ($key) {
                    try {
                        $value = if ($ValueName) { $key.GetValue($ValueName) } else { $null }
                        return [pscustomobject]@{ Found = $true; Value = $value; View = $view }
                    }
                    finally { $key.Dispose() }
                }
            }
            finally { $base.Dispose() }
        }
        catch {
            # Registry64 does not exist on 32-bit Windows; try the other view rather than giving up.
        }
    }
    return [pscustomobject]@{ Found = $false; Value = $null; View = $null }
}

# PowerPoint first, because it is the thing AlignPro attaches to. Checking the VSTO runtime before it
# meant a machine with no PowerPoint at all was told the runtime was missing - true, but not the
# reason, and not something installing the runtime would fix.
$powerPoint = Test-RegistryKey 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE'
if (-not $powerPoint.Found) {
    Fail 'PowerPoint does not appear to be installed.' `
         'AlignPro is a PowerPoint add-in and has nothing to attach to. It needs PowerPoint for Windows on the desktop; Microsoft 365 for the web cannot load add-ins of this kind.'
}
Write-Host "  PowerPoint                    present"

# .NET Framework 4.8 is release 528040 or higher, and ships with Windows 10 1903 onwards.
$ndp = Test-RegistryKey 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' 'Release'
$release = if ($ndp.Found -and $ndp.Value) { [int] $ndp.Value } else { 0 }
if ($release -lt 528040) {
    Fail 'The .NET Framework 4.8 is required and was not found.' `
         'Install it from https://dotnet.microsoft.com/download/dotnet-framework/net48 and run this again.'
}
Write-Host "  .NET Framework 4.8            present"

# The VSTO runtime ships with Office, and registers under v4R, or only v4, or only in the 32-bit view,
# depending on the machine - so look for all of it, plus the installer on disk, before concluding
# anything. Measured on one Click-to-Run machine: nothing at all in the native view, v4 and v4R both
# present in WOW6432Node, and VSTOInstaller.exe under both Program Files trees.
$vstoVersion = $null
foreach ($key in @('SOFTWARE\Microsoft\VSTO Runtime Setup\v4R', 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4')) {
    $probe = Test-RegistryKey $key 'Version'
    if ($probe.Found) {
        $vstoVersion = if ($probe.Value) { $probe.Value } else { 'present' }
        break
    }
}
if (-not $vstoVersion) {
    foreach ($base in @($env:CommonProgramFiles, ${env:CommonProgramFiles(x86)})) {
        if ($base -and (Test-Path (Join-Path $base 'Microsoft Shared\VSTO\10.0\VSTOInstaller.exe'))) {
            $vstoVersion = 'present (found on disk, not in the registry)'
            break
        }
    }
}

if ($vstoVersion) {
    Write-Host "  VSTO runtime                  $vstoVersion"
}
else {
    # A warning, not a failure. Office installs this runtime, so on a machine that has PowerPoint its
    # apparent absence is more often a detection gap than a real one - and refusing to install strands
    # someone whose machine is fine. Installing anyway costs a few files and some HKCU values, and if
    # the runtime really is missing the add-in simply does not appear, which the note below covers.
    $script:vstoWarning = $true
    Write-Host "  VSTO runtime                  NOT DETECTED - continuing anyway" -ForegroundColor Yellow
}

# --- PowerPoint must not be holding the files ----------------------------------------------------
if ((Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue) -and -not $Force) {
    Fail 'PowerPoint is running.' 'Close PowerPoint and run this again, so the files are not locked.'
}

# --- copy ----------------------------------------------------------------------------------------
if (-not (Test-Path $InstallPath)) { New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null }

$installed = @()
foreach ($file in $required) {
    Copy-Item (Join-Path $Source $file) (Join-Path $InstallPath $file) -Force
    $installed += $file
}
foreach ($file in $extras) {
    $from = Join-Path $Source $file
    if (Test-Path $from) {
        Copy-Item $from (Join-Path $InstallPath $file) -Force
        $installed += $file
    }
}
Write-Host ''
Write-Host "  installed to                  $InstallPath"

# Windows marks anything that arrived from the internet, and that mark survives both unzipping and
# copying. The .NET loader refuses to load a marked assembly, PowerPoint gives up and sets
# LoadBehavior to 2, and the add-in simply never appears - with nothing to show why.
#
# Nothing should be marked when the package came through install.ps1: Invoke-WebRequest writes no
# Zone.Identifier, so the zip has no mark to pass on. A zip the browser downloaded does, and Explorer's
# extractor copies it onto every file, so this stays as the safety net for the hand-unzip route.
$blocked = 0
foreach ($file in $installed) {
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
$manifestUrl = 'file:///' + ((Join-Path $InstallPath 'AlignPro.AddIn.vsto') -replace '\\', '/')
# PowerPoint's registration wants the |vstolocal suffix; the inclusion list wants the bare URL.
$manifest = $manifestUrl + '|vstolocal'
Set-ItemProperty -Path $registryKey -Name 'FriendlyName' -Value 'AlignPro'
Set-ItemProperty -Path $registryKey -Name 'Description'  -Value 'Align, distribute, size and tidy shapes in PowerPoint'
# Always force 3. If a previous attempt failed - blocked files being the usual reason - PowerPoint
# will have set this to 2, meaning "do not try again", and would ignore the add-in forever otherwise.
Set-ItemProperty -Path $registryKey -Name 'LoadBehavior' -Value 3 -Type DWord
Set-ItemProperty -Path $registryKey -Name 'Manifest'     -Value $manifest

Write-Host "  registered for                $env:USERNAME (no admin rights used)"

# --- grant trust ---------------------------------------------------------------------------------
# The public key is read out of the manifest being installed rather than shipped beside it, so the
# entry can never drift from the build it is meant to trust.
$inclusionRoot = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
try {
    [xml] $manifestXml = Get-Content (Join-Path $InstallPath 'AlignPro.AddIn.vsto') -Raw
    $ns = New-Object System.Xml.XmlNamespaceManager($manifestXml.NameTable)
    $ns.AddNamespace('ds', 'http://www.w3.org/2000/09/xmldsig#')
    $rsa = $manifestXml.SelectSingleNode('//ds:Signature/ds:KeyInfo/ds:KeyValue/ds:RSAKeyValue', $ns)
    if (-not $rsa) { throw 'the manifest carries no signature public key' }

    $publicKey = '<RSAKeyValue><Modulus>' + $rsa.Modulus + '</Modulus><Exponent>' + $rsa.Exponent + '</Exponent></RSAKeyValue>'
    if (-not (Test-Path $inclusionRoot)) { New-Item -Path $inclusionRoot -Force | Out-Null }

    # Drop any entry for this same manifest first, so reinstalling does not accumulate duplicates.
    Get-ChildItem $inclusionRoot -ErrorAction SilentlyContinue | ForEach-Object {
        if ((Get-ItemProperty $_.PSPath).Url -eq $manifestUrl) { Remove-Item -Path $_.PSPath -Recurse -Force }
    }

    $entry = Join-Path $inclusionRoot ([Guid]::NewGuid().ToString('B'))
    New-Item -Path $entry -Force | Out-Null
    Set-ItemProperty -Path $entry -Name 'Url' -Value $manifestUrl
    Set-ItemProperty -Path $entry -Name 'PublicKey' -Value $publicKey

    Write-Host "  trusted                       this add-in only (no certificate store touched)"
}
catch {
    Fail "Could not grant trust: $($_.Exception.Message)" `
         'Without it PowerPoint loads nothing and gives no reason. The files are installed; re-run this script to retry.'
}

# --- Add/Remove Programs -------------------------------------------------------------------------
# So removal is "Settings > Apps > Uninstall" like anything else, rather than a script the user has to
# still have lying around. Only worth writing when the uninstaller travelled with the package, since
# the entry is a promise that it is there.
$uninstaller = Join-Path $InstallPath 'Uninstall-AlignPro.ps1'
if (Test-Path $uninstaller) {
    if (-not (Test-Path $uninstallKey)) { New-Item -Path $uninstallKey -Force | Out-Null }
    $bytes = (Get-ChildItem -LiteralPath $InstallPath -File | Measure-Object -Property Length -Sum).Sum

    Set-ItemProperty -Path $uninstallKey -Name 'DisplayName'     -Value 'AlignPro'
    Set-ItemProperty -Path $uninstallKey -Name 'Publisher'       -Value 'AlignPro'
    Set-ItemProperty -Path $uninstallKey -Name 'InstallLocation' -Value $InstallPath
    Set-ItemProperty -Path $uninstallKey -Name 'DisplayIcon'     -Value (Join-Path $InstallPath 'AlignPro.AddIn.dll')
    Set-ItemProperty -Path $uninstallKey -Name 'URLInfoAbout'    -Value 'https://github.com/MrR1974/AlignPro'
    Set-ItemProperty -Path $uninstallKey -Name 'EstimatedSize'   -Value ([int]($bytes / 1KB)) -Type DWord
    Set-ItemProperty -Path $uninstallKey -Name 'NoModify'        -Value 1 -Type DWord
    Set-ItemProperty -Path $uninstallKey -Name 'NoRepair'        -Value 1 -Type DWord
    if ($Version) { Set-ItemProperty -Path $uninstallKey -Name 'DisplayVersion' -Value $Version }

    # -FromArp tells the uninstaller it was launched by Windows rather than by a person at a prompt,
    # so it holds the window open instead of vanishing with its output.
    $command = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}" -FromArp' -f $uninstaller
    Set-ItemProperty -Path $uninstallKey -Name 'UninstallString' -Value $command

    Write-Host "  listed in                     Settings > Apps > Installed apps"
}

Write-Host ''
Write-Host 'Done. Start PowerPoint and look for the AlignPro tab.' -ForegroundColor Green
Write-Host ''
if (Test-Path (Join-Path $InstallPath 'AlignPro-Sample.pptx')) {
    Write-Host 'A sample deck is installed alongside it, one slide per capability:' -ForegroundColor DarkGray
    Write-Host "  $(Join-Path $InstallPath 'AlignPro-Sample.pptx')" -ForegroundColor DarkGray
    Write-Host ''
}
Write-Host 'If the tab does not appear, check File > Options > Add-ins > Disabled Items.' -ForegroundColor DarkGray
Write-Host ''
if ($script:vstoWarning) {
    Write-Host 'One thing was not confirmed: the Visual Studio Tools for Office runtime.' -ForegroundColor Yellow
    Write-Host 'It ships with Office and is almost certainly present - this check has been wrong before -' -ForegroundColor DarkGray
    Write-Host 'so the install went ahead. If the AlignPro tab does not appear, that is the thing to fix:' -ForegroundColor DarkGray
    Write-Host '  https://aka.ms/VSTORuntimeDownload' -ForegroundColor DarkGray
    Write-Host 'Then run the installer again. Nothing needs uninstalling first.' -ForegroundColor DarkGray
    Write-Host ''
}
