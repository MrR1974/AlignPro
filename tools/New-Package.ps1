<#
.SYNOPSIS
    Packages an already-built Release into the release zip and its checksum.

.DESCRIPTION
    Produces the two assets a GitHub release carries:

        dist\AlignPro-<version>.zip
        dist\AlignPro-<version>.zip.sha256

    The zip is what install.ps1 downloads, and it is also the file a person can download and unzip by
    hand. It holds the add-in, the sample deck, both scripts and a double-click shim.

    A zip rather than an MSI, deliberately. SmartScreen's reputation check applies to executables the
    browser downloads, and an unsigned file has no reputation to check: Microsoft documents that
    reputation "must build for each new version of your files, starting with zero", and that a
    self-signed certificate behaves exactly like no signature. An unsigned MSI is therefore warned
    about on every release forever, and no amount of download volume fixes it. Nothing here is
    executable, and install.ps1 never involves the browser at all.

    The checksum is written sha256sum-style ("<hash>  <filename>") so it is readable both by
    install.ps1 and by anyone who wants to check the download themselves. It proves the download
    arrived intact; it is not a supply-chain guarantee, since it travels in the same release as the
    file it describes.

.PARAMETER Version
    Version being packaged, e.g. 1.0.0.

.PARAMETER Configuration
    Build configuration to package. Defaults to Release.

.EXAMPLE
    .\New-Package.ps1 -Version 1.0.0
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string] $Version,

    [string] $Configuration = 'Release'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path $PSScriptRoot -Parent
$binDir = Join-Path $projectRoot "src\AlignPro.AddIn\bin\$Configuration"
$deck = Join-Path $projectRoot 'sample\AlignPro-Sample.pptx'
$outputDir = Join-Path $projectRoot 'dist'

$addInFiles = @(
    'AlignPro.AddIn.dll'
    'AlignPro.AddIn.dll.manifest'
    'AlignPro.AddIn.vsto'
    'AlignPro.Geometry.dll'
    'Microsoft.Office.Tools.Common.v4.0.Utilities.dll'
)

$missing = $addInFiles | Where-Object { -not (Test-Path (Join-Path $binDir $_)) }
if ($missing) { throw "Build output is missing from '$binDir': $($missing -join ', '). Build $Configuration first." }
if (-not (Test-Path $deck)) { throw "The sample deck is missing from '$deck'. Run tools\New-SampleDeck.ps1 first." }

# --- stage ---------------------------------------------------------------------------------------
$staging = Join-Path ([System.IO.Path]::GetTempPath()) ("AlignProPkg-" + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $staging -Force | Out-Null

try {
    foreach ($file in $addInFiles) { Copy-Item (Join-Path $binDir $file) $staging }
    Copy-Item $deck (Join-Path $staging 'AlignPro-Sample.pptx')
    Copy-Item (Join-Path $PSScriptRoot 'Install-AlignPro.ps1') $staging
    Copy-Item (Join-Path $PSScriptRoot 'Uninstall-AlignPro.ps1') $staging
    Copy-Item (Join-Path $PSScriptRoot 'Install AlignPro.cmd') $staging

    $title = "AlignPro $Version"
    $readme = @"
$title
$('=' * $title.Length)

A PowerPoint add-in for aligning, distributing, sizing and tidying shapes.

INSTALLING

  Close PowerPoint, then double-click "Install AlignPro.cmd".

  If Windows objects to running it, the zip was extracted while still marked as downloaded from the
  internet. Delete this folder, right-click the zip, choose Properties, tick Unblock, click OK, and
  extract it again.

  There is also a one-line install that avoids all of that, because nothing is downloaded by the
  browser. In PowerShell:

    irm https://raw.githubusercontent.com/MrR1974/AlignPro/main/install.ps1 | iex

WHAT IT DOES

  Copies these files to %LOCALAPPDATA%\AlignPro and writes a few registry values under
  HKEY_CURRENT_USER so PowerPoint can find them. No administrator rights are used, and nothing is
  written outside your own user account.

  It also grants trust, because VSTO will not load an add-in whose signing certificate the machine
  does not know, and AlignPro's is self-signed. It does this with a single VSTO inclusion-list entry
  naming this add-in and the key that signed it - not by importing a certificate, which would trust
  everything that certificate ever signs. Uninstalling revokes it.

REMOVING IT

  Settings > Apps > Installed apps > AlignPro > Uninstall.
  Or run Uninstall-AlignPro.ps1 from %LOCALAPPDATA%\AlignPro.

REQUIREMENTS

  Windows with PowerPoint (desktop), .NET Framework 4.8, and the VSTO runtime. The last two are
  already present on any machine that runs Office.

  The installer checks all three. PowerPoint and the .NET Framework are checked strictly; the VSTO
  runtime is only warned about, because it registers itself differently from machine to machine and
  the check has produced false negatives. If it cannot be found the install goes ahead anyway, and
  says what to do should the AlignPro tab not appear.

SOURCE AND LICENCE

  https://github.com/MrR1974/AlignPro   MIT
"@
    Set-Content -Path (Join-Path $staging 'README.txt') -Value $readme -Encoding UTF8

    # --- zip -------------------------------------------------------------------------------------
    if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }
    $zip = Join-Path $outputDir "AlignPro-$Version.zip"
    if (Test-Path $zip) { Remove-Item $zip -Force }

    Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $zip -CompressionLevel Optimal

    # --- checksum --------------------------------------------------------------------------------
    $hash = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash.ToLowerInvariant()
    $sumFile = "$zip.sha256"
    # Two spaces between hash and name is the sha256sum convention, and `sha256sum -c` accepts it.
    # No trailing newline fuss: install.ps1 splits on whitespace and takes the first field.
    Set-Content -Path $sumFile -Value ("{0}  {1}" -f $hash, (Split-Path $zip -Leaf)) -Encoding ASCII

    Write-Host ''
    Write-Host "Packaged: $zip" -ForegroundColor Green
    Write-Host ("  {0:N0} KB" -f ((Get-Item $zip).Length / 1KB))
    Write-Host "  sha256  $hash"
    Write-Host ''
}
finally {
    Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
}
