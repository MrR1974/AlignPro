<#
.SYNOPSIS
    Tests, builds Release, and packages everything a GitHub release carries.

.DESCRIPTION
    Runs the geometry tests, builds Release, and writes three files to dist\:

        AlignPro-<version>.zip          the release payload - what install.ps1 downloads
        AlignPro-<version>.zip.sha256   its checksum, which install.ps1 verifies
        AlignPro-<version>.msi          the same install, for managed deployment

    The zip is the distribution route. The MSI is kept for Intune and Group Policy, where an installer
    package is what the tooling expects, and where SmartScreen is not in the path at all. It is not
    what a person downloads: an unsigned MSI is warned about on every release forever, because
    SmartScreen reputation for an unsigned file starts at zero for each new build and a self-signed
    certificate counts as no signature.

    The build needs Visual Studio's MSBuild and a signing certificate, because VSTO refuses to produce
    manifests unsigned. That certificate is a build-time formality: it is self-signed, it never leaves
    this machine, and what installs trust is the inclusion-list entry, not the certificate. Run
    New-DevSigningCertificate.ps1 once if the build complains about ClickOnce manifest signing.

    Deliberately publishes nothing. It writes the files and tells you where they are; creating the
    GitHub release is a separate, deliberate act.

.PARAMETER Version
    Version string, e.g. 1.0.0.

.PARAMETER OutputPath
    Where to write the output. Defaults to dist\ beside the project, which is gitignored.

.PARAMETER SkipInstaller
    Skip the MSI. Useful when WiX is not installed and only the zip is wanted.

.EXAMPLE
    .\New-Release.ps1 -Version 1.0.0
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string] $Version,

    [string] $OutputPath,

    [switch] $SkipInstaller
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path $PSScriptRoot -Parent
$solution = Join-Path $projectRoot 'AlignPro.sln'
if (-not $OutputPath) { $OutputPath = Join-Path $projectRoot 'dist' }

$payload = @(
    'AlignPro.AddIn.dll'
    'AlignPro.AddIn.dll.manifest'
    'AlignPro.AddIn.vsto'
    'AlignPro.Geometry.dll'
    'Microsoft.Office.Tools.Common.v4.0.Utilities.dll'
)

# --- locate MSBuild ------------------------------------------------------------------------------
# The dotnet CLI cannot load a VSTO project - its $(VSToolsPath) import resolves into the dotnet SDK -
# so this has to be Visual Studio's MSBuild.
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path $vswhere)) { throw 'Visual Studio is not installed; a VSTO project cannot be built without it.' }
$vsPath = & $vswhere -products * -requires Microsoft.VisualStudio.Workload.Office -property installationPath | Select-Object -First 1
if (-not $vsPath) { throw 'No Visual Studio installation with the Office/SharePoint development workload was found.' }
$msbuild = Join-Path $vsPath 'MSBuild\Current\Bin\MSBuild.exe'
if (-not (Test-Path $msbuild)) { throw "MSBuild not found at $msbuild." }

Write-Host ''
Write-Host "Building AlignPro $Version" -ForegroundColor Cyan
Write-Host "  msbuild: $msbuild" -ForegroundColor DarkGray

# --- test, then build ----------------------------------------------------------------------------
# The engine tests run on the dotnet CLI and are fast; a release that fails them should not be built.
Write-Host ''
Write-Host 'Running the geometry tests...' -ForegroundColor Cyan
$testProject = Join-Path $projectRoot 'tests\AlignPro.Geometry.Tests\AlignPro.Geometry.Tests.csproj'
& dotnet test $testProject --nologo -v q
if ($LASTEXITCODE -ne 0) { throw 'Tests failed. Release aborted.' }

Write-Host ''
Write-Host 'Building Release...' -ForegroundColor Cyan
& $msbuild $solution /p:Configuration=Release /p:VisualStudioVersion=17.0 /v:minimal /nologo
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }

$binaries = Join-Path $projectRoot 'src\AlignPro.AddIn\bin\Release'
$missing = $payload | Where-Object { -not (Test-Path (Join-Path $binaries $_)) }
if ($missing) { throw "Build output is missing: $($missing -join ', ')" }

# --- package -------------------------------------------------------------------------------------
& (Join-Path $PSScriptRoot 'New-Package.ps1') -Version $Version -Configuration Release
$zip = Join-Path $OutputPath "AlignPro-$Version.zip"
$sum = "$zip.sha256"
if (-not (Test-Path $zip)) { throw 'The release zip was not produced.' }

$assets = @($zip, $sum)

if (-not $SkipInstaller) {
    & (Join-Path $PSScriptRoot 'New-Installer.ps1') -Version $Version -Configuration Release
    $msi = Join-Path $OutputPath "AlignPro-$Version.msi"
    if (-not (Test-Path $msi)) { throw 'The MSI was not produced.' }
    $assets += $msi
}

Write-Host ''
Write-Host 'Next: create a GitHub release and attach these.' -ForegroundColor Cyan
Write-Host '  The zip and its .sha256 must both be attached, and named exactly as they are here -' -ForegroundColor DarkGray
Write-Host '  install.ps1 looks them up by name and refuses to install without the checksum.' -ForegroundColor DarkGray
Write-Host ''
Write-Host '  With the gh CLI:' -ForegroundColor DarkGray
Write-Host "    gh release create v$Version ``" -ForegroundColor DarkGray
foreach ($asset in $assets) {
    Write-Host "        `"$asset`" ``" -ForegroundColor DarkGray
}
Write-Host "        --title `"AlignPro $Version`" --notes-file <notes.md>" -ForegroundColor DarkGray
Write-Host ''
