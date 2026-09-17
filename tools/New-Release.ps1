<#
.SYNOPSIS
    Builds AlignPro in Release and packages a zip to attach to a GitHub release.

.DESCRIPTION
    Runs the tests, builds Release, and produces the MSI an end user downloads. That is a per-user
    install: no administrator rights, an entry in Add/Remove Programs, and deployable to managed
    machines through Intune or Group Policy.

    The build itself does need Visual Studio's MSBuild and a signing certificate, because VSTO refuses
    to produce manifests unsigned. That certificate is a build-time formality only - it is self-signed,
    it never leaves this machine, and installation does not check it. Run
    New-DevSigningCertificate.ps1 once if the build complains about ClickOnce manifest signing.

    Deliberately does not publish anything. It writes a zip and tells you where it is; uploading it to
    a GitHub release is a separate, deliberate act.

.PARAMETER Version
    Version string for the zip name, e.g. 1.0.0.

.PARAMETER OutputPath
    Where to write the zip. Defaults to dist\ beside the project, which is gitignored.

.EXAMPLE
    .\New-Release.ps1 -Version 1.0.0
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^\d+\.\d+\.\d+$')]
    [string] $Version,

    [string] $OutputPath
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
& (Join-Path $PSScriptRoot 'New-Installer.ps1') -Version $Version -Configuration Release
$msi = Join-Path $OutputPath "AlignPro-$Version.msi"
if (-not (Test-Path $msi)) { throw 'The installer was not produced.' }

Write-Host ''
Write-Host 'Next: create a GitHub release and attach that MSI.' -ForegroundColor Cyan
Write-Host '  With the gh CLI:' -ForegroundColor DarkGray
Write-Host ("    gh release create v$Version `"$msi`" --title `"AlignPro $Version`" --notes-file <notes.md>") -ForegroundColor DarkGray
Write-Host ''
