<#
.SYNOPSIS
    Builds AlignPro in Release and packages a zip to attach to a GitHub release.

.DESCRIPTION
    Produces the zip an end user downloads: the five add-in files, the install and uninstall scripts,
    the sample deck and a short readme. Nothing in it needs a compiler, Visual Studio or a certificate
    to use.

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

# --- stage ---------------------------------------------------------------------------------------
$staging = Join-Path ([System.IO.Path]::GetTempPath()) ("AlignPro-$Version-" + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $staging | Out-Null

foreach ($file in $payload) { Copy-Item (Join-Path $binaries $file) $staging -Force }
Copy-Item (Join-Path $PSScriptRoot 'Install-AlignPro.ps1') $staging -Force
Copy-Item (Join-Path $PSScriptRoot 'Uninstall-AlignPro.ps1') $staging -Force

# The sample deck ships in the zip so it is to hand the moment someone installs, rather than being a
# separate trip back to the repository. Fail rather than quietly shipping without it.
$deck = Join-Path $projectRoot 'sample\AlignPro-Sample.pptx'
if (-not (Test-Path $deck)) {
    throw "The sample deck is missing from '$deck'. Run tools\New-SampleDeck.ps1 first."
}
Copy-Item $deck $staging -Force

@"
AlignPro $Version
=================

Align, distribute, size and tidy shapes in PowerPoint - the things the built-in
tools will not do: aligning to a shape you choose, exact numeric spacing,
matching sizes, grid tidying, and aligning on what you actually see rather than
on PowerPoint's own rectangle, which ignores rotation.


To install
----------

1. Close PowerPoint.
2. Right-click Install-AlignPro.ps1 and choose "Run with PowerShell".
   Or from a PowerShell prompt:  .\Install-AlignPro.ps1
3. Start PowerPoint. There will be an AlignPro tab on the ribbon.

No administrator rights are needed. Nothing is installed machine-wide: the files
are copied to your local application data and one registry value is written
under HKEY_CURRENT_USER.

If PowerShell refuses to run the script, it is the execution policy rather than
anything to do with AlignPro. This runs it for that one command only:

    powershell -ExecutionPolicy Bypass -File .\Install-AlignPro.ps1


To remove
---------

Close PowerPoint, then run Uninstall-AlignPro.ps1. It removes the registry entry
and the files, and touches nothing else.


About signing
-------------

AlignPro is not signed by a certificate authority, so Windows cannot tell you who
published it. The source is on GitHub and can be read and built by anyone who
would rather verify it than take it on trust.

Installation does not involve a certificate at all: the add-in is loaded from
your own disk rather than installed as a ClickOnce package.


Try it
------

AlignPro-Sample.pptx is included here: ten slides, one per capability, each
captioned with what to try and what should happen. Slide 2 is the one to start
with - align left with Measure = Shape frame, undo, then again with Visual
bounds, and watch the rotated shape.


One thing worth knowing
-----------------------

Use AlignPro's own Undo button rather than Ctrl+Z after an AlignPro command.
PowerPoint groups changes made by add-ins into a single undo entry that can cover
far more than your last action.
"@ | Set-Content -Path (Join-Path $staging 'README.txt') -Encoding UTF8

# --- zip -----------------------------------------------------------------------------------------
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$zip = Join-Path $OutputPath "AlignPro-$Version.zip"
if (Test-Path $zip) { Remove-Item -LiteralPath $zip -Force }
Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $zip -CompressionLevel Optimal
[System.IO.Directory]::Delete($staging, $true)

$size = (Get-Item $zip).Length
Write-Host ''
Write-Host "Packaged: $zip" -ForegroundColor Green
Write-Host ("  {0:N0} KB, {1} files" -f ($size / 1KB), ($payload.Count + 4))
Write-Host ''
Write-Host 'Next: create a GitHub release and attach that zip.' -ForegroundColor Cyan
Write-Host '  With the gh CLI:' -ForegroundColor DarkGray
Write-Host ("    gh release create v$Version `"$zip`" --title `"AlignPro $Version`" --notes-file <notes.md>") -ForegroundColor DarkGray
Write-Host ''
