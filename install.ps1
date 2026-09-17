<#
    AlignPro - remote installer

        irm https://raw.githubusercontent.com/MrR1974/AlignPro/main/install.ps1 | iex

    Downloads the latest release, checks it against its published SHA-256, and hands the extracted
    folder to the installer inside it. Nothing is installed by this file itself: it is a courier.

    Why this rather than a downloaded installer
    -------------------------------------------
    SmartScreen's reputation check applies to executables the browser downloads, and an unsigned file
    has no reputation - Microsoft's own documentation is explicit that reputation "must build for each
    new version of your files, starting with zero", so an unsigned MSI is warned about on every
    release, forever. A self-signed certificate behaves identically to no signature.

    Fetching over HTTPS from inside PowerShell avoids that path entirely, and also avoids the mark of
    the web: Invoke-WebRequest writes no Zone.Identifier stream, and Expand-Archive does not propagate
    one, so no file ever needs unblocking. A script downloaded to disk and run by path would be refused
    by the AuthorizationManager check even with the execution policy at Unrestricted - which is the
    problem this replaces.

    What the hash does and does not prove
    -------------------------------------
    The zip and its checksum come from the same release, so this is an integrity check, not a
    supply-chain one: it catches a truncated or corrupted download, not a compromised repository. The
    trust anchor here is HTTPS and GitHub, exactly as it is for the script you are reading.

    Overrides, for the rare case you need one:

        $env:ALIGNPRO_VERSION = '1.0.0'   # install a specific release instead of the latest
        $env:ALIGNPRO_PATH    = 'D:\...'  # install somewhere other than %LOCALAPPDATA%\AlignPro
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = 'MrR1974/AlignPro'
$staging = Join-Path ([System.IO.Path]::GetTempPath()) ("AlignPro-" + [Guid]::NewGuid().ToString('N').Substring(0, 8))

function Fail {
    param([string] $Message, [string] $Remedy)
    Write-Host ''
    Write-Host $Message -ForegroundColor Red
    if ($Remedy) { Write-Host $Remedy -ForegroundColor Yellow }
    Write-Host ''
    exit 1
}

Write-Host ''
Write-Host 'AlignPro' -ForegroundColor Cyan
Write-Host ''

# Windows PowerShell 5.1 still negotiates TLS 1.0 by default on some builds, which github.com refuses.
try {
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
}
catch {
    # .NET Core picks the protocol itself and the property may be read-only; nothing to do.
}

try {
    # --- find the release ------------------------------------------------------------------------
    $wanted = $env:ALIGNPRO_VERSION
    $api = if ($wanted) {
        "https://api.github.com/repos/$repo/releases/tags/v$wanted"
    } else {
        "https://api.github.com/repos/$repo/releases/latest"
    }

    $headers = @{ 'User-Agent' = 'AlignPro-Installer'; 'Accept' = 'application/vnd.github+json' }
    $release = Invoke-RestMethod -Uri $api -Headers $headers -UseBasicParsing

    $version = $release.tag_name -replace '^v', ''
    Write-Host "  release                       $($release.tag_name)"

    $zipName = "AlignPro-$version.zip"
    $sumName = "$zipName.sha256"
    $zipAsset = $release.assets | Where-Object { $_.name -eq $zipName } | Select-Object -First 1
    $sumAsset = $release.assets | Where-Object { $_.name -eq $sumName } | Select-Object -First 1

    if (-not $zipAsset) { Fail "Release $($release.tag_name) has no asset named $zipName." 'The release may still be uploading. Try again shortly.' }
    if (-not $sumAsset) { Fail "Release $($release.tag_name) has no checksum asset named $sumName." 'Refusing to install something that cannot be checked.' }

    # --- download --------------------------------------------------------------------------------
    New-Item -ItemType Directory -Path $staging -Force | Out-Null
    $zipFile = Join-Path $staging $zipName
    $sumFile = Join-Path $staging $sumName

    # Invoke-WebRequest is markedly faster with the progress bar suppressed, and its output here is
    # noise in front of the installer's own.
    $previousProgress = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    try {
        Invoke-WebRequest -Uri $zipAsset.browser_download_url -OutFile $zipFile -Headers $headers -UseBasicParsing
        Invoke-WebRequest -Uri $sumAsset.browser_download_url -OutFile $sumFile -Headers $headers -UseBasicParsing
    }
    finally {
        $ProgressPreference = $previousProgress
    }
    Write-Host ("  downloaded                    {0:N0} KB" -f ((Get-Item $zipFile).Length / 1KB))

    # --- verify ----------------------------------------------------------------------------------
    # The checksum file is written sha256sum-style: "<hash>  <filename>".
    $expected = (((Get-Content -LiteralPath $sumFile -Raw) -split '\s+') | Where-Object { $_ })[0].ToLowerInvariant()
    $actual = (Get-FileHash -LiteralPath $zipFile -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($expected -ne $actual) {
        Fail "The download does not match its published checksum.`n  expected  $expected`n  got       $actual" `
             'Nothing has been installed. This is usually an interrupted download - run it again.'
    }
    Write-Host "  verified                      SHA-256 matches the published checksum"

    # --- extract ---------------------------------------------------------------------------------
    $extracted = Join-Path $staging 'package'
    Expand-Archive -LiteralPath $zipFile -DestinationPath $extracted -Force

    $installer = Join-Path $extracted 'Install-AlignPro.ps1'
    if (-not (Test-Path $installer)) { Fail "The package does not contain Install-AlignPro.ps1." 'The release asset looks wrong; please report it.' }

    # --- hand over -------------------------------------------------------------------------------
    # Run the installer as a script block built from its text rather than by path. Executing a .ps1
    # file is governed by the execution policy - which is Restricted by default on Windows PowerShell,
    # and would refuse - while a script block created here is not. Same code, same parameters, no
    # policy to argue with and no child process.
    $script = [ScriptBlock]::Create((Get-Content -LiteralPath $installer -Raw))
    $arguments = @{ Source = $extracted; Version = $version }
    if ($env:ALIGNPRO_PATH) { $arguments['InstallPath'] = $env:ALIGNPRO_PATH }

    & $script @arguments
}
catch {
    Fail "Installation failed: $($_.Exception.Message)" 'Nothing was left behind. Please report this if it repeats.'
}
finally {
    if (Test-Path $staging) { Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue }
}
