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
    the web: Invoke-WebRequest writes no Zone.Identifier stream, and .NET's ZipFile does not propagate
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
        $env:ALIGNPRO_NO_TUTORIAL = '1'   # do not open the tutorial deck when the install finishes
#>

# `irm | iex` runs this text inside the user's own interactive session, not as a script. Two things
# follow. `exit` there closes their PowerShell window, taking the error message with it - so nothing
# below exits unless this really is running as a file. And anything set at this level would stay set
# in their session afterwards - so it all lives in a script block, whose scope ends with it.
& {
    param([bool] $RunAsFile)

    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    $repo = 'MrR1974/AlignPro'
    $staging = Join-Path ([System.IO.Path]::GetTempPath()) ("AlignPro-" + [Guid]::NewGuid().ToString('N').Substring(0, 8))

    # Says what went wrong, then stops by throwing, which the catch at the bottom recognises and does
    # not repeat. Install-AlignPro.ps1 reports its own failures the same way.
    function Report {
        param([string] $Message, [string] $Remedy)
        Write-Host ''
        Write-Host $Message -ForegroundColor Red
        if ($Remedy) { Write-Host $Remedy -ForegroundColor Yellow }
        Write-Host ''
    }

    function Fail {
        param([string] $Message, [string] $Remedy)
        Report $Message $Remedy
        $reported = [System.Exception]::new($Message)
        $reported.Data['AlignPro.Reported'] = $true
        throw $reported
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
        # .NET rather than Get-FileHash, for the reason given at the extract below.
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $stream = [System.IO.File]::OpenRead($zipFile)
        try { $actual = -join ($sha.ComputeHash($stream) | ForEach-Object { $_.ToString('x2') }) }
        finally { $stream.Dispose(); $sha.Dispose() }
        if ($expected -ne $actual) {
            Fail "The download does not match its published checksum.`n  expected  $expected`n  got       $actual" `
                 'Nothing has been installed. This is usually an interrupted download - run it again.'
        }
        Write-Host "  verified                      SHA-256 matches the published checksum"

        # --- extract ---------------------------------------------------------------------------------
        # .NET directly rather than Expand-Archive, which has been seen to hang here. Expand-Archive and
        # Get-FileHash both live in modules PowerShell loads on first use, and finding a module means
        # searching every folder on PSModulePath - the first of which is usually Documents, redirected
        # to OneDrive on managed machines, where a paused or signing-in OneDrive stalls the search. A
        # PSModulePath inherited from PowerShell 7 breaks it differently, by offering 5.1 modules it
        # cannot load. Expand-Archive also draws a progress bar that is slow on Windows PowerShell 5.1.
        # ExtractToDirectory has none of these problems, and refuses entries that would land outside
        # the destination folder.
        $extracted = Join-Path $staging 'package'
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipFile, $extracted)
        Write-Host "  extracted                     $extracted"

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

        # --- open the tutorial -----------------------------------------------------------------------
        # The deck is the fastest way to understand what the add-in does, and someone who has just run a
        # one-line installer has nothing else in front of them. Failing to open it is never fatal: the
        # install has already succeeded by this point, and the installer has printed the path.
        if (-not $env:ALIGNPRO_NO_TUTORIAL) {
            $installedTo = if ($env:ALIGNPRO_PATH) { $env:ALIGNPRO_PATH } else { Join-Path $env:LOCALAPPDATA 'AlignPro' }
            $deck = Join-Path $installedTo 'AlignPro-Sample.pptx'

            if (Test-Path -LiteralPath $deck) {
                # PowerPoint loads add-ins when it starts, so a copy that was already running has no
                # AlignPro tab in it. Opening the tutorial there would show the slides telling you to
                # click a tab that is not on screen, which reads as a broken install.
                if (Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue) {
                    Write-Host 'PowerPoint is already running, so the AlignPro tab is not in it yet.' -ForegroundColor Yellow
                    Write-Host 'Close PowerPoint, then open the tutorial to get started:' -ForegroundColor DarkGray
                    Write-Host "  $deck" -ForegroundColor DarkGray
                    Write-Host ''
                }
                else {
                    try {
                        # Start-Process on the deck itself, so Windows picks whatever opens .pptx rather
                        # than this script guessing where PowerPoint is installed.
                        Start-Process -FilePath $deck | Out-Null
                        Write-Host 'Opening the tutorial deck to get you started.' -ForegroundColor DarkGray
                        Write-Host ''
                    }
                    catch {
                        Write-Host "The tutorial deck could not be opened: $($_.Exception.Message)" -ForegroundColor DarkYellow
                        Write-Host "  $deck" -ForegroundColor DarkGray
                        Write-Host ''
                    }
                }
            }
        }
    }
    catch {
        if (-not $_.Exception.Data['AlignPro.Reported']) {
            Report "Installation failed: $($_.Exception.Message)" 'Nothing was left behind. Please report this if it repeats.'
        }
        if ($RunAsFile) { exit 1 }
    }
    finally {
        if (Test-Path $staging) { Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue }
    }
} ($MyInvocation.MyCommand.CommandType -eq 'ExternalScript')
