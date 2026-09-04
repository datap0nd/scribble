#requires -Version 5.1
<#
.SYNOPSIS
Promotes one approved public Scribble release to Samsung GitHub.

.DESCRIPTION
Copies the selected stable tag and its commit to the internal repository,
creates the matching internal GitHub Release, uploads ScribbleSetup.exe, and
verifies both the authenticated asset and the anonymous latest-download URL
used by Scribble's updater. Existing internal releases are never overwritten.

.EXAMPLE
.\scripts\Promote-StableRelease.ps1 -Tag v2.0.35 -WhatIf

.EXAMPLE
.\scripts\Promote-StableRelease.ps1 -Tag v2.0.35
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "Medium")]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^v[0-9]+\.[0-9]+\.[0-9]+$")]
    [string]$Tag,

    [string]$InstallerPath = "",

    [string]$PublicRepository =
        "https://github.com/datap0nd/scribble.git",

    [string]$InternalRepository =
        "https://github.sec.samsung.net/r-cunha/scribble.git",

    [string]$InternalGhRepository =
        "github.sec.samsung.net/r-cunha/scribble",

    [string]$InternalInstallerUrl =
        "https://github.sec.samsung.net/r-cunha/scribble/releases/latest/download/ScribbleSetup.exe"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Invoke-CheckedTool {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$FilePath failed with exit code $LASTEXITCODE."
    }
}

function Assert-ScribbleInstaller {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $file = Get-Item -LiteralPath $Path
    if ($file.Length -lt (200 * 1024) -or
        $file.Length -gt (100 * 1024 * 1024)) {
        throw "ScribbleSetup.exe has an unexpected size: $($file.Length) bytes."
    }

    $stream = [System.IO.File]::OpenRead($file.FullName)
    try {
        if ($stream.ReadByte() -ne [byte][char]"M" -or
            $stream.ReadByte() -ne [byte][char]"Z") {
            throw "ScribbleSetup.exe is not a Windows executable."
        }
    }
    finally {
        $stream.Dispose()
    }
}

foreach ($command in @("git", "gh")) {
    if ($null -eq (Get-Command $command -ErrorAction SilentlyContinue)) {
        throw "Required command not found: $command"
    }
}

$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ("scribble-release-promotion-" + [Guid]::NewGuid().ToString("N"))
$temporaryRoot = [System.IO.Path]::GetFullPath($temporaryRoot)
$systemTemp = [System.IO.Path]::GetFullPath(
    [System.IO.Path]::GetTempPath()).TrimEnd("\") + "\"
if (-not $temporaryRoot.StartsWith(
        $systemTemp,
        [StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to use an unexpected temporary path: $temporaryRoot"
}

New-Item -ItemType Directory -Path $temporaryRoot `
    -Confirm:$false -WhatIf:$false | Out-Null

try {
    $releaseUri =
        "https://api.github.com/repos/datap0nd/scribble/releases/tags/" +
        [Uri]::EscapeDataString($Tag)
    $publicReleaseRequest = @{
        UseBasicParsing = $true
        Headers = @{ "User-Agent" = "Scribble-release-promotion" }
        Uri = $releaseUri
    }
    $publicRelease = Invoke-RestMethod @publicReleaseRequest
    if ($publicRelease.draft -or $publicRelease.prerelease) {
        throw "$Tag is not a published stable release."
    }

    $asset = @($publicRelease.assets) |
        Where-Object { $_.name -eq "ScribbleSetup.exe" } |
        Select-Object -First 1
    if ($null -eq $asset) {
        throw "Public release $Tag does not contain ScribbleSetup.exe."
    }

    $approvedInstaller = Join-Path $temporaryRoot "approved-ScribbleSetup.exe"
    $publicAssetRequest = @{
        UseBasicParsing = $true
        Headers = @{ "User-Agent" = "Scribble-release-promotion" }
        Uri = $asset.browser_download_url
        OutFile = $approvedInstaller
    }
    Invoke-WebRequest @publicAssetRequest
    Assert-ScribbleInstaller -Path $approvedInstaller
    $approvedHash = (Get-FileHash -Algorithm SHA256 `
        -LiteralPath $approvedInstaller).Hash

    if ([string]::IsNullOrWhiteSpace($InstallerPath)) {
        $resolvedInstaller = $approvedInstaller
    }
    else {
        $resolvedInstaller =
            (Resolve-Path -LiteralPath $InstallerPath).Path
        Assert-ScribbleInstaller -Path $resolvedInstaller
        $localHash = (Get-FileHash -Algorithm SHA256 `
            -LiteralPath $resolvedInstaller).Hash
        if ($localHash -ne $approvedHash) {
            throw (
                "The supplied installer does not match ScribbleSetup.exe " +
                "from public release $Tag.")
        }
    }

    $sourceHash = $approvedHash

    $mirrorPath = Join-Path $temporaryRoot "scribble.git"
    Invoke-CheckedTool -FilePath "git" -Arguments @(
        "clone",
        "--bare",
        $PublicRepository,
        $mirrorPath
    )

    $tagReference = "refs/tags/$Tag"
    $tagCommit = (& git -C $mirrorPath rev-parse --verify `
        ($tagReference + "^{commit}")).Trim()
    if ($LASTEXITCODE -ne 0 -or
        $tagCommit -notmatch "^[0-9a-f]{40,64}$") {
        throw "Stable tag $Tag could not be resolved to a commit."
    }

    $updaterObject =
        $tagReference + ":src/Scribble/Utilities/SelfUpdater.cs"
    $updaterSource = (& git -C $mirrorPath show $updaterObject) -join "`n"
    if ($LASTEXITCODE -ne 0 -or
        -not $updaterSource.Contains($InternalInstallerUrl)) {
        throw (
            "$Tag was built before Scribble's internal updater URL was " +
            "added. Publish a new stable release from the current main " +
            "branch, then promote that tag.")
    }

    Invoke-CheckedTool -FilePath "git" -Arguments @(
        "-C",
        $mirrorPath,
        "update-ref",
        "refs/heads/stable-promotion",
        $tagCommit
    )
    Invoke-CheckedTool -FilePath "git" -Arguments @(
        "-C",
        $mirrorPath,
        "remote",
        "add",
        "internal",
        $InternalRepository
    )

    & gh auth status --hostname github.sec.samsung.net
    if ($LASTEXITCODE -ne 0) {
        throw "GitHub CLI is not authenticated to github.sec.samsung.net."
    }

    & gh release view $Tag --repo $InternalGhRepository *> $null
    if ($LASTEXITCODE -eq 0) {
        throw "Internal release $Tag already exists. Nothing was overwritten."
    }

    $releaseName = [string]$publicRelease.name
    if ([string]::IsNullOrWhiteSpace($releaseName)) {
        $releaseName = "Scribble $Tag"
    }
    $notesPath = Join-Path $temporaryRoot "release-notes.md"
    [System.IO.File]::WriteAllText(
        $notesPath,
        [string]$publicRelease.body,
        (New-Object System.Text.UTF8Encoding($false)))

    $description =
        "$Tag to $InternalGhRepository with ScribbleSetup.exe"
    if (-not $PSCmdlet.ShouldProcess(
            $description,
            "Promote stable Scribble release")) {
        return
    }

    Invoke-CheckedTool -FilePath "git" -Arguments @(
        "-C",
        $mirrorPath,
        "push",
        "internal",
        "refs/heads/stable-promotion:refs/heads/main",
        ($tagReference + ":" + $tagReference)
    )

    Invoke-CheckedTool -FilePath "gh" -Arguments @(
        "release",
        "create",
        $Tag,
        $resolvedInstaller,
        "--repo",
        $InternalGhRepository,
        "--verify-tag",
        "--latest",
        "--title",
        $releaseName,
        "--notes-file",
        $notesPath
    )

    $authenticatedDirectory = Join-Path $temporaryRoot "authenticated"
    New-Item -ItemType Directory -Path $authenticatedDirectory | Out-Null
    Invoke-CheckedTool -FilePath "gh" -Arguments @(
        "release",
        "download",
        $Tag,
        "--repo",
        $InternalGhRepository,
        "--pattern",
        "ScribbleSetup.exe",
        "--dir",
        $authenticatedDirectory
    )
    $authenticatedInstaller = Join-Path $authenticatedDirectory `
        "ScribbleSetup.exe"
    Assert-ScribbleInstaller -Path $authenticatedInstaller
    $authenticatedHash = (Get-FileHash -Algorithm SHA256 `
        -LiteralPath $authenticatedInstaller).Hash
    if ($authenticatedHash -ne $sourceHash) {
        throw "The internal release asset does not match the approved installer."
    }

    $anonymousInstaller = Join-Path $temporaryRoot "anonymous-download.exe"
    try {
        $anonymousRequest = @{
            UseBasicParsing = $true
            Uri = $InternalInstallerUrl
            OutFile = $anonymousInstaller
        }
        Invoke-WebRequest @anonymousRequest
        Assert-ScribbleInstaller -Path $anonymousInstaller
    }
    catch {
        throw (
            "The release was promoted, but Scribble's unauthenticated " +
            "update URL could not download the installer. The internal " +
            "repository may require browser SSO. Do not ship the internal " +
            "updater until this URL works directly. " + $_.Exception.Message)
    }

    $anonymousHash = (Get-FileHash -Algorithm SHA256 `
        -LiteralPath $anonymousInstaller).Hash
    if ($anonymousHash -ne $sourceHash) {
        throw "The anonymous latest-download URL returned a different file."
    }

    Write-Host "Promoted $Tag successfully." -ForegroundColor Green
    Write-Host "Commit: $tagCommit"
    Write-Host "SHA256: $sourceHash"
    Write-Host "Updater URL: $InternalInstallerUrl"
}
finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force `
            -Confirm:$false -WhatIf:$false
    }
}
