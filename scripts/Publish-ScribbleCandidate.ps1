param(
    [Parameter(Mandatory=$true)][long]$BuildRunId,
    [string]$EvidencePath
)
$ErrorActionPreference = 'Stop'
$repo = 'datap0nd/scribble'
$run = (& gh api "repos/$repo/actions/runs/$BuildRunId" | ConvertFrom-Json)
if ($LASTEXITCODE -ne 0 -or $run.head_branch -ne 'main' -or $run.path -ne '.github/workflows/build.yml') {
    throw 'Promotion requires a successful main build from the installer workflow.'
}
$jobs = (& gh api "repos/$repo/actions/runs/$BuildRunId/jobs?per_page=100" | ConvertFrom-Json).jobs
if ($LASTEXITCODE -ne 0) { throw 'Could not verify build jobs.' }
foreach ($name in @('build', 'Browser extension fixtures')) {
    $job = @($jobs | Where-Object { $_.name -eq $name })
    if ($job.Count -ne 1 -or $job[0].conclusion -ne 'success') { throw "Required job did not pass: $name" }
}
$main = (& gh api "repos/$repo/commits/main" --jq '.sha')
if ($LASTEXITCODE -ne 0 -or $main -ne $run.head_sha) { throw 'A newer main commit exists; do not publish an older build.' }
$promotionRoot = Join-Path $env:LOCALAPPDATA ('Scribble\Promotions\' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $promotionRoot -Force | Out-Null
& gh run download $BuildRunId --repo $repo --name ScribbleSetup --dir $promotionRoot
if ($LASTEXITCODE -ne 0) { throw 'Could not download the exact candidate artifact.' }
$candidate = Get-Content -LiteralPath (Join-Path $promotionRoot 'candidate.json') -Raw | ConvertFrom-Json
if ($candidate.commit -ne $run.head_sha -or [string]$candidate.build_run_id -ne [string]$BuildRunId) {
    throw 'Candidate identity does not match the successful build.'
}
$installer = Join-Path $promotionRoot 'ScribbleSetup.exe'
$hash = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToLowerInvariant()
$versionInfo = (Get-Item -LiteralPath $installer).VersionInfo
$fileVersion = '{0}.{1}.{2}.{3}' -f $versionInfo.FileMajorPart, $versionInfo.FileMinorPart, $versionInfo.FileBuildPart, $versionInfo.FilePrivatePart
if ($hash -ne $candidate.installer_sha256 -or $fileVersion -ne $candidate.version) {
    throw 'Installer hash or file version does not match the tested candidate.'
}
$acceptance = 'Automated browser, guardrail, and installer checks passed. Live native/model acceptance is not certified by this release.'
if ($EvidencePath) {
    & (Join-Path $PSScriptRoot 'Test-ReleaseEvidence.ps1') -CandidateDirectory $promotionRoot -EvidencePath $EvidencePath
    $acceptance = 'Native and model acceptance passed for these exact bits, in addition to automated checks.'
}
# Keep the previous public installer locally before replacing its download.
$rollback = Join-Path $promotionRoot 'rollback'
New-Item -ItemType Directory -Path $rollback | Out-Null
& gh release download continuous --repo $repo --pattern ScribbleSetup.exe --dir $rollback
if ($LASTEXITCODE -ne 0) { throw 'The previous installer could not be retained for rollback.' }
$notes = "Scribble public update.`nVersion: $($candidate.version)`nCommit: $($candidate.commit)`nInstaller SHA-256: $($candidate.installer_sha256)`nExtension: $($candidate.extension_version)`n$acceptance"
$notesPath = Join-Path $promotionRoot 'release-notes.md'
$notes | Set-Content -LiteralPath $notesPath -Encoding utf8
& gh release upload continuous (Join-Path $promotionRoot 'ScribbleSetup.exe') --repo $repo --clobber
if ($LASTEXITCODE -ne 0) { throw 'Installer promotion failed; the downloaded candidate and rollback remain available.' }
& gh api --method PATCH "repos/$repo/git/refs/tags/continuous" -f "sha=$($candidate.commit)" -F force=true --silent
if ($LASTEXITCODE -ne 0) { throw 'Installer uploaded, but the public release tag could not be aligned with its source commit.' }
& gh release edit continuous --repo $repo --notes-file $notesPath --latest
if ($LASTEXITCODE -ne 0) { throw 'The installer uploaded, but release metadata failed to update. Reconcile the release before retrying.' }
$verified = Join-Path $promotionRoot 'updater-download.exe'
Invoke-WebRequest -Uri "https://github.com/$repo/releases/latest/download/ScribbleSetup.exe" -OutFile $verified
if ((Get-FileHash -LiteralPath $verified -Algorithm SHA256).Hash.ToLowerInvariant() -ne $hash) {
    throw 'The updater URL is not serving the published installer. Reconcile the public release before declaring delivery.'
}
Write-Output "Promoted $($candidate.version). Exact candidate and previous installer retained at $promotionRoot."
