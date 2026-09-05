param(
    [Parameter(Mandatory=$true)][long]$BuildRunId,
    [Parameter(Mandatory=$true)][string]$EvidencePath
)
$ErrorActionPreference = 'Stop'
$repo = 'datap0nd/scribble'
$run = (& gh api "repos/$repo/actions/runs/$BuildRunId" | ConvertFrom-Json)
if ($LASTEXITCODE -ne 0 -or $run.head_branch -ne 'main' -or $run.conclusion -ne 'success' -or $run.path -ne '.github/workflows/build.yml') {
    throw 'Promotion requires a successful main build from the installer workflow.'
}
$promotionRoot = Join-Path $env:LOCALAPPDATA ('Scribble\Promotions\' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $promotionRoot -Force | Out-Null
& gh run download $BuildRunId --repo $repo --name ScribbleSetup --dir $promotionRoot
if ($LASTEXITCODE -ne 0) { throw 'Could not download the exact candidate artifact.' }
$candidate = Get-Content -LiteralPath (Join-Path $promotionRoot 'candidate.json') -Raw | ConvertFrom-Json
if ($candidate.commit -ne $run.head_sha -or [string]$candidate.build_run_id -ne [string]$BuildRunId) {
    throw 'Candidate identity does not match the successful build.'
}
& (Join-Path $PSScriptRoot 'Test-ReleaseEvidence.ps1') -CandidateDirectory $promotionRoot -EvidencePath $EvidencePath
# Keep the previous public installer locally before replacing its download.
$rollback = Join-Path $promotionRoot 'rollback'
New-Item -ItemType Directory -Path $rollback | Out-Null
& gh release download continuous --repo $repo --pattern ScribbleSetup.exe --dir $rollback
if ($LASTEXITCODE -ne 0) { throw 'The previous installer could not be retained for rollback.' }
$notes = "Validated Scribble candidate.`nVersion: $($candidate.version)`nCommit: $($candidate.commit)`nInstaller SHA-256: $($candidate.installer_sha256)`nExtension: $($candidate.extension_version)`nNative and model acceptance passed for these exact bits."
$notesPath = Join-Path $promotionRoot 'release-notes.md'
$notes | Set-Content -LiteralPath $notesPath -Encoding utf8
& gh release upload continuous (Join-Path $promotionRoot 'ScribbleSetup.exe') --repo $repo --clobber
if ($LASTEXITCODE -ne 0) { throw 'Installer promotion failed; the downloaded candidate and rollback remain available.' }
& gh release edit continuous --repo $repo --notes-file $notesPath --latest
if ($LASTEXITCODE -ne 0) { throw 'The installer uploaded, but release metadata failed to update. Reconcile the release before retrying.' }
Write-Output "Promoted $($candidate.version). Exact candidate and previous installer retained at $promotionRoot."
