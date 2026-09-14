$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$publisher = Join-Path $root 'scripts/Publish-ScribbleCandidate.ps1'

# No GitHub calls or release writes may occur, even with a passing build or
# an evidence argument. A fake CLI makes accidental publication observable.
$script:githubCalled = $false
function gh { $script:githubCalled = $true; throw 'Unexpected GitHub access.' }
try {
    foreach ($runId in @(33855311737L, 34785528938L)) {
        $rejected = $false
        try { & $publisher -BuildRunId $runId -EvidencePath 'synthetic-evidence.json' }
        catch {
            if ($_.Exception.Message -notlike 'Public updates are frozen at Scribble 2.0.91.*') { throw }
            $rejected = $true
        }
        if (-not $rejected) { throw 'The frozen publisher accepted a promotion.' }
    }
    if ($script:githubCalled) { throw 'The frozen publisher contacted GitHub.' }
} finally { Remove-Item Function:gh }

$workflow = Get-Content (Join-Path $root '.github/workflows/build.yml') -Raw
if ($workflow -match 'contents:\s*write|publish-update:|action-gh-release|Publish-ScribbleCandidate|gh release') {
    throw 'The build workflow can publish public releases.'
}
if ($workflow -notmatch 'contents:\s*read' -or $workflow -notmatch '(?m)^      - codex/development\s*$') {
    throw 'Development builds must run with read-only repository permissions.'
}
Write-Output 'PASS: stable and development promotions refused before GitHub access; CI builds development without release publication.'
