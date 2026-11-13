$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$publisher = Join-Path $root 'scripts/Publish-ScribbleCandidate.ps1'

# The retired operator script must refuse without any GitHub call, even
# with a passing build or an evidence argument. A fake CLI makes
# accidental publication observable.
$script:githubCalled = $false
function gh { $script:githubCalled = $true; throw 'Unexpected GitHub access.' }
try {
    foreach ($runId in @(33855311737L, 34785528938L)) {
        $rejected = $false
        try { & $publisher -BuildRunId $runId -EvidencePath 'synthetic-evidence.json' }
        catch {
            if ($_.Exception.Message -notlike 'This legacy publisher is retired.*') { throw }
            $rejected = $true
        }
        if (-not $rejected) { throw 'The retired publisher accepted a promotion.' }
    }
    if ($script:githubCalled) { throw 'The retired publisher contacted GitHub.' }
} finally { Remove-Item Function:gh }

$workflow = Get-Content (Join-Path $root '.github/workflows/build.yml') -Raw
if ($workflow -notmatch '(?m)^permissions:\s*\r?\n\s+contents:\s*read\s*$' -or
    $workflow -notmatch '(?m)^      - codex/development\s*$') {
    throw 'Development builds must run with read-only repository permissions.'
}

# Exactly one job may write, and it is the owner-dispatched continuous
# publisher that runs after every build and test gate.
$job = [regex]::Match(
    $workflow,
    '(?ms)^  publish-continuous:\r?\n(.*?)(?=^  [A-Za-z0-9_-]+:\s*\r?\n|\z)')
if (-not $job.Success) { throw 'The continuous publisher job is missing.' }
$jobText = $job.Groups[1].Value
$outside = $workflow.Remove($job.Index, $job.Length)
if ($outside -match 'contents:\s*write|gh release|action-gh-release|GH_TOKEN') {
    throw 'Only the continuous publisher may write releases.'
}
foreach ($gate in @(
    'needs: build',
    "github.event_name == 'workflow_dispatch'",
    'inputs.publish_continuous',
    "github.ref == 'refs/heads/codex/development'",
    'contents: write'
)) {
    if (-not $jobText.Contains($gate)) {
        throw "The continuous publisher is missing its gate: $gate"
    }
}
if ($workflow -notmatch '(?ms)publish_continuous:.*?type:\s*boolean\s*\r?\n\s*default:\s*false') {
    throw 'Publishing must be off unless the owner ticks it.'
}

# It may only replace the continuous assets. Latest (stable 2.0.91),
# tags and other releases stay untouched.
$releaseCommands = [regex]::Matches($jobText, 'gh release \S+ \S+')
if ($releaseCommands.Count -ne 2) { throw 'Unexpected release commands in the continuous publisher.' }
foreach ($command in $releaseCommands) {
    if ($command.Value -ne 'gh release upload continuous' -and
        $command.Value -ne 'gh release edit continuous') {
        throw "The continuous publisher may not run: $($command.Value)"
    }
}
if (-not $jobText.Contains('--latest=false') -or
    $jobText -match '--latest(?!=false)|stable-2\.0\.91|git push|git tag|gh api') {
    throw 'The continuous publisher must never change Latest, tags or other releases.'
}
if (-not $jobText.Contains('installer_sha256 -ne $hash') -or
    -not $jobText.Contains('commit -ne $env:GITHUB_SHA')) {
    throw 'The continuous publisher must verify the installer against its manifest.'
}

Write-Output 'PASS: the legacy publisher refuses without GitHub access; only the owner-dispatched codex/development publisher can replace continuous, and Latest stays stable.'
