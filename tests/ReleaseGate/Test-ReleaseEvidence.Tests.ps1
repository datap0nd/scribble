$ErrorActionPreference = 'Stop'
$temporaryBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$fixtureDirectory = Join-Path $temporaryBase ('scribble-release-gate-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixtureDirectory | Out-Null
try {
    $installer = Join-Path $fixtureDirectory 'ScribbleSetup.exe'
    'Synthetic test bytes, not an installer' | Set-Content -LiteralPath $installer
    $hash = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToLowerInvariant()
    $candidate = @{ schema=1;commit=('a'*40);version='0.0.0.0';extension_version='0.0.0';installer_sha256=$hash }
    $candidate | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $fixtureDirectory 'candidate.json')
    $routes = @()
    foreach ($origin in @('outlook','excel','powerpoint','word','chrome')) {
        foreach ($destination in @('outlook','excel','powerpoint','word','chrome')) {
            if ($origin -eq $destination) { continue }
            foreach ($state in @('stopped','running')) {
                $routes += @{origin=$origin;destination=$destination;initial_state=$state;passed=$true;content_verified=$true;destination_visible=$true;receipt_sha256=$hash}
            }
        }
    }
    $runs = @()
    foreach ($scenario in @('browser_controls','morning_summary','five_slide_launch','outlook_to_powerpoint','recovery')) {
        1..20 | ForEach-Object { $runs += @{scenario=$scenario;run_id="$scenario-$_";passed=$true;native=$true;receipt_sha256=$hash} }
    }
    # This is a validator fixture, not an assertion that native apps ran.
    $evidence = @{schema=1;commit=$candidate.commit;version=$candidate.version;extension_version=$candidate.extension_version;installer_sha256=$hash;
        execution_kind='native';unresolved_failures=0;false_completions=0;routes=$routes;update_path_verified=$true;
        models=@(@{configuration_fingerprint=$hash;synthetic_contract_passed=$true;vision_passed=$true;streaming_passed=$true;runs=$runs})}
    $evidencePath = Join-Path $fixtureDirectory 'evidence.json'
    $validator = Join-Path $PSScriptRoot '..\..\scripts\Test-ReleaseEvidence.ps1'
    $evidence | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $evidencePath
    & $validator -CandidateDirectory $fixtureDirectory -EvidencePath $evidencePath | Out-Null
    $evidence.routes = @($routes | Select-Object -Skip 1)
    $evidence | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $evidencePath
    $rejected = $false
    try { & $validator -CandidateDirectory $fixtureDirectory -EvidencePath $evidencePath | Out-Null } catch { $rejected = $true }
    if (-not $rejected) { throw 'Gate accepted a missing native route.' }
    $evidence.routes = $routes
    $evidence | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $evidencePath
    'Different candidate bytes' | Set-Content -LiteralPath $installer
    $rejected = $false
    try { & $validator -CandidateDirectory $fixtureDirectory -EvidencePath $evidencePath | Out-Null } catch { $rejected = $true }
    if (-not $rejected) { throw 'Gate accepted evidence for different installer bytes.' }
    Write-Output 'PASS: Release gate rejects missing routes and mismatched installer bits (synthetic validator fixtures).'
} finally {
    $resolved = [IO.Path]::GetFullPath($fixtureDirectory)
    if ($resolved -ne $fixtureDirectory -or -not $resolved.StartsWith($temporaryBase, [StringComparison]::OrdinalIgnoreCase) -or
        [IO.Path]::GetFileName($resolved) -notlike 'scribble-release-gate-*') { throw 'Unexpected temporary cleanup target.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
