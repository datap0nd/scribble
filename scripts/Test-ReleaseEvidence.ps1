param(
    [Parameter(Mandatory=$true)][string]$CandidateDirectory,
    [Parameter(Mandatory=$true)][string]$EvidencePath
)
$ErrorActionPreference = 'Stop'
$candidate = Get-Content -LiteralPath (Join-Path $CandidateDirectory 'candidate.json') -Raw | ConvertFrom-Json
$evidence = Get-Content -LiteralPath $EvidencePath -Raw | ConvertFrom-Json
$hash = (Get-FileHash -LiteralPath (Join-Path $CandidateDirectory 'ScribbleSetup.exe') -Algorithm SHA256).Hash.ToLowerInvariant()
if ($candidate.schema -ne 1 -or $evidence.schema -ne 1 -or $hash -ne $candidate.installer_sha256 -or
    $hash -ne $evidence.installer_sha256 -or $candidate.commit -ne $evidence.commit -or
    $candidate.version -ne $evidence.version -or $candidate.extension_version -ne $evidence.extension_version) {
    throw 'Acceptance evidence does not identify these exact installer and extension bits.'
}
if ($evidence.execution_kind -ne 'native' -or $evidence.unresolved_failures -ne 0 -or $evidence.false_completions -ne 0) {
    throw 'Real native acceptance with zero unresolved failures and false completions is required.'
}
$apps = @('outlook','excel','powerpoint','word','chrome')
$required = @()
foreach ($origin in $apps) {
    foreach ($destination in $apps) {
        if ($origin -eq $destination) { continue }
        foreach ($state in @('stopped','running')) { $required += "$origin/$destination/$state" }
    }
}
$actual = @($evidence.routes | ForEach-Object {
    if ($_.passed -ne $true -or $_.content_verified -ne $true -or $_.destination_visible -ne $true -or
        [string]::IsNullOrWhiteSpace($_.receipt_sha256) -or $_.receipt_sha256 -notmatch '^[a-f0-9]{64}$') {
        throw 'Every route needs verified native content, visible destination and a receipt hash.'
    }
    "$($_.origin)/$($_.destination)/$($_.initial_state)"
})
if ($actual.Count -ne 40 -or (Compare-Object ($required | Sort-Object) ($actual | Sort-Object))) {
    throw 'All forty distinct native origin/destination/startup cases are required.'
}
$scenarios = @('browser_controls','morning_summary','five_slide_launch','outlook_to_powerpoint','recovery')
if (@($evidence.models).Count -eq 0) { throw 'No actual model configuration was tested.' }
foreach ($model in $evidence.models) {
    if ($model.configuration_fingerprint -notmatch '^[a-f0-9]{64}$' -or $model.synthetic_contract_passed -ne $true -or
        $model.vision_passed -ne $true -or $model.streaming_passed -ne $true) {
        throw 'A tested model, vision and streaming contract are required.'
    }
    foreach ($scenario in $scenarios) {
        $runs = @($model.runs | Where-Object { $_.scenario -eq $scenario })
        if ($runs.Count -lt 20 -or @($runs | Where-Object { $_.passed -ne $true -or $_.native -ne $true -or $_.receipt_sha256 -notmatch '^[a-f0-9]{64}$' }).Count -gt 0 -or
            @($runs.run_id | Sort-Object -Unique).Count -ne $runs.Count) {
            throw "Scenario $scenario requires twenty distinct successful native runs for each model."
        }
    }
}
if ($evidence.update_path_verified -ne $true) { throw 'Public update/download/install/restart verification is missing.' }
Write-Output "PASS: Native and model evidence matches candidate $($candidate.version) ($hash)."
