$ErrorActionPreference = 'Stop'

# Mechanism reproductions for the September 5 audit, not native Office tests.
# These deliberately reproduce the audited rules instead of loading the add-in.
$auditRepo = (Resolve-Path (Join-Path $PSScriptRoot '../../..')).Path
$progressSource = Get-Content (Join-Path $auditRepo 'src/Scribble/Chat/TaskContextManager.cs') -Raw
$slideSource = Get-Content (Join-Path $auditRepo 'src/Scribble/Office/SamsungPresentationReview.cs') -Raw
if (-not $progressSource.Contains('_stalled = failed || signature == _previousExchange ? _stalled + 1 : 0;') -or
    -not $progressSource.Contains('if (_stalled >= 6)') -or
    -not $slideSource.Contains('!(actualSource ?? "").Contains(evidence)')) {
    throw 'Audited rules changed. Recheck the audit before interpreting this reproduction.'
}

$pages = [System.Collections.Generic.List[object]]::new()
$previousSignature = ''
$stalled = 0
for ($page = 1; $page -le 10; $page++) {
    # Nine hundred nonmatching rows precede the first matching row.
    # The first call creates a cursor; subsequent calls use the same cursor.
    # The production response contains no count of scanned nonmatching rows.
    $signature = if ($page -eq 1) { 'initial-call-empty-incomplete' } else { 'same-cursor-empty-incomplete' }
    $stalled = if ($signature -eq $previousSignature) { $stalled + 1 } else { 0 }
    $previousSignature = $signature
    $pages.Add([ordered]@{ page = $page; rows_scanned = 100 * $page; matching_results = 0; stalled = $stalled; pauses = ($stalled -ge 6) })
    if ($stalled -ge 6) { break }
}
if ($pages.Count -ne 8 -or $pages[-1].rows_scanned -ge 901) {
    throw 'The sparse-cursor mechanism did not reproduce as expected.'
}

$numberPattern = '(?<![A-Za-z0-9])[-+]?(?:\d+(?:[,.]\d+)*|\.\d+)(?:[eE][-+]?\d+)?%?'
if (-not $slideSource.Contains($numberPattern)) { throw 'The audited numeric rule changed.' }
$evidence = 'Progress is 95% complete.'
$numberedContent = '1. Review status 2. Agree next steps 3. Confirm owner'
$dateContent = 'Source date: 2026-09-04'
$allowed = @([regex]::Matches($evidence, $numberPattern) | ForEach-Object { $_.Value.Replace(',', '').TrimStart('+').TrimEnd('%') })
$ordinalFalsePositives = @([regex]::Matches($numberedContent, $numberPattern) | ForEach-Object { $_.Value } | Where-Object { $_ -notin $allowed })
$dateTokens = @([regex]::Matches($dateContent, $numberPattern) | ForEach-Object { $_.Value })
$sourceText = "Progress is 95%`ncomplete."
$whitespaceRejected = -not $sourceText.Contains($evidence)
if (($ordinalFalsePositives -join ',') -ne '1,2,3' -or -not $whitespaceRejected) {
    throw 'The slide validation mechanisms did not reproduce as expected.'
}

[ordered]@{
    verification = 'Source-guarded rule reproduction; no Office, model endpoint, or production task loop executed'
    sparse_cursor = [ordered]@{ first_matching_row = 901; stopped_at_row = $pages[-1].rows_scanned; pages = $pages }
    slide_validation = [ordered]@{ rejected_list_ordinals = $ordinalFalsePositives; date_tokens = $dateTokens; whitespace_only_quote_difference_rejected = $whitespaceRejected }
} | ConvertTo-Json -Depth 8
