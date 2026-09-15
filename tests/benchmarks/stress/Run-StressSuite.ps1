param(
    [string]$CorpusRoot = (Join-Path $PSScriptRoot '..\generated\stress-corpus'),
    [string]$HostPath = (Join-Path $env:LOCALAPPDATA 'Programs\Scribble\ScribbleBrowserHost.exe'),
    [string[]]$CaseIds,
    [string]$ResumeFrom,
    [string]$ExpectedOfficeBuild,
    [ValidateSet('x64','x86')][string]$ExpectedOfficePlatform,
    [switch]$OpenReport
)
$ErrorActionPreference = 'Stop'
trap {
    if ($OpenReport) {
        Add-Type -AssemblyName System.Windows.Forms
        [void][Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Scribble stress tests')
    }
    throw
}
$CorpusRoot = (Resolve-Path -LiteralPath $CorpusRoot).Path
$HostPath = (Resolve-Path -LiteralPath $HostPath).Path
$manifestPath = Join-Path $CorpusRoot 'manifest.json'
$hash = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash.ToLowerInvariant()
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
if ($manifest.suite_id -ne 'scribble-stress-v1') { throw 'Choose a verified Scribble stress corpus.' }
$cases = @(Get-Content -LiteralPath (Join-Path $CorpusRoot 'operator\cases.json') -Raw | ConvertFrom-Json)
if ($cases.Count -ne 200) { throw 'The full stress catalog must contain exactly 200 cases.' }
$office = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
if ($ExpectedOfficeBuild -and $office.VersionToReport -ne $ExpectedOfficeBuild) { throw ('Office build mismatch: expected ' + $ExpectedOfficeBuild + ', found ' + $office.VersionToReport) }
if ($ExpectedOfficePlatform -and $office.Platform -ne $ExpectedOfficePlatform) { throw ('Office platform mismatch: expected ' + $ExpectedOfficePlatform + ', found ' + $office.Platform) }
if ($ResumeFrom) {
    if ($CaseIds) { throw 'Choose either explicit CaseIds or ResumeFrom.' }
    $previous = Get-Content -LiteralPath $ResumeFrom -Raw | ConvertFrom-Json
    $previousHash = if ($previous.suite) { $previous.suite.kitHash } else { $previous.kit_sha256 }
    if (-not $previousHash -and $previous.folder) {
        $previousSuite = Get-Content -LiteralPath (Join-Path $previous.folder 'suite.json') -Raw | ConvertFrom-Json
        $previousHash = $previousSuite.suite.kitHash
    }
    if ($previousHash -ne $hash) { throw 'Resume requires the identical corpus manifest; changed inputs need a new run.' }
    # A previous failure remains a failure. Only cases never submitted are resumed.
    $CaseIds = @($previous.cases | Where-Object status -eq 'not_run' | ForEach-Object id)
    if ($CaseIds.Count -eq 0) { Write-Output 'No unsubmitted cases remain. Existing failures were not silently retried.'; exit 0 }
}
if ($CaseIds) {
    $CaseIds = @($CaseIds | ForEach-Object { $_.ToUpperInvariant() })
    if (@($CaseIds | Select-Object -Unique).Count -ne $CaseIds.Count -or @($CaseIds | Where-Object { $_ -notmatch '^[A-Z]{2}[0-9]{2}$' -or $_ -notin $cases.id }).Count) {
        throw 'Every selected case must be a unique catalog ID.'
    }
}
$outputRoot = Join-Path (Split-Path -Parent $CorpusRoot) 'stress-runs'
[void][IO.Directory]::CreateDirectory($outputRoot)
$resultPath = Join-Path $outputRoot ('run-' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0,8) + '.json')
[IO.File]::WriteAllText($resultPath + '.environment.json', ([ordered]@{
    office_build = $office.VersionToReport; office_platform = $office.Platform; office_product = $office.ProductReleaseIds
    expected_office_build = $ExpectedOfficeBuild; expected_office_platform = $ExpectedOfficePlatform
    office_parity_verified = [bool]($ExpectedOfficeBuild -and $ExpectedOfficePlatform)
    host_version = (Get-Item -LiteralPath $HostPath).VersionInfo.FileVersion; corpus_sha256 = $hash
} | ConvertTo-Json), (New-Object Text.UTF8Encoding($false)))
$argumentList = @('--test-lab-run','--result-json',$resultPath,'--kit',$CorpusRoot,'--kit-sha256',$hash)
if ($CaseIds) { $argumentList += @('--cases',($CaseIds -join ',')) }
foreach ($value in $argumentList) { if ($value.Contains('"') -or $value.EndsWith('\')) { throw 'Unsupported quote or trailing separator in an operator path.' } }
$start = New-Object Diagnostics.ProcessStartInfo
$start.FileName = $HostPath
$start.Arguments = ($argumentList | ForEach-Object { '"' + $_ + '"' }) -join ' '
$start.UseShellExecute = $false
$start.CreateNoWindow = $true
$start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
Write-Output ('Starting ' + $(if ($CaseIds) {$CaseIds.Count} else {200}) + ' Qwen Office cases. Progress result: ' + $resultPath)
Write-Output 'The visible Test Lab window provides Stop. API budget verification occurs before each case. No credits are purchased by this launcher.'
$process = [Diagnostics.Process]::Start($start)
$process.WaitForExit()
if (-not (Test-Path -LiteralPath $resultPath)) { throw ('The host did not create an operator result. Install the current stress-suite development build. Exit: ' + $process.ExitCode) }
$result = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
Write-Output ($result | Select-Object status,requested_count,completed_count,harness_complete,correctness_status,pdf,error | ConvertTo-Json -Depth 6)
if ($OpenReport) {
    if ($result.pdf_valid -and $result.pdf -and [IO.Path]::GetExtension($result.pdf) -eq '.pdf' -and (Test-Path -LiteralPath $result.pdf)) {
        $reportStart = New-Object Diagnostics.ProcessStartInfo
        $reportStart.FileName = $result.pdf
        $reportStart.UseShellExecute = $true
        [void][Diagnostics.Process]::Start($reportStart)
    } else {
        Add-Type -AssemblyName System.Windows.Forms
        [void][Windows.Forms.MessageBox]::Show(('The run stopped before a PDF was available. ' + $result.error + [Environment]::NewLine + 'Details: ' + $resultPath), 'Scribble stress tests')
    }
}
exit $process.ExitCode
