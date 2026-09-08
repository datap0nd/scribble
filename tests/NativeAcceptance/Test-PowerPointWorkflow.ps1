param(
    [string]$TestExecutable = (Join-Path $PSScriptRoot '../GuardrailTests/bin/Release/GuardrailTests.exe'),
    [string]$OutputPath = (Join-Path $PSScriptRoot 'PowerPointWorkflow.json'),
    [int]$TimeoutSeconds = 180,
    [switch]$EnableRevision
)
$ErrorActionPreference = 'Stop'
$testPath = [IO.Path]::GetFullPath($TestExecutable)
$reportPath = [IO.Path]::GetFullPath($OutputPath)
if (-not (Test-Path -LiteralPath $testPath)) { throw 'Build the guardrail test executable first.' }
$start = New-Object Diagnostics.ProcessStartInfo
$start.FileName = $testPath
$start.Arguments = '--native-powerpoint "' + $reportPath + '"'
$start.UseShellExecute = $false
$start.CreateNoWindow = $true
$process = New-Object Diagnostics.Process
$process.StartInfo = $start
try {
    [void]$process.Start()
    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
        $process.Kill()
        throw 'Native PowerPoint acceptance timed out. No acceptance receipt was installed. Inspect any surviving unsaved test decks; the PowerPoint process was not terminated.'
    }
    if ($process.ExitCode -ne 0) { throw "Native PowerPoint acceptance failed. See $reportPath. Revision remains disabled." }
    $report = Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json
    if ($report.execution_kind -ne 'native' -or $report.all_operations_passed -ne $true -or
        $report.preservation_passed -ne $true -or $report.rollback_passed -ne $true -or $report.revision_passed -ne $true) {
        throw 'The report does not establish native revision readiness.'
    }
    if ($EnableRevision) {
        $destination = Join-Path $env:LOCALAPPDATA 'Scribble/PowerPointAcceptance.json'
        New-Item -ItemType Directory -Path (Split-Path $destination) -Force | Out-Null
        Copy-Item -LiteralPath $reportPath -Destination $destination -Force
    }
    Write-Output 'PASS: Native revision operations and recovery. Model/UI benchmarks and real Samsung fidelity remain separate acceptance checks.'
} finally { $process.Dispose() }
