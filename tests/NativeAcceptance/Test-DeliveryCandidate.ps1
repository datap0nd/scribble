param(
    [string]$TestExecutable = (Join-Path $PSScriptRoot '../GuardrailTests/bin/Release/GuardrailTests.exe'),
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '../benchmarks/generated/delivery-candidate'),
    [string]$SourcePresentation,
    [string]$SourceWorkbook,
    [switch]$RunNative,
    [switch]$SkipOffline,
    [int]$TimeoutSeconds = 240
)
$ErrorActionPreference = 'Stop'
$testPath = (Resolve-Path -LiteralPath $TestExecutable).Path
$assemblyPath = Join-Path (Split-Path $testPath) 'Scribble.dll'
$candidateHash = (Get-FileHash -LiteralPath $assemblyPath -Algorithm SHA256).Hash.ToLowerInvariant()
$outputPath = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $outputPath | Out-Null
$report = [ordered]@{
    schema = 1
    assembly_sha256 = $candidateHash
    guardrail_executable_sha256 = (Get-FileHash -LiteralPath $testPath -Algorithm SHA256).Hash.ToLowerInvariant()
    offline = 'pending'
    native_scope = 'pending'
    production_route = 'pending'
    paid_model = 'not_run'
    visual = 'pending'
    generalization = 'pending'
    full_acceptance_passed = $false
}

function Invoke-NativeCheck([string[]]$Arguments, [string]$ResultPath) {
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $testPath
    $start.Arguments = ($Arguments | ForEach-Object {
        if ($_ -match '["\r\n]') { throw 'Invalid native argument.' }
        '"' + $_ + '"'
    }) -join ' '
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    try {
        [void]$process.Start()
        $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
        while (-not $process.WaitForExit(1000)) {
            if ([DateTime]::UtcNow -gt $deadline) {
                $process.Kill()
                throw 'Native test timed out. Only the test runner was stopped; inspect surviving test-owned Office drafts.'
            }
        }
        if (-not (Test-Path -LiteralPath $ResultPath)) { throw 'Native test produced no receipt.' }
        $receipt = Get-Content -LiteralPath $ResultPath -Raw | ConvertFrom-Json
        if ($process.ExitCode -ne 0 -or $receipt.assembly_sha256 -ne $candidateHash) {
            throw "Native test failed or belongs to another assembly: $ResultPath"
        }
        return $receipt
    } finally { $process.Dispose() }
}

try {
    if (-not $SkipOffline) {
        & $testPath *> (Join-Path $outputPath 'guardrails.log')
        if ($LASTEXITCODE -ne 0) { throw 'Offline guardrails failed.' }
        $report.offline = 'passed'
    }
    if ($RunNative) {
        if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
            throw 'PowerPoint is already running. Native acceptance requires a fresh test-owned session; no existing process was touched.'
        }
        if (-not $SourcePresentation -or -not $SourceWorkbook) { throw 'Supply disposable source fixtures for native acceptance.' }
        $SourcePresentation = (Resolve-Path -LiteralPath $SourcePresentation).Path
        $SourceWorkbook = (Resolve-Path -LiteralPath $SourceWorkbook).Path
        $scopePath = Join-Path $outputPath 'chartless-acceptance.json'
        $scope = Invoke-NativeCheck @('--native-powerpoint-chartless', $scopePath) $scopePath
        if ($scope.scope -ne 'chartless-v1' -or -not $scope.chartless_operations_passed -or
            $scope.all_operations_passed -or -not $scope.preservation_passed -or
            -not $scope.rollback_passed -or -not $scope.revision_passed -or $scope.powerpoint_exited) {
            throw 'Chartless native receipt does not establish the supported scope.'
        }
        $report.native_scope = 'passed'
        # Temporarily install only the measured scope for this exact candidate.
        # Restore the user's installed-build receipt even when the route fails.
        $receiptPath = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Scribble/PowerPointAcceptance.json'
        $hadReceipt = Test-Path -LiteralPath $receiptPath
        $oldReceipt = if ($hadReceipt) { [IO.File]::ReadAllBytes($receiptPath) } else { $null }
        try {
            New-Item -ItemType Directory -Force -Path (Split-Path $receiptPath) | Out-Null
            [IO.File]::WriteAllBytes($receiptPath, [IO.File]::ReadAllBytes($scopePath))
            $routePath = Join-Path $outputPath 'production-route.json'
            $route = Invoke-NativeCheck @('--native-pilot-route', $SourcePresentation, $SourceWorkbook, $routePath) $routePath
            if (-not $route.production_route_passed -or -not $route.terminal_receipt_passed -or
                -not $route.source_preserved -or -not $route.workbook_preserved -or
                $route.paid_model_calls -ne 0 -or $route.model_requests -gt 18) { throw 'Production route receipt is incomplete.' }
            $report.production_route = 'passed'
        } finally {
            if ($hadReceipt) { [IO.File]::WriteAllBytes($receiptPath, $oldReceipt) }
            elseif (Test-Path -LiteralPath $receiptPath) { Remove-Item -LiteralPath $receiptPath }
        }
    }
} finally {
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $outputPath 'candidate.json') -Encoding UTF8
}
