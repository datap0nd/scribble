param(
    [string]$HostPath = (Join-Path $PSScriptRoot '..\..\src\Scribble.BrowserHost\bin\Release\ScribbleBrowserHost.exe'),
    [string]$OutputRoot = (Join-Path $PSScriptRoot 'generated')
)
$ErrorActionPreference = 'Stop'
$HostPath = (Resolve-Path -LiteralPath $HostPath).Path
$payload = Split-Path -Parent $HostPath
foreach ($required in @('ScribbleBrowserHost.exe.config', 'Scribble.dll', 'System.ValueTuple.dll')) {
    if (-not (Test-Path -LiteralPath (Join-Path $payload $required) -PathType Leaf)) {
        throw "Standalone Test Lab payload is missing $required."
    }
}
$folder = [IO.Path]::GetFullPath((Join-Path $OutputRoot ('standalone-pdf-' + [guid]::NewGuid().ToString('N'))))
[IO.Directory]::CreateDirectory($folder) | Out-Null
$cases = @('EX01','EX02','EX03','EX04','PP01','PP02','PP03','OL01','XA01','XA02','XA04','RB01','RC01','EX05','PP04','OL02') | ForEach-Object {
    [ordered]@{ id = $_; host = 'Runtime fixture'; status = 'blocked'; failureKind = 'environment';
        error = 'Runtime packaging check only. No Office or model request was submitted.' }
}
$manifest = [ordered]@{ schema = 1; suite = [ordered]@{
    id = [guid]::NewGuid().ToString('N'); folder = $folder; commit = 'standalone runtime packaging test'; kitHash = 'not a model run'
}; cases = @($cases) }
$manifestPath = Join-Path $folder 'suite.json'
[IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($false)))
$start = New-Object Diagnostics.ProcessStartInfo
$start.FileName = $HostPath
$start.Arguments = '--test-lab-report "' + $manifestPath + '"'
$start.WorkingDirectory = $folder
$start.UseShellExecute = $false
$start.CreateNoWindow = $true
$start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
$start.RedirectStandardOutput = $true
$start.RedirectStandardError = $true
$child = [Diagnostics.Process]::Start($start)
try {
    $stdout = $child.StandardOutput.ReadToEndAsync()
    $stderr = $child.StandardError.ReadToEndAsync()
    if (-not $child.WaitForExit(30000)) {
        $child.Kill()
        [void]$child.WaitForExit(5000)
        throw 'The standalone report-only child did not finish within 30 seconds.'
    }
    $output = $stdout.GetAwaiter().GetResult()
    $errorOutput = $stderr.GetAwaiter().GetResult()
    [IO.File]::WriteAllText((Join-Path $folder 'stdout.txt'), $output)
    [IO.File]::WriteAllText((Join-Path $folder 'stderr.txt'), $errorOutput)
    if ($child.ExitCode -ne 0) { throw "Standalone PDF rendering failed (exit $($child.ExitCode)): $errorOutput" }
    $report = $output | ConvertFrom-Json
    $expectedPdf = Join-Path $folder 'report.pdf'
    if ($report.execution_kind -ne 'report_only' -or $report.pdf_valid -ne $true -or $report.pdf -ne $expectedPdf) {
        throw "The standalone child did not confirm its own PDF: $output"
    }
    $bytes = [IO.File]::ReadAllBytes($expectedPdf)
    if ($bytes.Length -lt 5000 -or [Text.Encoding]::ASCII.GetString($bytes, 0, 5) -ne '%PDF-') {
        throw 'The standalone PDF is missing or invalid.'
    }
    Write-Host "PASS: standalone installed-runtime PDF rendered and reopened in its own child process: $expectedPdf"
} finally { $child.Dispose() }
