#requires -Version 5.1
$ErrorActionPreference = 'Stop'
$testRoot = Join-Path $env:TEMP ('scribble-loader-' + [guid]::NewGuid().ToString('N'))
$current = Join-Path $testRoot 'Programs\Scribble\Scribble.dll'
$legacy = Join-Path $testRoot 'Scribble\Scribble.dll'
New-Item -ItemType Directory -Path (Split-Path $current), (Split-Path $legacy) -Force | Out-Null
Add-Type -TypeDefinition 'namespace Scribble.Testing { public static class TestLab {} }' -OutputAssembly $current
Add-Type -TypeDefinition 'namespace OldScribble { public class NoTestLab {} }' -OutputAssembly $legacy
$runner = Join-Path $testRoot 'probe.ps1'
@'
param($Loader, $LocalRoot, $SelectedAssembly)
$ErrorActionPreference = 'Stop'
$env:LOCALAPPDATA = $LocalRoot
try {
    . $Loader -AssemblyPath $SelectedAssembly
    Write-Output ('TYPE_LOCATION=' + ('Scribble.Testing.TestLab' -as [type]).Assembly.Location)
} catch { Write-Output $_.Exception.Message; exit 1 }
'@ | Set-Content -LiteralPath $runner -Encoding UTF8
$loader = Join-Path $PSScriptRoot 'operator\TestLab.Common.ps1'
$shell = Join-Path $PSHOME 'powershell.exe'
$output = & $shell -NoProfile -File $runner $loader $testRoot 2>&1
if ($LASTEXITCODE -ne 0 -or -not ($output -contains "TYPE_LOCATION=$current")) {
    throw "Current installation was not preferred over legacy DLL: $output"
}
$output = & $shell -NoProfile -File $runner $loader $testRoot $legacy 2>&1
if ($LASTEXITCODE -eq 0 -or ($output -join "`n") -notmatch 'does not include Test Lab' -or ($output -join "`n") -notmatch [regex]::Escape($legacy)) {
    throw "Old DLL did not produce actionable diagnostics: $output"
}
Write-Output 'PASS: current install beats legacy DLL; explicit old DLL fails with path and upgrade guidance.'
