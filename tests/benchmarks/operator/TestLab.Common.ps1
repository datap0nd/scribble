#requires -Version 5.1
param([string]$AssemblyPath)
$ErrorActionPreference = 'Stop'
if (-not $AssemblyPath) {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Scribble\Scribble.dll'),
        (Join-Path $env:LOCALAPPDATA 'Programs\Scribble\Scribble.dll'),
        (Join-Path $PSScriptRoot '..\..\..\src\Scribble\bin\Release\Scribble.dll'))
    $AssemblyPath = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
}
if (-not $AssemblyPath -or -not (Test-Path -LiteralPath $AssemblyPath)) { throw 'Pass -AssemblyPath pointing to the installed Scribble.dll from the Test Lab build.' }
Add-Type -Path (Resolve-Path -LiteralPath $AssemblyPath).Path
