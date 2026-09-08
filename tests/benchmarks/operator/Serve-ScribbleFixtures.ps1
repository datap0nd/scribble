#requires -Version 5.1
param([string]$FixtureRoot=(Split-Path $PSScriptRoot -Parent), [int]$Port=8765, [string]$Python='python')
$ErrorActionPreference='Stop'
& $Python (Join-Path $PSScriptRoot 'serve_fixtures.py') --root (Join-Path $FixtureRoot 'inputs\browser') --port $Port
if ($LASTEXITCODE -ne 0) { throw 'Fixture server stopped with an error.' }
