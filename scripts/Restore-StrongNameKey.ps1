$ErrorActionPreference = "Stop"

$encodedPath = Join-Path $PSScriptRoot "..\build\Scribble.snk.base64"
$keyPath = Join-Path $PSScriptRoot "..\src\Scribble\Properties\Scribble.snk"
$encoded = (Get-Content $encodedPath -Raw).Trim()
$bytes = [Convert]::FromBase64String($encoded)
[IO.File]::WriteAllBytes($keyPath, $bytes)
Write-Host "Restored the strong-name key used for stable COM identity."
