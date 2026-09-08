#requires -Version 5.1
param([string]$AssemblyPath)
. (Join-Path $PSScriptRoot 'TestLab.Common.ps1') -AssemblyPath $AssemblyPath
[Scribble.Testing.TestLab]::Disable()
Write-Output 'Test Lab disabled. Existing evidence remains available for export.'
