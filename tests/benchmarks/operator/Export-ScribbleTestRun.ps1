#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$RunId,[Parameter(Mandatory=$true)][string]$DestinationDirectory,[string]$AssemblyPath)
. (Join-Path $PSScriptRoot 'TestLab.Common.ps1') -AssemblyPath $AssemblyPath
[Scribble.Testing.TestLab]::Export($RunId, [IO.Path]::GetFullPath($DestinationDirectory))
