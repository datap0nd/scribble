#requires -Version 5.1
param([string]$FixtureRoot = (Split-Path $PSScriptRoot -Parent), [string]$AssemblyPath)
. (Join-Path $PSScriptRoot 'TestLab.Common.ps1') -AssemblyPath $AssemblyPath
[Scribble.Testing.TestLab]::Enable((Resolve-Path -LiteralPath $FixtureRoot).Path)
[Scribble.Testing.TestLab]::Serialize([Scribble.Testing.TestLab]::Status())
Write-Output 'Enabled for eight hours. Open the Test Lab button in any Scribble pane. Recording begins only when you start a case.'
