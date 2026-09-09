#requires -Version 5.1
param([Parameter(Mandatory=$true)][string]$RunId,[Parameter(Mandatory=$true)][string]$DestinationDirectory,[string]$AssemblyPath)
. (Join-Path $PSScriptRoot 'TestLab.Common.ps1') -AssemblyPath $AssemblyPath
$zip=[Scribble.Testing.TestLab]::Export($RunId, [IO.Path]::GetFullPath($DestinationDirectory))
Write-Output "Evidence ZIP: $zip"
if ('Scribble.Testing.TestLabReport' -as [type]) {
    $pdf=[Scribble.Testing.TestLabReport]::Create($zip)
    Write-Output "PDF report: $pdf"
    Write-Output ('Pasteable summary: '+[IO.Path]::ChangeExtension($zip,$null)+'-summary.txt')
} else { Write-Warning 'This older Scribble installation exports ZIP only. Update Scribble to enable the automatic PDF report.' }
