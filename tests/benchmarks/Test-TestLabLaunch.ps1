#requires -Version 5.1
param([string]$AssemblyPath=(Join-Path $PSScriptRoot '../../src/Scribble/bin/Release/Scribble.dll'))
$ErrorActionPreference='Stop'
Add-Type -Path (Resolve-Path -LiteralPath $AssemblyPath).Path
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class LabWindowVisibility {
 [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr handle);
}
"@
function Assert($value,$message){if(-not $value){throw $message}}
Assert ($null -eq [Scribble.Testing.TestLab]::Status()) 'Do not run this test during an operator session.'
$folder=Join-Path $PSScriptRoot ('generated/launch-test-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $folder -Force | Out-Null
$folder=(Resolve-Path $folder).Path
$zip=(Resolve-Path (Join-Path $PSScriptRoot 'releases/scribble-test-kit-v1.zip')).Path
$kit=[Scribble.Testing.TestLabSuite]::Extract($zip,(Join-Path $folder 'kit'))
$bin=Split-Path (Resolve-Path $AssemblyPath).Path
$hostBin=Join-Path $PSScriptRoot '../../src/Scribble.BrowserHost/bin/Release'
Copy-Item -LiteralPath (Join-Path $hostBin 'ScribbleBrowserHost.exe') -Destination $bin -Force
if(Test-Path (Join-Path $hostBin 'ScribbleBrowserHost.exe.config')) { Copy-Item -LiteralPath (Join-Path $hostBin 'ScribbleBrowserHost.exe.config') -Destination $bin -Force }
$child=$null
try {
 [Scribble.Testing.TestLab]::Enable($kit)
 $run=[Scribble.Testing.TestLab]::Start('EX01','Excel',$true)
 [Scribble.Testing.TestLab]::Record($run.run_id,'synthetic','task_started',@{note='Launch/recovery regression only; no Office or model request'})
 $before=@(Get-Process ScribbleBrowserHost -ErrorAction SilentlyContinue | ForEach-Object Id)
 [Scribble.Testing.TestLabSuiteWindow]::Open()
 for($i=0;$i -lt 60;$i++) {
  $child=Get-Process ScribbleBrowserHost -ErrorAction SilentlyContinue | Where-Object { $_.Id -notin $before } | Select-Object -First 1
  if($child) { $child.Refresh(); if($child.MainWindowTitle -like '*finished*' -and $child.MainWindowHandle -ne [IntPtr]::Zero){break} }
  Start-Sleep -Milliseconds 250
 }
 Assert ($null -ne $child) 'Test Lab button did not start an operator window.'
 Assert ([LabWindowVisibility]::IsWindowVisible($child.MainWindowHandle)) 'Test Lab operator window is hidden.'
 Assert ($child.MainWindowTitle -like '*finished*') 'Unfinished capture did not reach the recovery screen.'
 Assert ([Scribble.Testing.TestLab]::ActiveRunId() -eq $run.run_id) 'Opening the recovery UI altered the active capture.'
 $rejected=$false;try{[void][Scribble.Testing.TestLabSuite]::RecoverIncomplete($folder)}catch{$rejected=$true}
 Assert $rejected 'Recovery must refuse while another operator/model host is still running.'
 [void]$child.CloseMainWindow(); Assert ($child.WaitForExit(10000)) 'Operator window did not close normally.'
 $report=[Scribble.Testing.TestLabSuite]::RecoverIncomplete($folder)
 Assert (Test-Path $report) 'Recovery report missing.'
 Assert ((Get-Content $report -Raw).Contains('Recovered after the operator closed')) 'Recovery did not preserve the incomplete status.'
 Assert ($null -eq [Scribble.Testing.TestLab]::Status()) 'Recovery did not release the old capture.'
 Write-Output 'PASS: visible operator window, unfinished capture retained, live-host recovery refused, incomplete evidence preserved, capture released.'
} finally {
 if($child -and -not $child.HasExited){[void]$child.CloseMainWindow();if(-not $child.WaitForExit(3000)){$child.Kill()}}
 [Scribble.Testing.TestLab]::Disable()
}
