#requires -Version 5.1
param([string]$AssemblyPath=(Join-Path $PSScriptRoot '../../src/Scribble/bin/Release/Scribble.dll'))
$ErrorActionPreference='Stop'
$resolvedAssembly=(Resolve-Path -LiteralPath $AssemblyPath).Path
Add-Type -TypeDefinition 'using System;using System.IO;using System.Reflection;public static class TestLabAssemblyResolver{public static void Install(string folder){AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(folder,new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};}}'
[TestLabAssemblyResolver]::Install((Split-Path $resolvedAssembly))
foreach($dependency in @('Microsoft.Extensions.Logging.Abstractions.dll','PdfSharp.Shared.dll','PdfSharp.System.dll','PdfSharp-gdi.dll')) {
 [void][Reflection.Assembly]::LoadFrom((Join-Path (Split-Path $resolvedAssembly) $dependency))
}
Add-Type -Path $resolvedAssembly
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class LabWindowVisibility {
 [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr handle);
 public delegate bool EnumProc(IntPtr handle,IntPtr value);
 [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent,EnumProc callback,IntPtr value);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr handle,System.Text.StringBuilder text,int length);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr handle,System.Text.StringBuilder text,int length);
 public static string[] Buttons(IntPtr parent){var values=new System.Collections.Generic.List<string>();EnumChildWindows(parent,(h,v)=>{var c=new System.Text.StringBuilder(64);GetClassName(h,c,c.Capacity);if(c.ToString().StartsWith("WindowsForms10.BUTTON")){var t=new System.Text.StringBuilder(128);GetWindowText(h,t,t.Capacity);values.Add(t.ToString());}return true;},IntPtr.Zero);return values.ToArray();}
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
 # Exercise the installed Start-menu command, which intentionally generates
 # its own launch nonce instead of requiring an Office add-in to provide one.
 $started=[Diagnostics.Process]::Start((Join-Path $bin 'ScribbleBrowserHost.exe'),'--test-lab-suite')
 for($i=0;$i -lt 60;$i++) {
  $child=Get-Process ScribbleBrowserHost -ErrorAction SilentlyContinue | Where-Object { $_.Id -notin $before } | Select-Object -First 1
  if($child) { $child.Refresh(); if($child.MainWindowTitle -like '*idle*' -and $child.MainWindowHandle -ne [IntPtr]::Zero){break} }
  Start-Sleep -Milliseconds 250
 }
 Assert ($null -ne $child) 'Test Lab button did not start an operator window.'
 Assert ([LabWindowVisibility]::IsWindowVisible($child.MainWindowHandle)) 'Test Lab operator window is hidden.'
 Assert ($child.MainWindowTitle -like '*idle*') 'Unfinished capture did not reach the idle/recheck window.'
 $buttons=@([LabWindowVisibility]::Buttons($child.MainWindowHandle) | Sort-Object)
 Assert (($buttons -join '|') -eq 'Start|Stop|View final PDF') ('Unexpected Test Lab actions: '+($buttons -join ', '))
 Assert ([Scribble.Testing.TestLab]::ActiveRunId() -eq $run.run_id) 'Opening the recovery UI altered the active capture.'
 for($click=0;$click -lt 20;$click++){[Scribble.Testing.TestLabSuiteWindow]::Open()}
 $after=@(Get-Process ScribbleBrowserHost -ErrorAction SilentlyContinue | Where-Object { $_.Id -notin $before })
 Assert ($after.Count -eq 1) 'Rapid Test clicks launched more than one window.'
 $rejected=$false;try{[void][Scribble.Testing.TestLabSuite]::RecoverIncomplete($folder)}catch{$rejected=$true}
 Assert $rejected 'Recovery must refuse while another operator/model host is still running.'
 [void]$child.CloseMainWindow(); Assert ($child.WaitForExit(10000)) 'Operator window did not close normally.'
 $report=[Scribble.Testing.TestLabSuite]::RecoverIncomplete($folder)
 Assert (Test-Path $report) 'Recovery report missing.'
 Assert ($report.EndsWith('.pdf') -and [Scribble.Testing.TestLabPdfWriter]::IsValid($report)) 'Recovery did not create a valid PDF.'
 Assert ((Get-Content (Join-Path $folder 'report.html') -Raw).Contains('Recovered after the operator closed')) 'Recovery did not preserve the incomplete status.'
 Assert ($null -eq [Scribble.Testing.TestLab]::Status()) 'Recovery did not release the old capture.'
 Write-Output 'PASS: standalone Start-menu launch, one idle three-action window, rapid-click focus, no launch inference, live-host recovery refusal, incomplete PDF, capture release.'
} finally {
 if($child -and -not $child.HasExited){[void]$child.CloseMainWindow();if(-not $child.WaitForExit(3000)){$child.Kill()}}
 [Scribble.Testing.TestLab]::Disable()
}
