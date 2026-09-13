#requires -Version 5.1
param([string]$AssemblyPath=(Join-Path $PSScriptRoot '../../src/Scribble/bin/Release/Scribble.dll'))
$ErrorActionPreference='Stop'
Add-Type -TypeDefinition 'using System;using System.IO;using System.Reflection;public static class ReliabilityResolver{public static void Install(string folder){AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(folder,new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};}}'
[ReliabilityResolver]::Install((Split-Path (Resolve-Path $AssemblyPath).Path))
Add-Type -Path (Resolve-Path $AssemblyPath).Path
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Assert($value,$message){if(-not $value){throw $message}}
function Reject([scriptblock]$action,$message){$rejected=$false;try{& $action}catch{$rejected=$true};Assert $rejected $message}
Assert ($null -eq [Scribble.Testing.TestLab]::Status()) 'Do not run reliability checks during an operator session.'
$cancel=New-Object Threading.CancellationTokenSource
$filter=New-Object Scribble.Testing.TestLabComMessageFilter($cancel.Token)
try {
 Assert ($filter.RetryRejectedCall([IntPtr]::Zero,100,2) -eq 250) 'A temporarily busy Office call was not retried.'
 Assert ($filter.RetryRejectedCall([IntPtr]::Zero,30000,2) -eq -1) 'A busy Office call can retry indefinitely.'
 $cancel.Cancel()
 Assert ($filter.RetryRejectedCall([IntPtr]::Zero,100,1) -eq -1) 'Stop did not cancel rejected-call retries.'
} finally {$filter.Dispose();$cancel.Dispose()}
$folder=Join-Path $PSScriptRoot ('generated/reliability-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $folder -Force | Out-Null
$folder=(Resolve-Path $folder).Path
$state=New-Object Scribble.Testing.SuiteState
$state.id=[guid]::NewGuid().ToString('N');$state.folder=$folder
$zip=[Scribble.Testing.TestLabSuite]::DownloadKit($state,[Threading.CancellationToken]::None)
Assert ([Scribble.Testing.TestLab]::FileHash($zip) -eq $state.kitHash) 'Bundled kit did not verify without a network request.'
$kit=[Scribble.Testing.TestLabSuite]::Extract($zip,(Join-Path $folder 'kit'))
$flags=[Reflection.BindingFlags]'Static,NonPublic'
try {
 [Scribble.Testing.TestLab]::Enable($kit)
 $run=[Scribble.Testing.TestLab]::Start('EX03','Excel',$true)
 # Reproduce the screenshot: the session outlives the recorder process/pipe.
 $session=[Scribble.Testing.TestLab]::Status()
 $session.transport_pipe='scribble-dead-recorder-'+[guid]::NewGuid().ToString('N')
 $session.transport_pid=2147483647;$session.transport_process_start=1
 [Scribble.Testing.TestLab].GetMethod('SaveSession',$flags).Invoke($null,@($session)) | Out-Null
 $recovered=[Scribble.Testing.TestLabSuite]::RecoverIncomplete((Join-Path $folder 'recovered'))
 Assert ([Scribble.Testing.TestLabPdfWriter]::IsValid($recovered)) 'Dead-recorder recovery did not produce a valid PDF.'
 Assert ($null -eq [Scribble.Testing.TestLab]::ActiveRunId()) 'Recovery left the stale run latched.'
 Assert (-not [Scribble.Testing.TestLab]::GetRun($run.run_id).trace_complete) 'Recovered evidence was incorrectly certified complete.'
 [Scribble.Testing.TestLab]::Enable($kit)
 $next=[Scribble.Testing.TestLab]::Start('EX03','Excel',$true)
 Assert ($next.run_id -ne $run.run_id) 'A recovered run was reused.'
 Reject { [Scribble.Testing.TestLabMail]::OpenItem($null,('scribble-fixture:'+$run.run_id+':0'),'') } 'A stale synthetic message escaped its run boundary.'
 [Scribble.Testing.TestLab]::Record($next.run_id,'source','input_verified',@{text='95000'})
 [Scribble.Testing.TestLab]::Record($next.run_id,'pane','pane_event',@{type='user';text='95000'})
 [Scribble.Testing.TestLab]::Record($next.run_id,'pane','pane_event',@{type='assistant';text='The complete June revenue is 42.'})
 [Scribble.Testing.TestLab]::Finish($false)
 $evidence=[Scribble.Testing.TestLab]::Export($next.run_id,(Join-Path $folder 'wrong-answer'))
 $archive=[IO.Compression.ZipFile]::OpenRead($evidence)
 try { Assert (-not @($archive.Entries | Where-Object { $_.FullName.Contains('\') }).Count) 'Evidence ZIP contains host-dependent backslash paths.' } finally {$archive.Dispose()}
 $evaluation=[Scribble.Testing.TestLabEvaluator]::Evaluate($evidence,(Join-Path $folder 'wrong-answer'))
 Assert ($evaluation.status -eq 'failed') 'A correct number in input evidence certified a wrong final answer.'
 $output=[Scribble.Testing.TestLabEvaluator]::OutputText('xlsx',"Worksheet: Sales`nR1C1: 120000`nWorksheet: Scribble Draft`nR1C1: 42")
 Assert ($output.Contains('42') -and -not $output.Contains('120000')) 'Source worksheet values leaked into output grading.'
 $results=@()
 foreach($case in [Scribble.Testing.TestLabSuite]::SelectCases([Scribble.Testing.TestLab]::Cases(),'')) {
  $result=New-Object Scribble.Testing.SuiteCaseResult
  $result.id=$case.id;$result.host=$case.host;$result.status='blocked';$result.error='Synthetic reliability failure, not a live model run.'
  $results+=$result
 }
 $state.folder=Join-Path $folder 'all-case-report';[IO.Directory]::CreateDirectory($state.folder) | Out-Null
 $pdf=[Scribble.Testing.TestLabSuiteReport]::Create($state,$results)
 Assert ([Scribble.Testing.TestLabPdfWriter]::IsValid($pdf)) 'The 16-case report was not readable.'
 Write-Output ('Reliability report: '+$pdf)
 Write-Output 'PASS: offline bundled kit, dead-pipe recovery, new-run restart, stale mail rejection, final-answer grading, output-only values and 16-case PDF.'
} finally {[Scribble.Testing.TestLab]::Disable()}
