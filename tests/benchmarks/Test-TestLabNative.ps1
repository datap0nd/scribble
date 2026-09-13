#requires -Version 5.1
# Explicit local Office integration smoke. No model is called and no mailbox is
# imported/searched. Do not equate these manufactured outputs with Qwen quality.
param([switch]$RunOffice,[switch]$SkipOutlook,[string]$AssemblyPath=(Join-Path $PSScriptRoot '../../src/Scribble/bin/Release/Scribble.dll'))
$ErrorActionPreference='Stop'
if(-not $RunOffice){throw 'Pass -RunOffice to open synthetic documents in the installed desktop apps.'}
$names=if($SkipOutlook){@('EXCEL','POWERPNT','WINWORD')}else{@('EXCEL','POWERPNT','OUTLOOK','WINWORD')}
if(@(Get-Process -Name $names -ErrorAction SilentlyContinue).Count){throw 'Close existing Office sessions before this isolated smoke test.'}
Add-Type -TypeDefinition 'using System;using System.IO;using System.Reflection;public static class NativeResolver{public static void Install(string folder){AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(folder,new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};}}'
[NativeResolver]::Install((Split-Path (Resolve-Path $AssemblyPath).Path))
Add-Type -Path (Resolve-Path $AssemblyPath).Path
function Assert($value,$message){if(-not $value){throw $message}}
Assert ($null -eq [Scribble.Testing.TestLab]::Status()) 'Do not run smoke checks during an operator session.'
$root=Join-Path $PSScriptRoot ('generated/native-smoke-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root -Force | Out-Null
$root=(Resolve-Path $root).Path
$zip=(Resolve-Path (Join-Path $PSScriptRoot 'releases/scribble-test-kit-v1.zip')).Path
$state=New-Object Scribble.Testing.SuiteState
$state.id=[guid]::NewGuid().ToString('N');$state.folder=$root;$state.pid=$PID
$state.processStart=(Get-Process -Id $PID).StartTime.ToUniversalTime().Ticks;$state.expires=[DateTime]::UtcNow.AddHours(1)
$state.commit='native integration smoke - no model';$state.kitHash=[Scribble.Testing.TestLab]::FileHash($zip)
$results=@()
foreach($pair in @(@('EX01','Excel'),@('PP01','PowerPoint'),@('OL02','Outlook'))) {
 if($SkipOutlook -and $pair[1] -eq 'Outlook'){continue}
 $id=$pair[0];$appName=$pair[1];$application=$null;$document=$null;$mail=$null;$environment=$null
 $caseFolder=Join-Path $root ('cases/'+$id)
 $result=New-Object Scribble.Testing.SuiteCaseResult;$result.id=$id;$result.host=$appName;$result.status='blocked'
 try {
  $kit=[Scribble.Testing.TestLabSuite]::Extract($zip,$caseFolder)
  [Scribble.Testing.TestLab]::Enable($kit)
  $case=[Scribble.Testing.TestLab]::Cases() | Where-Object id -eq $id
  $environmentType=[Scribble.Testing.TestLab].Assembly.GetType('Scribble.Testing.TestLabOfficeEnvironment')
  $environment=[Activator]::CreateInstance($environmentType,[Reflection.BindingFlags]'Instance,NonPublic',$null,@([Action[string]]{param($text) Write-Host $text}),$null)
  $application=$environmentType.GetMethod('Connect',[Reflection.BindingFlags]'Instance,NonPublic').Invoke($environment,@($appName))
  Write-Output ('Native '+$appName+' '+$application.Version)
  if($appName -eq 'Excel'){$application.Visible=$true;$document=$application.Workbooks.Open((Join-Path $kit 'inputs/excel/Atlas-input.xlsx'),0,$true)}
  elseif($appName -eq 'PowerPoint'){$application.Visible=-1;$document=$application.Presentations.Open((Join-Path $kit 'inputs/powerpoint/Atlas-start.pptx'),-1,0,-1)}
  else {
   $prepare=[Scribble.Testing.TestLabMail].GetMethod('Prepare',[Reflection.BindingFlags]'Static,NonPublic')
   $prepare.Invoke($null,@($application,$case.PSObject.BaseObject,[Threading.CancellationToken]::None,[Action[string]]{param($text) Write-Host $text})) | Out-Null
  }
  $run=[Scribble.Testing.TestLab]::Start($id,$appName,$true)
  $state.caseId=$id;$state.host=$appName;$state.runId=$run.run_id
  [Scribble.Testing.TestLabSuite]::Save($state)
  [Scribble.Testing.TestLab]::Record($run.run_id,'smoke','manual_native_fixture',@{note='Manufactured native integration output. No model was used.'})
  [Scribble.Testing.BenchmarkArtifactCollector]::Capture($run.run_id,'source') | Write-Output
  if($appName -eq 'Excel') {
   $sheet=$document.Worksheets.Add();$sheet.Name='Scribble Draft'
   $sheet.Cells.Item(1,1)='June revenue';$sheet.Cells.Item(1,2).Formula='=SUM(Sales!E6:E9)'
   $sheet.Cells.Item(2,1)='Budget';$sheet.Cells.Item(2,2).Formula='=SUM(Budget!C2:C5)'
   $sheet.Cells.Item(3,1)='Profit';$sheet.Cells.Item(3,2).Formula='=B1-SUM(Sales!F6:F9)'
   $chart=$sheet.ChartObjects().Add(240,20,400,230).Chart;$chart.SetSourceData($sheet.Range('A1:B2'));$chart.ChartType=51
   Assert ([double]$sheet.Cells.Item(1,2).Value2 -eq 120000) 'Native fixture formula did not calculate June revenue.'
  } elseif($appName -eq 'PowerPoint') {
   for($i=1;$i -le 6;$i++){
    $slide=$document.Slides.Add($document.Slides.Count+1,12)
    $shape=$slide.Shapes.AddTextbox(1,40,60,600,180)
    $shape.TextFrame.TextRange.Text="[Scribble draft] Native smoke slide $i`rRevenue 120000 / Budget 130000 / Delivery 94%"
   }
  } else {
   $messages=[Scribble.Testing.TestLabMail]::Load($application,$case)
   Assert ($messages.Count -eq 3) 'Native MSG fixtures were not loaded as the exact working set.'
   Assert ($messages[1].AttachmentNames.Count -gt 0) 'Native MSG attachments are missing.'
   $mail=$application.CreateItem(0);$mail.To='review@example.test';$mail.Subject='Atlas June review'
   $mail.Body='Native smoke only. Revenue 120000; delivery 94%. This is a manufactured unsent draft, not model output.'
   [Scribble.Testing.TestLab]::RegisterMailOutput($mail);$mail.Display($false)
  }
  [Scribble.Testing.BenchmarkArtifactCollector]::Capture($run.run_id,'final') | Write-Output
  [Scribble.Testing.TestLab]::Finish($false)
  $result.evidence=[Scribble.Testing.TestLab]::Export($run.run_id,$caseFolder)
  $evaluation=[Scribble.Testing.TestLabEvaluator]::Evaluate($result.evidence,$caseFolder)
  $result.evaluation=Join-Path $caseFolder 'evaluation.json';$result.status='needs_review'
  Assert (-not @($evaluation.checks | Where-Object { $_.name -like 'required_final_*' -and -not $_.passed }).Count) 'Native capture omitted a required output.'
  Write-Output ('PASS native '+$appName+' fixture creation, readback, capture and evidence export.')
 } catch {$result.error=$_.Exception.ToString();Write-Output ('FAIL native '+$appName+': '+$result.error)}
 finally {
  if($mail){$mail.Close(1)}
  if($document){if($appName -eq 'PowerPoint'){$document.Saved=-1;$document.Close()}else{$document.Close($false)}}
  if($application -and $appName -eq 'Excel' -and $application.Workbooks.Count -eq 0){$application.Quit()}
  if($application -and $appName -eq 'PowerPoint' -and $application.Presentations.Count -eq 0){$application.Quit()}
  if($environment){$environment.Dispose()}
  [Scribble.Testing.TestLab]::Disable();$results+=$result
 }
}
$state.expires=[DateTime]::UtcNow;[Scribble.Testing.TestLabSuite]::Save($state)
$report=[Scribble.Testing.TestLabSuiteReport]::Create($state,$results)
Write-Output ('Native smoke report: '+$report)
if(@($results | Where-Object { $_.error }).Count){exit 1}
