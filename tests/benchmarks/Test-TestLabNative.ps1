#requires -Version 5.1
# Explicit local Office integration smoke. No model is called and no mailbox is
# imported/searched. Do not equate these manufactured outputs with Qwen quality.
param([switch]$RunOffice,[switch]$SkipOutlook,[switch]$ReuseOutlook,[switch]$InjectOfficeExit,[string]$AssemblyPath=(Join-Path $PSScriptRoot '../../src/Scribble/bin/Release/Scribble.dll'))
$ErrorActionPreference='Stop'
if(-not $RunOffice){throw 'Pass -RunOffice to open synthetic documents in the installed desktop apps.'}
$names=if($SkipOutlook){@('EXCEL','POWERPNT','WINWORD')}else{@('EXCEL','POWERPNT','OUTLOOK','WINWORD')}
if($ReuseOutlook){$names=@($names | Where-Object {$_ -ne 'OUTLOOK'})} # Only local MSG inputs and an unsent test draft; never quit Outlook.
if(@(Get-Process -Name $names -ErrorAction SilentlyContinue).Count){throw 'Close existing Office sessions before this isolated smoke test.'}
Add-Type -TypeDefinition 'using System;using System.IO;using System.Reflection;using System.Runtime.InteropServices;public static class NativeResolver{[DllImport("user32.dll")]private static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);public static int Owner(int h){uint p;GetWindowThreadProcessId(new IntPtr(h),out p);return (int)p;}public static void Install(string folder){AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(folder,new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};}}'
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
$nativeStarted=[DateTime]::UtcNow
$environmentType=[Scribble.Testing.TestLab].Assembly.GetType('Scribble.Testing.TestLabOfficeEnvironment')
$environment=[Activator]::CreateInstance($environmentType,[Reflection.BindingFlags]'Instance,NonPublic',$null,@([Action[string]]{param($text) Write-Host $text}),$null)
try {
foreach($pair in @(@('EX01','Excel'),@('PP01','PowerPoint'),@('PP03','PowerPoint'),@('OL02','Outlook'))) {
 if($SkipOutlook -and $pair[1] -eq 'Outlook'){continue}
 $id=$pair[0];$appName=$pair[1];$application=$null;$document=$null;$sourceDocument=$null;$mail=$null
 $caseFolder=Join-Path $root ('cases/'+$id)
 $result=New-Object Scribble.Testing.SuiteCaseResult;$result.id=$id;$result.host=$appName;$result.status='blocked'
 try {
  $kit=[Scribble.Testing.TestLabSuite]::Extract($zip,$caseFolder)
  [Scribble.Testing.TestLab]::Enable($kit)
  $case=[Scribble.Testing.TestLab]::Cases() | Where-Object id -eq $id
  $application=$environmentType.GetMethod('Connect',[Reflection.BindingFlags]'Instance,NonPublic').Invoke($environment,@($appName))
  $officeOwner=if($appName -eq 'PowerPoint'){[Diagnostics.Process]::GetProcessById([NativeResolver]::Owner($application.HWND))}else{$null}
  Write-Output ('Native '+$appName+' '+$application.Version)
  if($appName -eq 'Excel'){$application.Visible=$true;$document=$application.Workbooks.Open((Join-Path $kit 'inputs/excel/Atlas-input.xlsx'),0,$true)}
  elseif($appName -eq 'PowerPoint'){$application.Visible=-1;$pptInput=$case.inputs | Where-Object {$_ -like '*.pptx'} | Select-Object -First 1;$document=$application.Presentations.Open((Join-Path $kit $pptInput),-1,0,-1)}
  else {
   $prepare=[Scribble.Testing.TestLabMail].GetMethod('Prepare',[Reflection.BindingFlags]'Static,NonPublic')
   $prepare.Invoke($null,@($application,$case.PSObject.BaseObject,[Threading.CancellationToken]::None,[Action[string]]{param($text) Write-Host $text})) | Out-Null
  }
  if($appName -eq 'Excel') {
   # Exercise actual VT_ERROR and body month cells before starting evidence.
   # This workbook is created and discarded solely by this integration test.
   $probe=$application.Workbooks.Add()
   try {
    $probeRows=New-Object 'Collections.Generic.List[Collections.Generic.IReadOnlyList[string]]'
    foreach($row in @(@('Month','Value'),@('2026-06','=1/0'))) {$probeRows.Add([string[]]$row)}
    $writer=[Scribble.Testing.TestLab].Assembly.GetType('Scribble.Office.WorkbookDraftWriter')
    $write=$writer.GetMethods([Reflection.BindingFlags]'Static,NonPublic') | Where-Object {$_.Name -eq 'WriteDraftSheet' -and $_.GetParameters().Count -eq 3}
    $probeStatus=$write.Invoke($null,@($application.PSObject.BaseObject,'Native error regression',$probeRows.PSObject.BaseObject))
    $probeSheet=$application.ActiveSheet
    Assert ($probeSheet.Cells.Item(4,1).Value2 -ceq '2026-06') 'A body month key became a date serial.'
    Assert ($probeSheet.Cells.Item(4,2).Formula -ceq '=1/0') 'The writer hid the broken formula as text.'
    Assert ([Scribble.Office.ExcelErrorValue]::Text($probeSheet.Cells.Item(4,2).Value2) -eq '#DIV/0!') 'The actual Excel error value was not recognized.'
    Assert ($probeStatus.Contains('remain visible')) 'The writer did not disclose its broken formula.'
   } finally { $probe.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($probe); $document.Activate() }
  }
  $run=[Scribble.Testing.TestLab]::Start($id,$appName,$true)
  $state.caseId=$id;$state.host=$appName;$state.runId=$run.run_id
  [Scribble.Testing.TestLabSuite]::Save($state)
  [Scribble.Testing.TestLab]::Record($run.run_id,'smoke','manual_native_fixture',@{note='Manufactured native integration output. No model was used.'})
  [Scribble.Testing.BenchmarkArtifactCollector]::Capture($run.run_id,'source') | Write-Output
  if($id -eq 'PP03'){[Scribble.Testing.BenchmarkArtifactCollector]::Capture($run.run_id,'intermediate') | Write-Output}
  if($appName -eq 'Excel') {
   $rows=New-Object 'Collections.Generic.List[Collections.Generic.IReadOnlyList[string]]'
   foreach($row in @(@('Metric','2026-05','2026-06','June Budget','Gap'),@('Revenue','100000','=SUM(Sales!E6:E9)','=SUM(Budget!C2:C5)','=C4-D4'),@('Cost','60000','=SUM(Sales!F6:F9)','',''),@('Profit','=B4-B5','=C4-C5','',''),@('Margin','=B6/B4','=C6/C4','',''),@('Growth','','=(C4-B4)/B4','',''),@('Budget gap percent','','=E4/D4','',''))) {$rows.Add([string[]]$row)}
   $writer=[Scribble.Testing.TestLab].Assembly.GetType('Scribble.Office.WorkbookDraftWriter')
   $write=$writer.GetMethods([Reflection.BindingFlags]'Static,NonPublic') | Where-Object {$_.Name -eq 'WriteDraftSheet' -and $_.GetParameters().Count -eq 3}
   [void]$write.Invoke($null,@($application.PSObject.BaseObject,'Native writer smoke - manufactured output',$rows.PSObject.BaseObject))
   $sheet=$application.ActiveSheet
   Assert ($sheet.Cells.Item(3,2).Value2 -ceq '2026-05' -and $sheet.Cells.Item(3,3).Value2 -ceq '2026-06') 'Native writer converted month labels into date serials.'
   Assert ([double]$sheet.Cells.Item(4,3).Value2 -eq 120000 -and [double]$sheet.Cells.Item(6,3).Value2 -eq 46000 -and [double]$sheet.Cells.Item(4,5).Value2 -eq -10000) 'The real draft writer changed formula addresses or numeric values.'
   $chart=$sheet.ChartObjects().Add(520,20,400,230).Chart;$chart.SetSourceData($sheet.Range('A3:C6'));$chart.ChartType=51
  } elseif($appName -eq 'PowerPoint') {
   if($id -eq 'PP03') {$sourceDocument=$document;$document=$application.Presentations.Add();[Scribble.Testing.TestLab]::RegisterOutput($document,'pptx')}
   for($i=1;$i -le 6;$i++){
    $slide=$document.Slides.Add($document.Slides.Count+1,12)
    $shape=$slide.Shapes.AddTextbox(1,40,60,600,180)
    $shape.TextFrame.TextRange.Text="[Scribble draft] Native smoke slide $i`rRevenue 120000; budget 130000; gap -10000 (-7.69%); cost 74000; profit 46000; margin 38.33%; growth 20%; delivery 94% against 97%."
   }
   [void][Scribble.Testing.TestLabSuite]::SaveSelectedDeckForMail($application,$run.run_id)
  } else {
   $messages=[Scribble.Testing.TestLabMail]::Load($application,$case)
   Assert ($messages.Count -eq 3) 'Native MSG fixtures were not loaded as the exact working set.'
   Assert ($messages[1].AttachmentNames.Count -gt 0) 'Native MSG attachments are missing.'
   $mail=$application.CreateItem(0);$mail.To='review@example.test';$mail.Subject='Atlas June review'
   $mail.Body="Native smoke only. Revenue: EUR 120000; budget gap: -10000 (-7.69%); gross profit: EUR 46000; cost: EUR 74000; weighted gross margin: 38.33%; delivery: 94% against target 97%.`r`nPlanned actions: Mira Cole reviews freight costs by 10 July 2026. Leon Park confirms recovery by 12 July 2026. This is a manufactured unsent draft, not model output."
   [Scribble.Testing.TestLab]::RegisterMailOutput($mail);$mail.Display($false)
  }
  [Scribble.Testing.BenchmarkArtifactCollector]::Capture($run.run_id,'final') | Write-Output
  [Scribble.Testing.TestLab]::Finish($false)
  $result.evidence=[Scribble.Testing.TestLab]::Export($run.run_id,$caseFolder)
  if($id -eq 'PP03') {
   $evidenceZip=[IO.Compression.ZipFile]::OpenRead($result.evidence)
   try {
    Assert (@($evidenceZip.Entries | Where-Object {$_.Name -like '*-intermediate-output-*'}).Count -eq 0) 'Untouched starter slides were attributed to this run as output.'
    Assert (@($evidenceZip.Entries | Where-Object {$_.Name -like '*-final-output-*-readback.json'}).Count -eq 1) 'Final evidence included the untouched crowded starter as a new draft.'
   } finally {$evidenceZip.Dispose()}
  }
  $evaluation=[Scribble.Testing.TestLabEvaluator]::Evaluate($result.evidence,$caseFolder)
  $result.evaluation=Join-Path $caseFolder 'evaluation.json';$result.status='needs_review'
  Assert (-not @($evaluation.checks | Where-Object { $_.hard -and -not $_.passed }).Count) ('Native evidence failed: '+($evaluation.findings -join '; '))
  Write-Output ('PASS native '+$appName+' fixture creation, readback, capture and evidence export.')
 } catch {$result.error=$_.Exception.ToString();Write-Output ('FAIL native '+$appName+': '+$result.error)}
 finally {
  try {
  if($mail){$mail.Close(1)}
  if($document){if($appName -eq 'PowerPoint'){$document.Saved=-1;$document.Close()}else{$document.Close($false)}}
  if($sourceDocument){$sourceDocument.Saved=-1;$sourceDocument.Close()}
  if($application -and $appName -eq 'Excel' -and $application.Workbooks.Count -eq 0){$application.Quit()}
  if($application -and $appName -eq 'PowerPoint' -and $application.Presentations.Count -eq 0){
   if($InjectOfficeExit -and $id -eq 'PP01') {
    # Explicit fault injection only into the empty PowerPoint process created
    # by this guarded smoke. Never terminate a pre-existing Office process.
    Assert ($officeOwner.StartTime.ToUniversalTime() -ge $nativeStarted -and $officeOwner.Id -eq [NativeResolver]::Owner($application.HWND)) 'Fault injection could not establish ownership.'
    $officeOwner.Kill();Assert ($officeOwner.WaitForExit(5000)) 'Owned PowerPoint fault injection did not end the executor.'
    Write-Output 'Injected exit of the empty, smoke-owned PowerPoint process; the next case must reconnect.'
   } else {$application.Quit()}
  }
  } catch {$result.error+="`nNative cleanup: "+$_.Exception.ToString()}
  [Scribble.Testing.TestLab]::Disable();$results+=$result
 }
}
} finally {$environment.Dispose()}
$state.expires=[DateTime]::UtcNow;[Scribble.Testing.TestLabSuite]::Save($state)
$report=[Scribble.Testing.TestLabSuiteReport]::Create($state,$results)
Write-Output ('Native smoke report: '+$report)
if(@($results | Where-Object { $_.error }).Count){exit 1}
