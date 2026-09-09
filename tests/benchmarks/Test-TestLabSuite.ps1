#requires -Version 5.1
param([string]$AssemblyPath=(Join-Path $PSScriptRoot '../../src/Scribble/bin/Release/Scribble.dll'),[switch]$RenderPdf)
$ErrorActionPreference='Stop'
Add-Type -Path (Resolve-Path -LiteralPath $AssemblyPath).Path
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Assert($value,$message) { if (-not $value) { throw $message } }
function Reject([scriptblock]$action,$message) { $rejected=$false; try { & $action } catch { $rejected=$true }; Assert $rejected $message }
$labRoot=[Scribble.Testing.TestLab]::Root
Assert (-not (Test-Path -LiteralPath (Join-Path $labRoot 'session.bin'))) 'Do not run suite checks during an operator session.'
Assert ($null -eq [Scribble.Testing.TestLabSuite]::Active()) 'Do not run checks during an active suite.'
$folder=Join-Path $PSScriptRoot ('generated/suite-test-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $folder -Force | Out-Null
$folder=(Resolve-Path -LiteralPath $folder).Path
$zip=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot 'releases/scribble-test-kit-v1.zip')).Path
$state=New-Object Scribble.Testing.SuiteState
$state.id=[guid]::NewGuid().ToString('N');$state.folder=$folder;$state.pid=$PID
$state.processStart=(Get-Process -Id $PID).StartTime.ToUniversalTime().Ticks
$state.expires=[DateTime]::UtcNow.AddMinutes(10);$state.commit='synthetic validation';$state.kitHash=[Scribble.Testing.TestLab]::FileHash($zip)
try {
    $kit=[Scribble.Testing.TestLabSuite]::Extract($zip,(Join-Path $folder 'cases/EX01'))
    [Scribble.Testing.TestLab]::Enable($kit)
    $cases=@([Scribble.Testing.TestLab]::Cases()); Assert ($cases.Count -eq 16) 'Suite catalog omitted cases.'
    foreach ($case in $cases) {
        $first=[Scribble.Testing.TestLabSuite]::Prompt($case,0)
        Assert ($first -eq $(if($case.prerequisite_prompt){$case.prerequisite_prompt}else{$case.prompt})) ('Wrong first prompt: '+$case.id)
        if($case.prerequisite_prompt) { Assert ([Scribble.Testing.TestLabSuite]::Prompt($case,1) -eq $case.prompt) ('Wrong follow-up: '+$case.id) }
        else { Reject { [Scribble.Testing.TestLabSuite]::Prompt($case,1) } 'Unexpected second phase accepted.' }
        Reject { [Scribble.Testing.TestLabSuite]::Prompt($case,2) } 'Arbitrary phase accepted.'
    }
    Assert ([Scribble.Testing.TestLabSuite]::PresetAnswer($cases[0],'Which currency?') -eq 'EUR excluding tax') 'Kit preset answer not used.'
    Assert ($null -eq [Scribble.Testing.TestLabSuite]::PresetAnswer($cases[0],'Should I contact someone?')) 'Invented a response outside the kit presets.'
    $run=[Scribble.Testing.TestLab]::Start('EX01','Excel',$true)
    $state.caseId='EX01';$state.host='Excel';$state.runId=$run.run_id
    [Scribble.Testing.TestLabSuite]::Save($state)
    Assert ([Scribble.Testing.TestLabSuite]::Require($state.id,'Excel').runId -eq $run.run_id) 'Valid pane command rejected.'
    Reject { [Scribble.Testing.TestLabSuite]::Require('wrong','Excel') } 'Wrong suite accepted.'
    Reject { [Scribble.Testing.TestLabSuite]::Require($state.id,'Word') } 'Wrong host accepted.'
    $fixtureInput=Join-Path $kit $run.input_paths[0]
    Assert ([Scribble.Testing.TestLabSuite]::OwnsSource($run.run_id,$fixtureInput)) 'Suite-owned source not collectable.'
    Assert (-not [Scribble.Testing.TestLabSuite]::OwnsSource($run.run_id,(Join-Path $folder 'unrelated.xlsx'))) 'Unrelated source accepted.'
    # Exercise the production pane driver without Office/model dependencies.
    $driverType=[Scribble.Testing.TestLab].Assembly.GetType('Scribble.Testing.TestLabSuitePane')
    $driver=[Activator]::CreateInstance($driverType,$true);$method=$driverType.GetMethod('Command')
    $script:sendCount=0;$script:stopCount=0
    $pending=New-Object 'Threading.Tasks.TaskCompletionSource[bool]'
    $busy=[Func[bool]]{return $false};$reset=[Action]{};$load=[Action[Scribble.Testing.LabCase]]{}
    $send=[Func[string,Threading.Tasks.Task]]{param($prompt);$script:sendCount++;return $pending.Task}
    $stop=[Action]{$script:stopCount++}
    $commandArgs=@($state.id,'load-once','load',0,'Excel',$true,$busy,$reset,$load,$send,$stop)
    $loaded=$method.Invoke($driver,$commandArgs) | ConvertFrom-Json
    Assert ($loaded.state -eq 'done') 'Pane did not load.'
    $commandArgs[1]='submit-once';$commandArgs[2]='submit'
    $started=$method.Invoke($driver,$commandArgs) | ConvertFrom-Json
    Assert ($started.state -eq 'running' -and $script:sendCount -eq 1) 'Submission was not tracked.'
    [void]$method.Invoke($driver,$commandArgs)
    Assert ($script:sendCount -eq 1) 'Duplicate command submitted twice.'
    $commandArgs[2]='stop'
    $stopping=$method.Invoke($driver,$commandArgs) | ConvertFrom-Json
    Assert ($stopping.state -eq 'running' -and $script:stopCount -eq 1) 'Stop falsely claimed the async operation finished.'
    $pending.SetResult($true);$commandArgs[2]='status'
    for($i=0;$i -lt 100;$i++) { $done=$method.Invoke($driver,$commandArgs) | ConvertFrom-Json;if($done.state -eq 'done'){break};Start-Sleep -Milliseconds 10 }
    Assert ($done.state -eq 'done') 'Completed task remained running.'
    [Scribble.Testing.TestLab]::Finish($false)
    [Scribble.Testing.TestLab]::Disable()
    $pptKit=[Scribble.Testing.TestLabSuite]::Extract($zip,(Join-Path $folder 'cases/PP01'))
    [Scribble.Testing.TestLab]::Enable($pptKit)
    $pptRun=[Scribble.Testing.TestLab]::Start('PP01','PowerPoint',$true)
    $state.caseId='PP01';$state.host='PowerPoint';$state.runId=$pptRun.run_id
    [Scribble.Testing.TestLabSuite]::Save($state)
    $csvInput=Join-Path $pptKit 'inputs/data/sales.csv'
    Assert ([Scribble.Testing.TestLabSuite]::OwnsSource($pptRun.run_id,$csvInput)) 'CSV must be an allowed supporting input for this test.'
    Assert (-not [Scribble.Testing.TestLabSuite]::OwnsNativeSource($pptRun.run_id,$csvInput,'Excel')) 'Supporting CSV would be mislabeled as XLSX.'
    [Scribble.Testing.TestLab]::Finish($false);[Scribble.Testing.TestLab]::Disable()
    $chromeKit=[Scribble.Testing.TestLabSuite]::Extract($zip,(Join-Path $folder 'cases/CH01'))
    [Scribble.Testing.TestLab]::Enable($chromeKit)
    $chromeRun=[Scribble.Testing.TestLab]::Start('CH01','Chrome',$true)
    $state.caseId='CH01';$state.host='Chrome';$state.runId=$chromeRun.run_id;$state.chromeToken=[guid]::NewGuid().ToString('N')
    [Scribble.Testing.TestLabSuite]::Save($state)
    $commandId=[guid]::NewGuid().ToString('N')
    @{id=$commandId;action='submit';prompt='Synthetic controller validation';runId=$chromeRun.run_id} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $folder 'cases/CH01/chrome-command.json')
    $request=@{suite=$state.id;token=$state.chromeToken;controller=[guid]::NewGuid().ToString('N')}
    $command=[Scribble.Testing.TestLabSuite]::Chrome(($request | ConvertTo-Json)) | ConvertFrom-Json
    Assert ($command.id -eq $commandId) 'Chrome controller did not receive command.'
    $other=$request.Clone();$other.controller=[guid]::NewGuid().ToString('N')
    Assert ([Scribble.Testing.TestLabSuite]::Chrome(($other | ConvertTo-Json)) -eq 'null') 'Second Chrome controller claimed the same case.'
    $other=$request.Clone();$other.token='old-case'
    Reject { [Scribble.Testing.TestLabSuite]::Chrome(($other | ConvertTo-Json)) } 'Old Chrome case token accepted.'
    $request.action='complete';$request.id=$commandId;$request.error='Synthetic connection error'
    [void][Scribble.Testing.TestLabSuite]::Chrome(($request | ConvertTo-Json))
    $reply=Get-Content -LiteralPath (Join-Path $folder ('cases/CH01/'+$commandId+'.reply.json')) -Raw | ConvertFrom-Json
    Assert ($reply.error -eq 'Synthetic connection error') 'Browser error not persisted.'
    [Scribble.Testing.TestLab]::Finish($false);[Scribble.Testing.TestLab]::Disable()
    $state.expires=[DateTime]::UtcNow.AddSeconds(-1);[Scribble.Testing.TestLabSuite]::Save($state)
    Reject { [Scribble.Testing.TestLabSuite]::Require($state.id,'Chrome') } 'Expired suite accepted.'

    # Reproduce a failure before the preparation script can write JSON. This never launches Office.
    $bootstrap=Join-Path $kit 'operator/Prepare-ScribbleTestCase.ps1'
    Set-Content -LiteralPath $bootstrap -Value 'throw "BOOTSTRAP_DIAGNOSTIC_SENTINEL"' -Encoding UTF8
    $manifestPath=Join-Path $kit 'manifest.json';$manifest=Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    $entry=$manifest.files | Where-Object path -eq 'operator/Prepare-ScribbleTestCase.ps1'
    $entry.sha256=[Scribble.Testing.TestLab]::FileHash($bootstrap);$entry.size=(Get-Item -LiteralPath $bootstrap).Length
    $manifest | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    [Scribble.Testing.TestLab]::Enable($kit)
    $preparation=[Scribble.Testing.TestLabPreparation]::Launch('EX01',$true)
    $report=$null
    for($i=0;$i -lt 100;$i++) { $report=[Scribble.Testing.TestLabPreparation]::ReadReport($preparation);if($report -and $report.status -eq 'failed'){break};Start-Sleep -Milliseconds 100 }
    Assert ($report.status -eq 'failed' -and $report.log.Contains('BOOTSTRAP_DIAGNOSTIC_SENTINEL')) 'Startup stderr was replaced with generic exit 1.'
    Assert (([Scribble.Testing.TestLabPreparation]::ReadReport($preparation)).status -eq 'failed') 'Failure status was not durable.'
    [Scribble.Testing.TestLab]::Disable()
    $badZip=Join-Path $folder 'traversal.zip'
    $archive=[IO.Compression.ZipFile]::Open($badZip,[IO.Compression.ZipArchiveMode]::Create)
    try { [void]$archive.CreateEntry('scribble-test-kit-v1/../../escaped.txt') } finally { $archive.Dispose() }
    Reject { [Scribble.Testing.TestLabSuite]::Extract($badZip,(Join-Path $folder 'rejected')) } 'Traversal archive accepted.'
    Assert (-not (Test-Path -LiteralPath (Join-Path $folder 'escaped.txt'))) 'Traversal wrote outside extraction root.'
    $result=New-Object Scribble.Testing.SuiteCaseResult
    $result.id='EX01';$result.host='Excel';$result.status='blocked';$result.started='2026-09-09T12:00:00Z';$result.finished='2026-09-09T12:00:01Z'
    $result.error='<script>alert(1)</script> synthetic preparation error '+('diagnostic line ' * 200)+' END_OF_ERROR'
    '2026-09-09T12:00:00Z Synthetic validation only. No Office or Qwen request was made.' | Set-Content -LiteralPath (Join-Path $folder 'suite.log')
    $html=[Scribble.Testing.TestLabSuiteReport]::BuildHtml($state,@($result))
    Assert ($html.Contains('END_OF_ERROR') -and $html.Contains('2026-09-09T12:00:01Z')) 'Report omitted diagnostic tail or timestamps.'
    Assert ($html.Contains('&lt;script&gt;') -and -not $html.Contains('<script>alert')) 'Report executes source text.'
    Assert ((Get-Content -LiteralPath (Join-Path $folder 'summary.txt') -Raw).Contains('EX01 Excel')) 'Pasteable summary missing case.'
    Assert ((Get-Content -LiteralPath (Join-Path $folder 'summary.txt') -Raw).Contains('END_OF_ERROR')) 'Copy summary lost the first failure diagnostic.'
    $reportPath=[Scribble.Testing.TestLabSuiteReport]::Create($state,@($result))
    Assert ($reportPath.EndsWith('report.html') -and (Test-Path $reportPath)) 'Suite must produce HTML without a PDF renderer.'
    Assert (-not (Test-Path (Join-Path $folder 'report.pdf'))) 'Suite unexpectedly created a PDF.'
    $diagnostics=Get-Content -LiteralPath (Join-Path $folder 'diagnostics.txt') -Raw
    Assert ($diagnostics.Contains('<script>alert(1)</script>') -and $diagnostics.Contains('END_OF_ERROR')) 'Diagnostic text lost escaped error content.'
    $payload=('recorded event ' * 3000)+'TAIL_SENTINEL'
    $parts=[Scribble.Testing.TestLabSuiteReport]::SplitDiagnostics($state.id,$payload)
    $reassembled=($parts | ForEach-Object { $_.Substring($_.IndexOf("`n")+1) }) -join ''
    Assert ($parts.Length -gt 1 -and $reassembled -ceq $payload) 'Relay chunks lost or duplicated diagnostic text.'
    Assert ($html.Contains('data-copy=') -and $html.Contains('script-src')) 'HTML relay controls missing.'
    Write-Output "Suite HTML sample: $reportPath"
    Write-Output 'PASS: 16 cases, exact phases, source ownership, lease expiry, Chrome controller exclusivity, startup stderr, ZIP traversal, suite report.'
} finally {
    [Scribble.Testing.TestLabPreparation]::Stop($preparation)
    [Scribble.Testing.TestLab]::Disable()
    $state.expires=[DateTime]::UtcNow.AddSeconds(-1);[Scribble.Testing.TestLabSuite]::Save($state)
}
