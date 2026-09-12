#requires -Version 5.1
param([string]$AssemblyPath=(Join-Path $PSScriptRoot '../../src/Scribble/bin/Release/Scribble.dll'),
      [string]$KitRoot=(Join-Path $PSScriptRoot 'generated/scribble-test-kit-v1'))
$ErrorActionPreference='Stop'
$assembly=(Resolve-Path -LiteralPath $AssemblyPath).Path
$kit=(Resolve-Path -LiteralPath $KitRoot).Path
Add-Type -Path $assembly
function Assert($value,$message) { if (-not $value) { throw $message } }
$root=[Scribble.Testing.TestLab]::Root
Assert (-not (Test-Path -LiteralPath (Join-Path $root 'session.bin'))) 'Do not run preparation tests during an operator session.'
foreach ($script in Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'operator') -Filter *.ps1) {
    $tokens=$null;$errors=$null
    [void][Management.Automation.Language.Parser]::ParseFile($script.FullName,[ref]$tokens,[ref]$errors)
    Assert ($errors.Count -eq 0) ("PowerShell syntax: " + $script.Name + " " + $errors)
}
$preparationSource=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'operator/Prepare-ScribbleTestCase.ps1') -Raw
Assert ($preparationSource -match "Start-Process\s+-FilePath\s+\`$executable\s+-PassThru") 'Office preparation must launch a missing app as an interactive process.'
Assert ($preparationSource -notmatch "New-Object\s+-ComObject\s+\(\`$name\+'\.Application'\)") 'Office preparation must not bind a new app lifetime only to the helper COM client.'
Assert ($preparationSource -notmatch "if\s*\(\`$started\.HasExited\)\s*\{\s*throw") 'An Office/DDE launcher handoff must not be mistaken for application failure.'
$cases=Get-Content -LiteralPath (Join-Path $kit 'operator/cases.json') -Raw | ConvertFrom-Json
$plans=@{}
foreach ($case in $cases) {
    $output=& (Join-Path $PSScriptRoot 'operator/Prepare-ScribbleTestCase.ps1') -CaseId $case.id -FixtureRoot $kit -AssemblyPath $assembly -PlanOnly
    $plan=($output | Where-Object { $_ -notlike 'Loaded Test Lab*' } | Out-String) | ConvertFrom-Json
    Assert ($plan.apps -contains $case.host) ('Missing source app for '+$case.id)
    foreach ($p in @($plan.documents)+@($plan.mail_paths)+@($plan.pages)+@($plan.pdfs)) { Assert ($p.StartsWith('inputs/') -and (Test-Path -LiteralPath (Join-Path $kit $p))) ('Invalid prepared input: '+$p) }
    $plans[$case.id]=$plan
}
Assert ($plans.XA01.apps -contains 'Excel' -and $plans.XA01.apps -contains 'PowerPoint' -and $plans.XA01.mail_paths.Count -eq 3) 'Email-to-deck setup incomplete.'
Assert ($plans.PP02.documents -contains 'inputs/powerpoint/Atlas-start.pptx' -and -not ($plans.PP02.apps -contains 'Excel')) 'Supporting CSV incorrectly created an Excel pane prerequisite.'
Assert ($plans.XA03.apps -contains 'Chrome' -and $plans.XA03.apps -contains 'Excel' -and $plans.XA03.apps -contains 'Outlook') 'Browser handoff apps missing.'
$worker=$null;$duplicate=$null
try {
    [Scribble.Testing.TestLab]::Enable($kit)
    foreach ($case in [Scribble.Testing.TestLab]::Cases()) {
        $context=@([Scribble.Testing.TestLabPreparation]::ContextFiles($case))
        Assert ($context.Count -le 3) ('Too many context attachments: '+$case.id)
        if ($case.id -eq 'PP01') { Assert ($context.Count -eq 3 -and @($context | Where-Object { $_.EndsWith('.pdf') }).Count -eq 1) 'PowerPoint context should include sales, budget and operations PDF.' }
    }
    $session=[Scribble.Testing.TestLab]::Status()
    $serverScript=Join-Path $PSScriptRoot 'operator/Start-ScribbleFixtureServer.ps1'
    $args=@('-NoProfile','-ExecutionPolicy','Bypass','-File',('"'+$serverScript+'"'),'-AssemblyPath',('"'+$assembly+'"'))
    $worker=Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $args -WindowStyle Hidden -PassThru
    $receipt=Join-Path $root ('fixture-server-'+$session.session_id+'.json')
    $server=$null
    for ($i=0;$i -lt 60;$i++) { if (Test-Path -LiteralPath $receipt) { try { $server=Get-Content $receipt -Raw | ConvertFrom-Json; break } catch {} }; Start-Sleep -Milliseconds 100 }
    Assert ($null -ne $server) 'Fixture server did not start.'
    $url='http://127.0.0.1:'+$server.port
    $client=New-Object Net.WebClient
    try {
        foreach ($page in @('index.html','operations.html','archive.html')) {
            $bytes=$client.DownloadData($url+'/'+$page)
            Assert ([Scribble.Testing.TestLab]::Hash($bytes) -eq [Scribble.Testing.TestLab]::FileHash((Join-Path $kit ('inputs/browser/'+$page)))) ('Served bytes differ: '+$page)
        }
        foreach ($path in @('/evaluator-only/answers.json','/../manifest.json','/','/index.html?source=oracle')) {
            if ($path -eq '/') { continue }
            $rejected=$false
            try { [void]$client.DownloadData($url+$path) } catch [Net.WebException] { $rejected=$_.Exception.Response.StatusCode -eq 404 }
            Assert $rejected ('Server exposed unlisted path '+$path)
        }
        $head=Invoke-WebRequest -UseBasicParsing -Method Head -Uri ($url+'/index.html') -TimeoutSec 3
        Assert ($head.RawContentLength -eq 0) 'HEAD returned a response body.'
    } finally { $client.Dispose() }
    $duplicate=Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $args -WindowStyle Hidden -PassThru
    Assert ($duplicate.WaitForExit(10000) -and $duplicate.ExitCode -eq 0) 'Second server did not reuse the active session.'
    Assert (-not $worker.HasExited) 'Original server exited unexpectedly.'
    [Scribble.Testing.TestLab]::Disable()
    Assert ($worker.WaitForExit(5000)) 'Server did not stop after disable.'
    Write-Output 'PASS: 16 case plans, durable interactive Office launch, prerequisite apps/files, bounded context, exact fixture HTTP bytes, oracle exclusion, HEAD, duplicate server and shutdown.'
} finally {
    [Scribble.Testing.TestLab]::Disable()
    foreach ($process in @($worker,$duplicate)) { if ($process) { if (-not $process.HasExited) { $process.Kill() }; $process.Dispose() } }
}
