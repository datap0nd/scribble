#requires -Version 5.1
[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$CaseId,
      [string]$FixtureRoot=(Split-Path $PSScriptRoot -Parent), [string]$AssemblyPath,
      [string]$ReportPath, [switch]$PlanOnly, [switch]$Suite)
. (Join-Path $PSScriptRoot 'TestLab.Common.ps1') -AssemblyPath $AssemblyPath
$fixture=(Resolve-Path -LiteralPath $FixtureRoot).Path
$manifest=[Scribble.Testing.TestLab]::VerifyKit($fixture)
$allCases=Get-Content -LiteralPath (Join-Path $fixture 'operator/cases.json') -Raw | ConvertFrom-Json
$case=@($allCases | Where-Object id -eq $CaseId)
if ($case.Count -ne 1) { throw 'Unknown case ID.' }; $case=$case[0]
$index=Get-Content -LiteralPath (Join-Path $fixture 'operator/mail-index.json') -Raw | ConvertFrom-Json
$mail=@($index | Where-Object { $case.inputs -contains $_.path })
$paths=@($case.inputs) + @($mail | ForEach-Object { $_.attachments })
$paths=@($paths | Select-Object -Unique)
foreach ($relative in $paths) {
    if (-not $relative.StartsWith('inputs/') -or -not @($manifest.files | Where-Object { $_.path -eq $relative -and $_.role -eq 'input' }).Count) { throw "Invalid case input: $relative" }
    [void][Scribble.Testing.TestLab]::SafeChild($fixture,$relative)
}
$formats=@{'.xlsx'='Excel';'.csv'='Excel';'.pptx'='PowerPoint';'.docx'='Word'}
$outputApps=@{xlsx='Excel';pptx='PowerPoint';docx='Word';msg='Outlook'}
$documents=@($paths | Where-Object { $formats.ContainsKey([IO.Path]::GetExtension($_)) })
$apps=@($case.host) + @($documents | ForEach-Object { $formats[[IO.Path]::GetExtension($_)] }) + @($case.artifacts | ForEach-Object { $outputApps[$_] })
$apps=@($apps | Where-Object { $_ } | Select-Object -Unique)
if (@($apps | Where-Object { $_ -notin @('Excel','PowerPoint','Word','Outlook','Chrome') }).Count) { throw 'Unsupported app in case.' }
$plan=[ordered]@{case_id=$CaseId;host=$case.host;apps=$apps;documents=$documents;mail_paths=@($mail | ForEach-Object { $_.path });pages=@($paths | Where-Object { $_.StartsWith('inputs/browser/') });pdfs=@($paths | Where-Object { $_.EndsWith('.pdf') })}
if ($PlanOnly) { $plan | ConvertTo-Json -Depth 6; return }
$session=[Scribble.Testing.TestLab]::Status()
if (-not $session -or $session.fixture_root -ne $fixture) { throw 'Enable Test Lab for this extracted kit first.' }
if ($session.run_id) { throw 'Finish the active run before preparing a case.' }
$completed=New-Object 'Collections.Generic.List[string]'
$remaining=New-Object 'Collections.Generic.List[string]'
function Report($state) {
    $value=[ordered]@{case_id=$CaseId;status=$state;completed=@($completed.ToArray());remaining=@($remaining.ToArray())}
    if ($ReportPath) { $value | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $ReportPath -Encoding UTF8 }
    else { $value | ConvertTo-Json -Depth 5 | Write-Output }
}
function Assert-Idle {
    $now=[Scribble.Testing.TestLab]::Status()
    if (-not $now -or $now.session_id -ne $session.session_id -or $now.run_id) { throw 'Preparation stopped because the test session changed or a run started.' }
}
function Progress($text) { Write-Output ([DateTime]::UtcNow.ToString('O')+' '+$text) }
function Get-App($name) {
    try { return ,([Runtime.InteropServices.Marshal]::GetActiveObject($name+'.Application')) }
    catch { return ,(New-Object -ComObject ($name+'.Application')) }
}
$progids=@{Excel='Scribble.ExcelAddIn';PowerPoint='Scribble.PowerPointAddIn';Word='Scribble.WordAddIn';Outlook='Scribble.AddIn'}
Report 'preparing'
$preparationLock=$null
trap {
    if ($null -ne $remaining) { $remaining.Add($_.Exception.Message); Report 'failed' }
    if ($preparationLock) { $preparationLock.Dispose() }
    throw $_
}
$preparationLock=[IO.File]::Open((Join-Path ([Scribble.Testing.TestLab]::Root) 'preparation.lock'),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
try {
foreach ($name in $apps | Where-Object { $_ -ne 'Chrome' }) {
    try {
        Assert-Idle
        Progress "Connecting to $name..."
        $app=Get-App $name
        if ($null -eq $app) { throw "No $name application object was returned." }
        if ($name -eq 'PowerPoint') { $app.Visible=-1 }
        elseif ($name -ne 'Outlook') { $app.Visible=$true }
        Progress "$name connected; checking Scribble add-in..."
        try {
            if (-not $app.COMAddIns.Item($progids[$name]).Connect) { $remaining.Add("Enable the Scribble add-in in $name.") }
            else { $completed.Add("$name`: Scribble add-in connected.") }
        } catch { $remaining.Add("Scribble add-in not found in $name. Install or repair Scribble.") }
        # Open CSVs before native workbooks so the case workbook is active last.
        foreach ($relative in @($documents | Where-Object { $formats[[IO.Path]::GetExtension($_)] -eq $name } | Sort-Object { [IO.Path]::GetExtension($_) -ne '.csv' })) {
            Assert-Idle
            Progress "Opening $relative in $name..."
            $path=[Scribble.Testing.TestLab]::SafeChild($fixture,$relative)
            if ($name -eq 'Excel') { $collection=$app.Workbooks }
            elseif ($name -eq 'PowerPoint') { $collection=$app.Presentations }
            else { $collection=$app.Documents }
            $document=$null
            for ($i=1;$i -le $collection.Count;$i++) { if ($collection.Item($i).FullName -eq $path) { $document=$collection.Item($i); break } }
            if ($document -and -not $document.Saved) { throw "The fixture $relative has unsaved edits. Save a separate output and close it before preparing again." }
            if (-not $document) {
                if ($name -eq 'Excel') {
                    for ($j=1;$j -le $app.Workbooks.Count;$j++) {
                        $opened=$app.Workbooks.Item($j)
                        if ([IO.Path]::GetFileName($opened.FullName) -eq [IO.Path]::GetFileName($path) -and $opened.FullName -ne $path) {
                            throw "Excel already has a different workbook named $([IO.Path]::GetFileName($path)) open at $($opened.FullName). Save its generated output and close that workbook before retrying this case."
                        }
                    }
                    $document=$app.Workbooks.Open($path,0,$true)
                }
                elseif ($name -eq 'PowerPoint') { $document=$app.Presentations.Open($path,-1,0,-1) }
                else { $document=$app.Documents.Open($path,$false,$true) }
            }
            if ($null -eq $document) { throw "$name did not open $relative. Check its visible file-open dialog." }
            if ($name -eq 'PowerPoint') { $document.Windows.Item(1).Activate() } else { $document.Activate() }
            $completed.Add("Opened $relative in $name (source opened read-only).")
        }
        if ($name -eq 'Outlook' -and $mail.Count -gt 0) {
            Assert-Idle
            Progress 'Importing and selecting Outlook fixtures...'
            & (Join-Path $PSScriptRoot 'Import-ScribbleTestMail.ps1') -FixtureRoot $fixture -AssemblyPath $resolvedAssembly -CaseId $CaseId
            $completed.Add("Selected $($mail.Count) synthetic messages in the isolated Outlook PST.")
        }
        elseif ($name -eq 'Outlook') {
            $explorer=$app.ActiveExplorer()
            if ($explorer) { $explorer.Display() } else { $app.Session.GetDefaultFolder(6).GetExplorer().Display() }
            $completed.Add('Opened Outlook for the expected unsent draft.')
        }
    } catch { $remaining.Add("$name`: $($_.Exception.Message) | $($_.InvocationInfo.PositionMessage)") }
    Report 'preparing'
}
foreach ($relative in $plan.pdfs) {
    try { Assert-Idle; Start-Process -FilePath ([Scribble.Testing.TestLab]::SafeChild($fixture,$relative)); $completed.Add("Opened $relative in the PDF viewer.") }
    catch { $remaining.Add("PDF viewer: $($_.Exception.Message)") }
}
if ($apps -contains 'Chrome') {
    try {
        Assert-Idle
        $chrome=@((Join-Path $env:ProgramFiles 'Google/Chrome/Application/chrome.exe'),(Join-Path ${env:ProgramFiles(x86)} 'Google/Chrome/Application/chrome.exe'),(Join-Path $env:LOCALAPPDATA 'Google/Chrome/Application/chrome.exe')) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
        if (-not $chrome) { throw 'Google Chrome was not found. Install Chrome and the Scribble extension.' }
        $serverScript=Join-Path $PSScriptRoot 'Start-ScribbleFixtureServer.ps1'
        $serverReceipt=Join-Path ([Scribble.Testing.TestLab]::Root) ('fixture-server-'+$session.session_id+'.json')
        $shell=Join-Path $PSHOME 'powershell.exe'
        Start-Process -FilePath $shell -WindowStyle Hidden -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"'+$serverScript+'"'),'-AssemblyPath',('"'+$resolvedAssembly+'"'))
        $ready=$false
        for ($i=0;$i -lt 30;$i++) {
            try {
                $server=Get-Content -LiteralPath $serverReceipt -Raw | ConvertFrom-Json
                if ($server.session_id -ne $session.session_id -or $server.manifest_sha256 -ne $session.manifest_sha256) { throw 'Stale server receipt.' }
                $baseUrl='http://127.0.0.1:'+[int]$server.port
                $response=Invoke-WebRequest -UseBasicParsing -Uri ($baseUrl+'/index.html') -TimeoutSec 1
                if ($response.Headers['X-Scribble-Fixture'] -ne $session.session_id) { throw 'Unexpected fixture server.' }
                $ready=$true; break
            } catch { Start-Sleep -Milliseconds 250 }
        }
        if (-not $ready) { throw 'The local fixture server did not become ready. Disable and re-enable Test Lab, then prepare again.' }
        $urls=@($plan.pages | ForEach-Object { $baseUrl+'/'+[IO.Path]::GetFileName($_) })
        if ($urls.Count -eq 0) { $urls=@($baseUrl+'/operations.html') }
        # Keep the operations page active after opening the case's supporting pages.
        $urls=@($urls | Sort-Object { $_.EndsWith('/operations.html') })
        Start-Process -FilePath $chrome -ArgumentList (@('--new-window') + $urls)
        $completed.Add("Opened Chrome fixture pages at $baseUrl. The server stops when Test Lab is disabled or expires.")
        if (-not $Suite) { $remaining.Add('Open the Scribble side panel in Chrome and verify its native connection; extension readiness cannot be confirmed from Office.') }
    } catch { $remaining.Add("Chrome: $($_.Exception.Message)") }
}
if ($case.host -ne 'Chrome') {
    try {
        Assert-Idle
        $origin=Get-App $case.host
        if ($case.host -eq 'Outlook') { $origin.ActiveExplorer().Activate() }
        else { $origin.ActiveWindow.Activate() }
    } catch { $remaining.Add("Return to $($case.host) to start this case: $($_.Exception.Message)") }
}
if (-not $Suite) { $remaining.Add("Open Scribble in $($case.host), select $CaseId and Start case. Record the screen before submitting the prepared prompt.") }
if ($Suite -and $remaining.Count -gt 0) { Report 'failed' } else { Report 'finished' }
} finally { $preparationLock.Dispose() }
