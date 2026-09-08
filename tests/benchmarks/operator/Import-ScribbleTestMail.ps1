#requires -Version 5.1
[CmdletBinding()]
param([string]$FixtureRoot=(Split-Path $PSScriptRoot -Parent),[string]$AssemblyPath)
. (Join-Path $PSScriptRoot 'TestLab.Common.ps1') -AssemblyPath $AssemblyPath
$fixture=(Resolve-Path -LiteralPath $FixtureRoot).Path
[void][Scribble.Testing.TestLab]::VerifyKit($fixture)
$index=Get-Content -LiteralPath (Join-Path $fixture 'operator\mail-index.json') -Raw | ConvertFrom-Json
$importRoot=Join-Path ([Scribble.Testing.TestLab]::Root) 'mail'
New-Item -ItemType Directory -Path $importRoot -Force | Out-Null
$pst=Join-Path $importRoot 'Atlas-v1.pst'
$outlook=New-Object -ComObject Outlook.Application
$session=$outlook.Session
$store=@($session.Stores | Where-Object { $_.FilePath -eq $pst }) | Select-Object -First 1
if (-not $store) { $session.AddStoreEx($pst,2); $store=@($session.Stores | Where-Object { $_.FilePath -eq $pst }) | Select-Object -First 1 }
if (-not $store) { throw 'Could not create the isolated Atlas local PST.' }
$root=$store.GetRootFolder()
$root.Name='Scribble synthetic Atlas v1'
$folder=@($root.Folders | Where-Object Name -eq 'Atlas fixtures') | Select-Object -First 1
if (-not $folder) { $folder=$root.Folders.Add('Atlas fixtures',6) }
$receipts=@()
foreach($source in $index) {
    $existing=@($folder.Items | Where-Object Subject -eq $source.subject)
    if ($existing.Count -gt 1) { throw "Duplicate imported subject: $($source.subject). Resolve manually in the test PST." }
    if ($existing.Count -eq 1) { $item=$existing[0] }
    else {
        $item=$folder.Items.Add('IPM.Note')
        $item.Subject=$source.subject; $item.Body=$source.body; $item.To='review@example.test'
        $pa=$item.PropertyAccessor
        $pa.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x0C1A001F',[string]$source.sender)
        $pa.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x0C1F001F',[string]$source.sender)
        $pa.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x0C1E001F','SMTP')
        $pa.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x00390040',[datetime]::Parse($source.date).ToUniversalTime())
        $pa.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x0E060040',[datetime]::Parse($source.date).ToUniversalTime())
        foreach($attachment in $source.attachments) {
            $path=[Scribble.Testing.TestLab]::SafeChild($fixture,[string]$attachment)
            [void]$item.Attachments.Add($path)
        }
        # Received/read message flags: never invoke Send or submit to a transport.
        $pa.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x0E070003',1)
        $item.Save()
    }
    if($item.SenderEmailAddress -ne $source.sender -or $item.Attachments.Count -ne $source.attachments.Count) { throw "Native import verification failed for $($source.subject). The incomplete fixture remains isolated in the local PST." }
    if(($item.Body -replace '\s+',' ').Trim() -ne ($source.body -replace '\s+',' ').Trim() -or [math]::Abs(($item.ReceivedTime.ToUniversalTime()-[datetime]::Parse($source.date).ToUniversalTime()).TotalSeconds) -gt 1) { throw "Body/date verification failed for $($source.subject)." }
    for($i=1;$i -le $item.Attachments.Count;$i++) {
        $temporary=Join-Path $importRoot ('verify-'+[guid]::NewGuid().ToString('N')+'.bin')
        try {
            $item.Attachments.Item($i).SaveAsFile($temporary)
            $original=[Scribble.Testing.TestLab]::SafeChild($fixture,[string]$source.attachments[$i-1])
            if([Scribble.Testing.TestLab]::FileHash($temporary) -ne [Scribble.Testing.TestLab]::FileHash($original)) { throw 'Imported attachment bytes differ from fixture.' }
        } finally { if(Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary } }
    }
    $receipts += [ordered]@{source_id=$source.id;entry_id=$item.EntryID;store_id=$folder.StoreID;subject=$item.Subject;sender=$item.SenderEmailAddress;received_utc=$item.ReceivedTime.ToUniversalTime().ToString('O');attachments=$item.Attachments.Count}
}
$receipts | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $importRoot 'import-receipts.json') -Encoding utf8
$explorer=$outlook.ActiveExplorer()
if($explorer) { $explorer.CurrentFolder=$folder } else { $folder.GetExplorer().Display() }
Write-Output 'Imported five synthetic messages into the isolated local PST. Select the exact case messages and Add email to Scribble. Unscoped mailbox search does not search this test folder.'
