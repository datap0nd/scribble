#requires -Version 5.1
param([string]$AssemblyPath=(Join-Path $PSScriptRoot '..\..\src\Scribble\bin\Release\Scribble.dll'),[string]$KitRoot=(Join-Path $PSScriptRoot 'generated\scribble-test-kit-v1'))
$ErrorActionPreference='Stop'
Add-Type -Path (Resolve-Path $AssemblyPath).Path
Add-Type -AssemblyName System.IO.Compression.FileSystem
$root=[Scribble.Testing.TestLab]::Root
if(Test-Path -LiteralPath (Join-Path $root 'session.bin')) { throw 'An operator descriptor already exists. Do not run infrastructure tests during a user lab session.' }
$testOutput=Join-Path $env:TEMP ('scribble-lab-tests-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testOutput | Out-Null
function Assert($value,$message) { if(-not $value) { throw $message } }
function Reject([scriptblock]$action,$message) { $rejected=$false;try { & $action }catch { $rejected=$true };Assert $rejected $message }
try {
    Assert ($null -eq [Scribble.Testing.TestLab]::Status()) 'Default must be disabled'
    Reject { [Scribble.Testing.TestLab]::SafeChild($testOutput,'..\escape') } 'Traversal accepted'
    Reject { [Scribble.Testing.TestLab]::Start('EX01','Excel',$false) } 'Unconfirmed start accepted'
    [Scribble.Testing.TestLab]::Enable((Resolve-Path $KitRoot).Path)
    Assert ([Scribble.Testing.TestLab]::Cases().Count -eq 16) 'Cases missing'
    $run=[Scribble.Testing.TestLab]::Start('EX01','Excel',$true)
    Reject { [Scribble.Testing.TestLab]::Start('PP01','PowerPoint',$true) } 'Concurrent active case accepted'
    Reject { [Scribble.Testing.TestLab]::Collect($run.run_id,(Join-Path $KitRoot 'evaluator-only\reference-deck.pptx')) } 'Reference accepted as generated output'
    [Scribble.Testing.TestLab]::CheckInputFile((Join-Path $KitRoot 'inputs\excel\Atlas-input.xlsx'))
    for($i=0;$i -lt 150;$i++) { [Scribble.Testing.TestLab]::Record($run.run_id,'test-task','roundtrip',@{index=$i;value='Bearer secret-fixture-token'}) }
    [Scribble.Testing.TestLab]::Marker('video zero')
    Reject { [Scribble.Testing.TestLab]::Export($run.run_id,$testOutput) } 'Exported active case'
    [Scribble.Testing.TestLab]::Finish($false)
    $archive=[Scribble.Testing.TestLab]::Export($run.run_id,$testOutput)
    $z=[IO.Compression.ZipFile]::OpenRead($archive)
    try {
        $reader=[IO.StreamReader]::new($z.GetEntry('timeline.jsonl').Open());$text=$reader.ReadToEnd();$reader.Dispose()
        Assert (-not $text.Contains('secret-fixture-token')) 'Secret leaked'
        Assert ($text.Contains('[REDACTED]')) 'Redaction missing'
        Assert (@($text -split "`n" | Where-Object { $_ -match 'roundtrip' }).Count -eq 150) 'Trace rotated away events'
        $reader=[IO.StreamReader]::new($z.GetEntry('run.json').Open());$report=$reader.ReadToEnd()|ConvertFrom-Json;$reader.Dispose()
        Assert $report.trace_complete 'Complete trace was lost'
        Assert ($report.missing_artifacts -contains 'xlsx') 'Missing file not reported'
    } finally { $z.Dispose() }
    $bad=[Scribble.Testing.TestLab]::Start('EX03','Excel',$true)
    Reject { [Scribble.Testing.TestLab]::CheckInputFile((Join-Path $KitRoot 'inputs\data\sales.csv')) } 'Different case source accepted'
    Reject { [Scribble.Testing.TestLab]::CheckInputFile((Join-Path $KitRoot 'evaluator-only\answers.json')) } 'Oracle accepted as source'
    [Scribble.Testing.TestLab]::Finish($true)
    Assert (-not [Scribble.Testing.TestLab]::GetRun($bad.run_id).trace_complete) 'Source mismatch certified complete'
    $quota=[Scribble.Testing.TestLab]::Start('EX03','Excel',$true)
    $quotaFile=Join-Path ([Scribble.Testing.TestLab]::RunDirectory($quota.run_id)) 'events\quota-test.bin'
    $stream=[IO.File]::Create($quotaFile);$stream.SetLength(250L*1024*1024+1);$stream.Dispose()
    try { [Scribble.Testing.TestLab]::Record($quota.run_id,'test','quota',@{});Assert ([Scribble.Testing.TestLab]::CaptureState() -eq 'incomplete') 'Quota silently rotated' }
    finally { Remove-Item -LiteralPath $quotaFile }
    [Scribble.Testing.TestLab]::Finish($false)
    $expired=[Scribble.Testing.TestLab]::Status();$expired.expires_utc=[datetime]::UtcNow.AddMinutes(-1)
    $bytes=[Text.Encoding]::UTF8.GetBytes([Scribble.Testing.TestLab]::Serialize($expired))
    [IO.File]::WriteAllBytes((Join-Path $root 'session.bin'),[Security.Cryptography.ProtectedData]::Protect($bytes,$null,[Security.Cryptography.DataProtectionScope]::CurrentUser))
    Assert ($null -eq [Scribble.Testing.TestLab]::Status()) 'Expired descriptor accepted'
    [Scribble.Testing.TestLab]::Disable()
    Assert ($null -eq [Scribble.Testing.TestLab]::Status()) 'Disable failed'
    $flat='<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Native fixture</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>'
    $flatPath=Join-Path $testOutput 'flat-opc.docx'
    [Scribble.Testing.BenchmarkArtifactCollector]::SaveFlatOpc($flat,$flatPath)
    $z=[IO.Compression.ZipFile]::OpenRead($flatPath)
    try { Assert ($null -ne $z.GetEntry('[Content_Types].xml')) 'Missing content types';Assert ($null -ne $z.GetEntry('word/document.xml')) 'Missing Word document part' } finally { $z.Dispose() }
    Reject { [Scribble.Testing.BenchmarkArtifactCollector]::SaveFlatOpc($flat.Replace('/word/document.xml','/../escape'),(Join-Path $testOutput 'invalid.docx')) } 'Unsafe Flat OPC part accepted'
    Write-Output 'PASS: default off, confirmation, path guard, single run, source hashes, 150 retained events, secret redaction, lifecycle, export and oracle exclusion.'
    Write-Output "Test evidence: $testOutput"
} finally { [Scribble.Testing.TestLab]::Disable() }
