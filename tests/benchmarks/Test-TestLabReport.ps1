#requires -Version 5.1
param([string]$AssemblyPath=(Join-Path $PSScriptRoot '../../src/Scribble/bin/Release/Scribble.dll'),
      [string]$KitRoot=(Join-Path $PSScriptRoot 'generated/scribble-test-kit-v1'),[switch]$RenderPdf)
$ErrorActionPreference='Stop'
$resolvedAssembly=(Resolve-Path -LiteralPath $AssemblyPath).Path
Add-Type -TypeDefinition 'using System;using System.IO;using System.Reflection;public static class TestLabAssemblyResolver{public static void Install(string folder){AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{var p=Path.Combine(folder,new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};}}'
[TestLabAssemblyResolver]::Install((Split-Path $resolvedAssembly))
foreach($dependency in @('Microsoft.Extensions.Logging.Abstractions.dll','PdfSharp.Shared.dll','PdfSharp.System.dll','PdfSharp-gdi.dll')) {
    [void][Reflection.Assembly]::LoadFrom((Join-Path (Split-Path $resolvedAssembly) $dependency))
}
Add-Type -Path $resolvedAssembly
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Assert($value,$message) { if(-not $value) {throw $message} }
$root=[Scribble.Testing.TestLab]::Root
Assert (-not (Test-Path -LiteralPath (Join-Path $root 'session.bin'))) 'Do not run report tests during an operator session.'
$destination=Join-Path $PSScriptRoot ('generated/report-test-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $destination -Force | Out-Null
try {
    [Scribble.Testing.TestLab]::Enable((Resolve-Path -LiteralPath $KitRoot).Path)
    $run=[Scribble.Testing.TestLab]::Start('EX01','Excel',$true)
    [Scribble.Testing.TestLab]::Record($run.run_id,'sample','task_started',@{note='Synthetic report validation, not a model run'})
    [Scribble.Testing.TestLab]::Record($run.run_id,'sample','tool_error',@{error=$true;message=('<script>alert(1)</script> synthetic timeout '+('Long diagnostic line ' * 100)+' END_OF_ERROR');unicode="M$([char]0x00fc)nchen / $([char]0x4e2d)$([char]0x6587)"})
    [Scribble.Testing.TestLab]::Record($run.run_id,'pane','pane_event',@{type='assistant';text='Synthetic report sample: the workbook could not be created. No actual Qwen request was made.'})
    [Scribble.Testing.TestLab]::Marker('Sample video marker 00:15')
    [Scribble.Testing.TestLab]::Record($run.run_id,'sample','task_completed',@{ok=$false})
    [Scribble.Testing.TestLab]::Finish($false)
    $zip=[Scribble.Testing.TestLab]::Export($run.run_id,$destination)
    $summary=Join-Path $destination 'summary.txt'
    $html=[Scribble.Testing.TestLabReport]::BuildHtml($zip,$summary)
    Assert ($html.Contains('END_OF_ERROR') -and $html.Contains('Sample video marker 00:15')) 'Error tail or timestamps missing.'
    Assert ($html.Contains('&lt;script&gt;') -and -not $html.Contains('<script>alert')) 'Report rendered untrusted HTML.'
    Assert ($html.Contains('MISSING OUTPUTS') -and $html.Contains('correctness not yet reviewed')) 'Report falsely implies pass.'
    Assert ((Get-Content $summary -Raw).Contains($run.run_id)) 'Pasteable summary lacks run identity.'
    if ($RenderPdf) {
        $pdf=[Scribble.Testing.TestLabReport]::Create($zip)
        Assert ((Get-Item -LiteralPath $pdf).Length -gt 1000) 'Empty PDF.'
        $archive=[IO.Compression.ZipFile]::OpenRead($zip)
        try {
            Assert ($null -ne $archive.GetEntry('report.pdf') -and $null -ne $archive.GetEntry('summary.txt')) 'Shareable reports missing from ZIP.'
            $reader=New-Object IO.StreamReader($archive.GetEntry('export-manifest.json').Open())
            try { $manifest=$reader.ReadToEnd() | ConvertFrom-Json } finally { $reader.Dispose() }
            foreach($file in $manifest.files) {
                $stream=$archive.GetEntry($file.path).Open();$sha=[Security.Cryptography.SHA256]::Create()
                try { $hash=[BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','').ToLowerInvariant();Assert ($hash -eq $file.sha256) ('Report export hash mismatch: '+$file.path) }
                finally { $stream.Dispose();$sha.Dispose() }
            }
        } finally { $archive.Dispose() }
        Write-Output "Sample PDF: $pdf"
    }
    $eventFolder=Join-Path ([Scribble.Testing.TestLab]::RunDirectory($run.run_id)) 'events'
    [IO.File]::WriteAllBytes((Join-Path $eventFolder 'corrupt.bin'),[byte[]](1,2,3,4))
    $brokenZip=[Scribble.Testing.TestLab]::Export($run.run_id,$destination)
    $brokenHtml=[Scribble.Testing.TestLabReport]::BuildHtml($brokenZip,$summary)
    Assert ($brokenHtml.Contains('Recorded event could not be read:') -and $brokenHtml.Contains('corrupt.bin')) 'Unreadable event was silently dropped.'
    Write-Output 'PASS: screenshot summary, full diagnostic tail, escaped HTML, missing-output status, video markers and report export.'
} finally { [Scribble.Testing.TestLab]::Disable() }
