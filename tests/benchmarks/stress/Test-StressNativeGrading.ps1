[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$ScribbleAssembly,
    [string]$CorpusRoot = '',
    [string]$ExpectedManifestHash = '',
    [string]$CaseId = 'PP03',
    [string]$CandidatePptx = ''
)

# This preflight exercises the presentation, native-chart, and artifact rules
# derived from a sealed case. It does not start a Test Lab session, contact a
# model, or claim the clean source reference solves the entire task.
$ErrorActionPreference='Stop'
if (-not $CorpusRoot) { $CorpusRoot=Join-Path $PSScriptRoot '../generated/stress-corpus' }
$root=[IO.Path]::GetFullPath($CorpusRoot)
$assemblyPath=[IO.Path]::GetFullPath($ScribbleAssembly)
$parentManifest=Join-Path $root 'manifest.json'
$parentHash=(Get-FileHash -LiteralPath $parentManifest -Algorithm SHA256).Hash.ToLowerInvariant()
if ($ExpectedManifestHash -and $parentHash -ne $ExpectedManifestHash.ToLowerInvariant()) { throw 'The source manifest hash differs from the requested seal' }
[void][Reflection.Assembly]::LoadFrom($assemblyPath)
Add-Type -AssemblyName System.Web.Extensions
Add-Type -AssemblyName System.IO.Compression
$verified=[Scribble.Testing.TestLab]::VerifyKit($root)
if ($verified.suite_id -ne 'scribble-stress-v1') { throw 'The native preflight requires the real sealed stress corpus' }
$serializer=New-Object Web.Script.Serialization.JavaScriptSerializer
$serializer.MaxJsonLength=8388608
$cases=$serializer.Deserialize((Get-Content -LiteralPath (Join-Path $root 'operator/cases.json') -Raw),[Scribble.Testing.LabCase[]])
$originalCase=@($cases | Where-Object id -eq $CaseId)
if ($originalCase.Count -ne 1 -or $originalCase[0].host -ne 'PowerPoint') { throw 'Choose exactly one real PowerPoint case' }
$originalCase=$originalCase[0]
$originalOraclePath=Join-Path $root $originalCase.oracle_ref
$originalOracle=Get-Content -LiteralPath $originalOraclePath -Raw | ConvertFrom-Json
$rules=@($originalOracle.checks | Where-Object { $_.kind -eq 'presentation' -or $_.kind -eq 'native_chart' -or ($_.kind -eq 'native_artifact' -and $_.extension -eq 'pptx') })
if (@($rules | Where-Object kind -eq 'presentation').Count -ne 1) { throw 'The case must have exactly one presentation rule' }
$presentationRule=@($rules | Where-Object kind -eq 'presentation')[0]
$reference=Join-Path $root $presentationRule.reference_pptx
$source=@($originalCase.inputs | Where-Object { $_.EndsWith('.pptx') })
if ($source.Count -ne 1) { throw 'The case must name one source deck' }
$source=Join-Path $root $source[0]
if ($CandidatePptx) {
    $CandidatePptx=[IO.Path]::GetFullPath($CandidatePptx)
    if (-not [IO.File]::Exists($CandidatePptx)) { throw 'Candidate PowerPoint file does not exist' }
}
$preflightRoot=Join-Path $root ('.build/native-grading-preflight/'+[Guid]::NewGuid().ToString('N'))
$projection=Join-Path $preflightRoot 'projection'
[IO.Directory]::CreateDirectory($projection) | Out-Null
function Write-Json([string]$Path,$Value) {
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($Path)) | Out-Null
    [IO.File]::WriteAllText($Path,(ConvertTo-Json -InputObject $Value -Depth 40 -Compress),[Text.UTF8Encoding]::new($false))
}
$projectedCase=New-Object Scribble.Testing.LabCase
$projectedCase.id=$CaseId
$projectedCase.host='PowerPoint'
$projectedCase.prompt=$originalCase.prompt
$projectedCase.inputs=@($originalCase.inputs)
$projectedCase.artifacts=@('pptx')
$projectedCase.oracle_ref=$originalCase.oracle_ref
$projectedCase.timeout_seconds=900
Write-Json (Join-Path $projection 'operator/cases.json') @($projectedCase)
Write-Json (Join-Path $projection 'operator/mail-index.json') @()
$derivedOracle=[pscustomobject]@{
    schema=1;suite_id='scribble-stress-v1';id=$CaseId;checks=$rules
    preflight_scope='Native presentation-grader only; native source copies, zero model calls'
    original_oracle_sha256=(Get-FileHash -LiteralPath $originalOraclePath -Algorithm SHA256).Hash.ToLowerInvariant()
    original_manifest_sha256=$parentHash
}
Write-Json (Join-Path $projection $projectedCase.oracle_ref) $derivedOracle
foreach ($relative in @($projectedCase.inputs)+@($presentationRule.theme_ref,$presentationRule.reference_pptx)) {
    if ($relative -notmatch '^(inputs|evaluator-only)/' -or $relative.Contains('..')) { throw 'Invalid sealed projection path' }
    $target=Join-Path $projection $relative
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target)) | Out-Null
    Copy-Item -LiteralPath (Join-Path $root $relative) -Destination $target
}
$files=@(Get-ChildItem -LiteralPath $projection -File -Recurse | ForEach-Object {
    [pscustomobject]@{path=$_.FullName.Substring($projection.Length+1).Replace('\','/');size=$_.Length;sha256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()}
})
Write-Json (Join-Path $projection 'manifest.json') ([pscustomobject]@{schema=1;suite_id='scribble-stress-v1';parent_manifest_sha256=$parentHash;files=$files})
[void][Scribble.Testing.TestLab]::VerifyKit($projection)

$helper=@'
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Scribble.Testing;
public static class NativeGradingPreflightHelper {
  public static TestLabCheck[] Evaluate(string root, LabCase c, StressNative n, string runId, string zipPath) {
    var text = string.Join("\n", n.slides.SelectMany(s => s.shapes).Select(s => s.text ?? "")) + "\n" +
      string.Join("\n", n.slides.SelectMany(s => s.charts).SelectMany(ch => ch.series).SelectMany(s => new[] { s.name }.Concat(s.categories).Concat(s.values)));
    var capture = new StressReadback { run_id=runId, artifact_extension="pptx", native_readback=true,
      run_created_output=true, output_boundary=true, stress_native=n, text=text };
    using(var stream=File.Create(zipPath)) using(var zip=new ZipArchive(stream, ZipArchiveMode.Create)) {
      using(var writer=new StreamWriter(zip.CreateEntry("artifacts/PowerPoint-final-output-1-readback.json").Open())) writer.Write(TestLab.Serialize(capture));
      using(var writer=new StreamWriter(zip.CreateEntry("timeline.jsonl").Open())) writer.Write("");
    }
    var run=new LabRun { schema=1, run_id=runId, case_id=c.id, host="PowerPoint", suite_id="scribble-stress-v1",
      fixture_root=root, manifest_sha256=TestLab.FileHash(Path.Combine(root,"manifest.json")), input_paths=c.inputs, required_artifacts=c.artifacts };
    using(var stream=File.OpenRead(zipPath)) using(var zip=new ZipArchive(stream,ZipArchiveMode.Read))
      return TestLabStressEvaluator.Evaluate(zip,run,c,new TestLabCheck[0]);
  }
}
'@
Add-Type -TypeDefinition $helper -ReferencedAssemblies @($assemblyPath,'System.Core','System.IO.Compression','System.Web.Extensions')
$records=[Collections.Generic.List[object]]::new()
$existing=@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
$app=New-Object -ComObject PowerPoint.Application
$quit=(-not $existing) -and $app.Presentations.Count -eq 0
$originalHashes=@{reference=(Get-FileHash -LiteralPath $reference -Algorithm SHA256).Hash;source=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash}
$normalizations=[Collections.Generic.List[object]]::new()
$baseline=Join-Path $preflightRoot 'normalized-reference-baseline.pptx'
try {
    Copy-Item -LiteralPath $reference -Destination $baseline
    $baselinePresentation=$app.Presentations.Open($baseline,0,0,0)
    try {
        # The source authoring canvas used18CSSpx for compact bylines, which is
        #13.5nativept. The output contract requires14pt body text. Raise only
        #those textboxes in this private baseline, leaving the sealed source
        #and oracle unchanged. All isolated mutants share this same baseline.
        for ($s=1;$s -le $baselinePresentation.Slides.Count;$s++) {
            $slide=$baselinePresentation.Slides.Item($s)
            for ($i=1;$i -le $slide.Shapes.Count;$i++) {
                $shape=$slide.Shapes.Item($i)
                if ($shape.HasTable -ne -1 -and $shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                    $size=[double]$shape.TextFrame.TextRange.Font.Size
                    if ([math]::Abs($size-13.5) -lt 0.01) {
                        $shape.TextFrame.TextRange.Font.Size=14
                        $normalizations.Add([pscustomobject]@{slide=$s;shape=$i;from_points=$size;to_points=14})
                    }
                }
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shape)
            }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($slide)
        }
        $baselinePresentation.Save()
    } finally {
        $baselinePresentation.Close()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($baselinePresentation)
    }
    $variants=if ($CandidatePptx) { @('candidate') } else { @('clean_reference','defective_source','pink_header_only','pink_chart_only') }
    foreach ($variant in $variants) {
        $presentation=$null
        try {
            $owned=Join-Path $preflightRoot ($variant+'.pptx')
            $copySource=if ($variant -eq 'candidate') { $CandidatePptx } elseif ($variant -eq 'defective_source') { $source } else { $baseline }
            Copy-Item -LiteralPath $copySource -Destination $owned
            $presentation=$app.Presentations.Open($owned,0,0,0)
            $mutated=$false
            if ($variant -eq 'pink_header_only') {
                foreach ($slide in $presentation.Slides) {
                    foreach ($shape in $slide.Shapes) {
                        if (-not $mutated -and $shape.HasTable -eq -1) {
                            for ($c=1;$c -le $shape.Table.Columns.Count;$c++) {
                                $shape.Table.Cell(1,$c).Shape.Fill.ForeColor.RGB=[Scribble.Office.MetoTheme]::Rgb('#E91E63')
                            }
                            $mutated=$true
                        }
                    }
                }
            }
            if ($variant -eq 'pink_chart_only') {
                foreach ($slide in $presentation.Slides) {
                    foreach ($shape in $slide.Shapes) {
                        if (-not $mutated -and $shape.HasChart -eq -1) {
                            $shape.Chart.SeriesCollection().Item(1).Format.Fill.ForeColor.RGB=[Scribble.Office.MetoTheme]::Rgb('#E91E63')
                            $mutated=$true
                        }
                    }
                }
            }
            if ($variant.StartsWith('pink_') -and -not $mutated) { throw "Unable to construct targeted $variant mutant" }
            if ($mutated) { $presentation.Save() }
            $native=[Scribble.Testing.TestLabStressEvidence]::Read($presentation,'PowerPoint',[string[]]@(),[int[]]@())
            $runId=[Guid]::NewGuid().ToString('N')
            $zipPath=Join-Path $preflightRoot ($variant+'-native-evidence.zip')
            $checks=[NativeGradingPreflightHelper]::Evaluate($projection,$projectedCase,$native,$runId,$zipPath)
            $presentationChecks=@($checks | Where-Object { $_.name -match '_presentation$' })
            if ($presentationChecks.Count -ne 1) { throw "Native preflight did not reach the presentation checker: $($checks | ConvertTo-Json -Compress)" }
            $accepted=@($checks | Where-Object { $_.hard -and -not $_.passed }).Count -eq 0
            $expectedAcceptance=$variant -eq 'clean_reference' -or $variant -eq 'candidate'
            $records.Add([pscustomobject]@{variant=$variant;expected_accepted=$expectedAcceptance;accepted=$accepted;passed=($accepted -eq $expectedAcceptance);checks=$checks;evidence_zip=$zipPath;native_measurement_error=$native.error;owned_copy=$owned})
            Write-Host "$variant accepted=$accepted expected=$expectedAcceptance"
        } finally {
            if ($null -ne $presentation) { $presentation.Close(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
        }
    }
} finally {
    if ($quit) { $app.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
if ($originalHashes.reference -ne (Get-FileHash -LiteralPath $reference -Algorithm SHA256).Hash -or $originalHashes.source -ne (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash) { throw 'An original corpus presentation changed' }
$report=Join-Path $preflightRoot 'native-presentation-grading-preflight.json'
Write-Json $report ([pscustomobject]@{schema_version=1;scope='Native presentation-grader only, zero model calls, no TestLab session mutation; clean_reference means normalized owned baseline, not byte-identical source';case_id=$CaseId;assembly_sha256=(Get-FileHash -LiteralPath $assemblyPath -Algorithm SHA256).Hash.ToLowerInvariant();parent_manifest_sha256=$parentHash;candidate_sha256=if ($CandidatePptx) { (Get-FileHash -LiteralPath $CandidatePptx -Algorithm SHA256).Hash.ToLowerInvariant() } else { $null };original_reference_sha256=$originalHashes.reference.ToLowerInvariant();normalized_baseline_sha256=(Get-FileHash -LiteralPath $baseline -Algorithm SHA256).Hash.ToLowerInvariant();baseline_normalizations=@($normalizations.ToArray());original_oracle_sha256=$derivedOracle.original_oracle_sha256;derived_oracle_sha256=(Get-FileHash -LiteralPath (Join-Path $projection $projectedCase.oracle_ref) -Algorithm SHA256).Hash.ToLowerInvariant();projection_manifest_sha256=(Get-FileHash -LiteralPath (Join-Path $projection 'manifest.json') -Algorithm SHA256).Hash.ToLowerInvariant();results=@($records.ToArray())})
Write-Output $report
if (@($records | Where-Object { -not $_.passed }).Count -gt 0) { throw "Native presentation grading preflight failed; see $report" }
