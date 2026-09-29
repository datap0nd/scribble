param(
    [Parameter(Mandatory=$true)][string]$SourcePresentation,
    [Parameter(Mandatory=$true)][string]$AssemblyPath,
    [Parameter(Mandatory=$true)][string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
$lock = 'C:\ScribbleRuns\MACHINE_LOCK.txt'
if (Test-Path -LiteralPath $lock) { throw 'Machine lock held.' }
$baseline = @(Get-Process EXCEL,POWERPNT,WINWORD -ErrorAction SilentlyContinue |
    Select-Object ProcessName,Id,StartTime)
if ($baseline.Count) { throw 'Office baseline is not empty.' }
[IO.Directory]::CreateDirectory($OutputDirectory) | Out-Null
$sourcePath = (Resolve-Path -LiteralPath $SourcePresentation).Path
$assembly = (Resolve-Path -LiteralPath $AssemblyPath).Path
$copyPath = Join-Path $OutputDirectory ('source-' + [Guid]::NewGuid().ToString('N') + '.pptx')
Copy-Item -LiteralPath $sourcePath -Destination $copyPath
$sourceHash = (Get-FileHash -LiteralPath $copyPath -Algorithm SHA256).Hash
@('lane=P1 native copy receipt regression',
  ('session=' + [Guid]::NewGuid().ToString('N')),
  ('start_utc=' + [DateTime]::UtcNow.ToString('o'))) |
    Set-Content -LiteralPath $lock
$app = $null; $source = $null; $draftCopy = $null
$stage = 'open'
try {
    Add-Type -Path $assembly
    $app = New-Object -ComObject PowerPoint.Application
    $app.Visible = -1
    $source = $app.Presentations.Open($copyPath, -1, 0, -1)
    $source.Windows.Item(1).Activate()
    $copyType = [Scribble.Office.PresentationInspection].Assembly.GetType(
        'Scribble.Office.PresentationDraftCopy', $true)
    $revisionType = [Scribble.Office.PresentationInspection].Assembly.GetType(
        'Scribble.Office.PresentationRevision', $true)
    $static = [Reflection.BindingFlags]'Static,NonPublic'
    $instance = [Reflection.BindingFlags]'Instance,NonPublic'
    $measured = [int[]]$copyType.GetMethod('MeasuredReplacementSlides', $static).
        Invoke($null, [object[]]@([object]$source))
    if ($measured.Length -eq 0) { throw 'No measured replacement slide.' }
    $target = [Scribble.Office.PresentationInspection]::FindSlide($source, $measured[0])
    $chartBystander = $false
    for ($i = 1; $i -le $source.Slides.Count; $i++) {
        $page = $source.Slides.Item($i)
        if ($measured -contains [int]$page.SlideID) { continue }
        foreach ($shape in $page.Shapes) {
            if ($shape.HasChart -ne 0) { $chartBystander = $true; break }
        }
    }
    if (-not $chartBystander) { throw 'No native chart bystander.' }
    $stage = 'copy'
    $draftCopy = $copyType.GetMethod('Create', $static).Invoke($null,
        [object[]]@([object]$app, [object]$source,
            [Guid]::NewGuid().ToString('N')))
    $operation = [Collections.Generic.Dictionary[string,object]]::new()
    $operation['kind'] = 'replace_slide'
    $operation['slide_id'] = [int]$target.SlideID
    $operation['fingerprint'] = [Scribble.Office.PresentationInspection]::Fingerprint($target)
    $operation['notes'] = 'Source note retained.'
    $slide = [Collections.Generic.Dictionary[string,object]]::new()
    $slide['layout'] = 'bullets'
    $slide['title'] = 'Operating review and evidence boundaries'
    $slide['subtitle'] = 'Source backed measures and interpretation'
    $slide['bullets'] = [string[]]@(
        'Current and prior periods must be distinguished.',
        'Incomplete source figures remain unknown.',
        'Planned actions do not establish completed results.')
    $slide['sources'] = 'Workbook source'
    $slide['footnote'] = 'Workbook source'
    $slide['evidence'] = 'Source evidence'
    $operation['slide'] = $slide
    $operations = [object[]]::new(1); $operations[0] = $operation
    $one = [object[]]::new(1); $one[0] = $operations
    $stage = 'bind'
    $bound = $copyType.GetMethod('BindOperations', $instance).Invoke($draftCopy, $one)
    $one[0] = $measured
    $style = $copyType.GetMethod('MeasuredNativeStyleOperations', $instance).
        Invoke($draftCopy, $one)
    $combined = [object[]]@($bound) + [object[]]@($style)
    $draft = $copyType.GetField('Draft',
        [Reflection.BindingFlags]'Instance,Public,NonPublic').GetValue($draftCopy)
    $draft.Windows.Item(1).Activate()
    $two = [object[]]::new(2); $two[0] = $draft; $two[1] = $combined
    $stage = 'patch'
    $changed = $revisionType.GetMethod('ApplyOwnedDraft', $static).Invoke($null, $two)
    $stage = 'receipt'
    $two[0] = $changed
    $copyType.GetMethod('AcceptDirectRevision', $instance).Invoke($draftCopy, $two) | Out-Null
    if ((Get-FileHash -LiteralPath $copyPath -Algorithm SHA256).Hash -ne $sourceHash)
    { throw 'Source bytes changed.' }
    [pscustomobject]@{status='passed';assembly_sha256=(Get-FileHash -LiteralPath $assembly -Algorithm SHA256).Hash;
        chart_bystander=$chartBystander;changed_slides=$changed.Count;source_preserved=$true;
        paid_model_calls=0} | ConvertTo-Json -Compress |
        Set-Content -LiteralPath (Join-Path $OutputDirectory 'receipt.json')
    Get-Content -LiteralPath (Join-Path $OutputDirectory 'receipt.json')
}
catch {
    [pscustomobject]@{status='failed';stage=$stage;error=$_.Exception.Message;
        source_preserved=((Get-FileHash -LiteralPath $copyPath -Algorithm SHA256).Hash -eq $sourceHash);
        paid_model_calls=0} | ConvertTo-Json -Compress |
        Set-Content -LiteralPath (Join-Path $OutputDirectory 'receipt.json')
    throw
}
finally {
    if ($null -ne $draftCopy) {
        try { $copyType.GetMethod('DiscardOwnedDraft', $instance).
            Invoke($draftCopy, @()) | Out-Null } catch { }
    }
    if ($null -ne $source) { try { $source.Close() } catch { } }
    if ($null -ne $app) {
        try {
            for ($i = $app.Presentations.Count; $i -ge 1; $i--) {
                $app.Presentations.Item($i).Close()
            }
            $app.Quit()
        } catch { }
    }
    if (Test-Path -LiteralPath $copyPath) { Remove-Item -LiteralPath $copyPath }
    Remove-Item -LiteralPath $lock -Force
}
