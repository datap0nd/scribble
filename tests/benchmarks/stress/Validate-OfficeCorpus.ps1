[CmdletBinding()]
param(
    [string]$CorpusRoot = '',
    [string[]]$Only = @(),
    [switch]$SkipExcel,
    [switch]$SkipPowerPoint
)

# Validate generated sources, not model output. Open only corpus-owned copies;
# never attach to or close the user's workbooks/presentations. No chart external
# data activation, refresh, link update, or mailbox action is performed.
$ErrorActionPreference = 'Stop'
if (-not $CorpusRoot) { $CorpusRoot = Join-Path $PSScriptRoot '../generated/stress-corpus' }
$root = [IO.Path]::GetFullPath($CorpusRoot)
$catalog = Get-Content -LiteralPath (Join-Path $root 'evaluator-only/office_catalog.json') -Raw | ConvertFrom-Json
$runRoot = Join-Path $root ('evaluator-only/native-validation/' + [Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($runRoot) | Out-Null
$results = [Collections.Generic.List[object]]::new()
$errors = [Collections.Generic.List[string]]::new()
$reportPath = Join-Path $runRoot 'native-office-validation.json'

function Save-Progress {
    [pscustomobject]@{
        schema_version = 1
        scope = 'Native Office source corpus validation; zero model calls; does not prove Qwen output quality'
        source_catalog = 'evaluator-only/office_catalog.json'
        results = @($script:results.ToArray())
        failures = @($script:errors.ToArray())
    } | ConvertTo-Json -Depth 24 | Set-Content -LiteralPath $script:reportPath -Encoding UTF8
}

function Release-Com($value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)
    }
}

function Assert-Numeric($Actual, $Expected, [string]$Label, [double]$Tolerance = 0.000001) {
    if ($null -eq $Expected) {
        if ($null -ne $Actual -and [string]$Actual -ne '') { throw "$Label must be blank/unknown, observed $Actual" }
        return
    }
    if ($null -eq $Actual -or $Actual -is [string] -or [math]::Abs([double]$Actual - [double]$Expected) -gt $Tolerance) {
        throw "$Label expected $Expected, observed $Actual"
    }
}

function Cell-Indices([string]$Address) {
    if ($Address -notmatch '^([A-Z]+)([0-9]+)$') { throw "Unsupported cell address $Address" }
    $column = 0
    foreach ($letter in $Matches[1].ToCharArray()) { $column = $column * 26 + [int]$letter - 64 }
    return @([int]$Matches[2], $column)
}

function Native-Array($Value) {
    if ($Value -is [Array]) { return @($Value | ForEach-Object { $_ }) }
    return @($Value)
}

function Hex-Color([int]$Color) {
    return '{0:X2}{1:X2}{2:X2}' -f ($Color -band 255), (($Color -shr 8) -band 255), (($Color -shr 16) -band 255)
}

$excel = $null
$powerPoint = $null
$quitExcel = $false
$quitPowerPoint = $false
try {
    if (-not $SkipExcel) {
        $excel = New-Object -ComObject Excel.Application
        $quitExcel = $excel.Workbooks.Count -eq 0
        if (-not $quitExcel) { throw 'Excel did not create an isolated empty instance; existing workbooks will not be changed' }
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.AutomationSecurity = 3
        foreach ($spec in $catalog.workbooks) {
            if ($Only.Count -gt 0 -and $spec.id -notin $Only) { continue }
            $workbook = $null
            try {
                $sourcePath = Join-Path $root $spec.path
                $beforeHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
                $copy = Join-Path $runRoot ($spec.id + '.xlsx')
                Copy-Item -LiteralPath $sourcePath -Destination $copy
                $workbook = $excel.Workbooks.Open($copy, 0, $false)
                $ledger = $workbook.Worksheets.Item('Ledger')
                $history = $workbook.Worksheets.Item('History')
                $ledger.Calculate()
                $history.Calculate()
                $used = $ledger.UsedRange
                $values = $used.Value2
                $formulas = $used.Formula
                foreach ($expected in $spec.native_formula_cells) {
                    $indices = Cell-Indices $expected.cell
                    $actual = $values.GetValue($indices[0],$indices[1])
                    $formula = [string]$formulas.GetValue($indices[0],$indices[1])
                    if (-not $formula.StartsWith('=')) { throw "$($spec.id) $($expected.cell) lost its formula" }
                    Assert-Numeric $actual $expected.expected "$($spec.id) $($expected.cell)" $expected.tolerance
                }
                foreach ($cell in $spec.facts.missing_cells) {
                    $indices = Cell-Indices $cell.cell
                    Assert-Numeric ($values.GetValue($indices[0],$indices[1])) $null "$($spec.id) missing $($cell.cell)"
                }
                $historyValues = $history.Range('B2:C6').Value2
                for ($month=0;$month -lt 5;$month++) {
                    Assert-Numeric ($historyValues.GetValue($month+1,1)) $spec.facts.monthly[$month].primary "$($spec.id) history primary month $month"
                    Assert-Numeric ($historyValues.GetValue($month+1,2)) $spec.facts.monthly[$month].secondary "$($spec.id) history secondary month $month"
                }
                $chartObjects = $history.ChartObjects()
                if ($chartObjects.Count -lt 1) { throw "$($spec.id) is missing its native chart" }
                $seriesCollection = $chartObjects.Item(1).Chart.SeriesCollection()
                $seriesEvidence = @()
                for ($seriesIndex=1;$seriesIndex -le $seriesCollection.Count;$seriesIndex++) {
                    $series = $seriesCollection.Item($seriesIndex)
                    $seriesValues = Native-Array $series.Values
                    $categories = Native-Array $series.XValues
                    if ($seriesValues.Count -ne 5 -or $categories.Count -ne 5) { throw "$($spec.id) history chart category/value count mismatch" }
                    for ($month=0;$month -lt 5;$month++) {
                        $expected = if ($seriesIndex -eq 1) { $spec.facts.monthly[$month].primary } else { $spec.facts.monthly[$month].secondary }
                        Assert-Numeric $seriesValues[$month] $expected "$($spec.id) chart series $seriesIndex month $month"
                        if ([string]$categories[$month] -ne [string]$spec.facts.monthly[$month].period) { throw "$($spec.id) native history chart period label differs" }
                    }
                    $seriesEvidence += [pscustomobject]@{ name=$series.Name; formula=$series.Formula; categories=$categories; values=$seriesValues }
                    Release-Com $series
                }
                $probes = @()
                foreach ($probe in $spec.formula_probes) {
                    $target = $ledger.Range($probe.cell)
                    $target.Value2 = [double]$probe.new_value
                    $ledger.Calculate()
                    foreach ($dependent in $probe.dependent_cells) {
                        Assert-Numeric ($ledger.Range($dependent.cell).Value2) $dependent.expected "$($spec.id) changed $($dependent.cell)"
                    }
                    $target.Value2 = [double]$probe.original_value
                    $ledger.Calculate()
                    $probes += [pscustomobject]@{ cell=$probe.cell; delta=$probe.delta; dependent_cells=@($probe.dependent_cells); passed=$true }
                    Release-Com $target
                }
                $workbook.Close($false)
                Release-Com $workbook
                $workbook = $null
                if ($beforeHash -ne (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash) { throw "$($spec.id) source changed" }
                $results.Add([pscustomobject]@{id=$spec.id;kind='excel';passed=$true;native_formula_cells=$spec.native_formula_cells.Count;missing_values=$spec.facts.missing_cells.Count;history_chart=$seriesEvidence;dependency_probes=$probes;source_sha256=$beforeHash})
                Write-Host "$($spec.id) native formula/cache/chart/dependency checks passed"
            } catch {
                $message = "$($spec.id): $($_.Exception.Message)"
                $errors.Add($message)
                $results.Add([pscustomobject]@{id=$spec.id;kind='excel';passed=$false;error=$message})
                Write-Warning $message
            } finally {
                if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} ; Release-Com $workbook }
                Save-Progress
            }
        }
    }
    if (-not $SkipPowerPoint) {
        $existingPowerPoint = @(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $quitPowerPoint = (-not $existingPowerPoint) -and $powerPoint.Presentations.Count -eq 0
        foreach ($spec in $catalog.presentations) {
            if ($Only.Count -gt 0 -and $spec.id -notin $Only) { continue }
            foreach ($variant in @('source','reference')) {
                $presentation = $null
                try {
                    $relative = if ($variant -eq 'source') { $spec.path } else { $spec.reference_path }
                    $sourcePath = Join-Path $root $relative
                    $beforeHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
                    $presentation = $powerPoint.Presentations.Open($sourcePath, -1, 0, 0)
                    Assert-Numeric $presentation.PageSetup.SlideWidth 960 "$($spec.id) canvas width" 0.01
                    Assert-Numeric $presentation.PageSetup.SlideHeight 540 "$($spec.id) canvas height" 0.01
                    if ($presentation.Slides.Count -ne $spec.slide_count) { throw "Slide count differs from source contract" }
                    $geometry = @()
                    $charts = @()
                    $fonts = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                    $headers = @()
                    $pngDir = Join-Path $runRoot ($spec.id + '-' + $variant)
                    [IO.Directory]::CreateDirectory($pngDir) | Out-Null
                    for ($slideIndex=1;$slideIndex -le $presentation.Slides.Count;$slideIndex++) {
                        $slide = $presentation.Slides.Item($slideIndex)
                        $slide.Export((Join-Path $pngDir ('slide-{0:D2}.png' -f $slideIndex)), 'PNG', 1280, 720)
                        for ($shapeIndex=1;$shapeIndex -le $slide.Shapes.Count;$shapeIndex++) {
                            $shape = $slide.Shapes.Item($shapeIndex)
                            $outside = $shape.Left -lt -0.5 -or $shape.Top -lt -0.5 -or ($shape.Left+$shape.Width) -gt 960.5 -or ($shape.Top+$shape.Height) -gt 540.5
                            if ($outside) { $geometry += [pscustomobject]@{slide=$slideIndex;shape=$shapeIndex;kind='out_of_bounds'} }
                            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                                $text = $shape.TextFrame.TextRange
                                [void]$fonts.Add([string]$text.Font.Name)
                                if ($text.BoundHeight -gt ($shape.Height - $shape.TextFrame.MarginTop - $shape.TextFrame.MarginBottom + 1.5)) {
                                    $geometry += [pscustomobject]@{slide=$slideIndex;shape=$shapeIndex;kind='text_overflow';bound_height=$text.BoundHeight;shape_height=$shape.Height;text=([string]$text.Text).Substring(0,[Math]::Min(70,([string]$text.Text).Length))}
                                }
                            }
                            if ($shape.HasTable -eq -1) {
                                $table=$shape.Table
                                $headers += [pscustomobject]@{slide=$slideIndex;color=(Hex-Color $table.Cell(1,1).Shape.Fill.ForeColor.RGB)}
                                for ($r=1;$r -le $table.Rows.Count;$r++) { for ($c=1;$c -le $table.Columns.Count;$c++) {
                                    $cellShape=$table.Cell($r,$c).Shape
                                    $cellText=$cellShape.TextFrame.TextRange
                                    [void]$fonts.Add([string]$cellText.Font.Name)
                                    if ($cellText.BoundHeight -gt ($cellShape.Height-$cellShape.TextFrame.MarginTop-$cellShape.TextFrame.MarginBottom+1.5)) {
                                        $geometry += [pscustomobject]@{slide=$slideIndex;shape=$shapeIndex;kind='table_text_overflow';row=$r;column=$c;bound_height=$cellText.BoundHeight;shape_height=$cellShape.Height}
                                    }
                                } }
                            }
                            if ($shape.HasChart -eq -1) {
                                $seriesCollection=$shape.Chart.SeriesCollection()
                                $seriesList=@()
                                for ($si=1;$si -le $seriesCollection.Count;$si++) {
                                    $series=$seriesCollection.Item($si)
                                    $color=if ($slideIndex -eq 10) { Hex-Color $series.Format.Line.ForeColor.RGB } else { Hex-Color $series.Format.Fill.ForeColor.RGB }
                                    $expectedColor=if ($si -eq 1) { '4F81BD' } else { '5B9BD5' }
                                    if ($color -ne $expectedColor) { throw "Chart series $si on slide $slideIndex has unexpected native color $color" }
                                    $seriesList += [pscustomobject]@{name=$series.Name;values=(Native-Array $series.Values);categories=(Native-Array $series.XValues);color=$color}
                                    Release-Com $series
                                }
                                $charts += [pscustomobject]@{slide=$slideIndex;series=$seriesList}
                            }
                            Release-Com $shape
                        }
                        Release-Com $slide
                    }
                    if ($charts.Count -ne $spec.native_chart_slides.Count -or $headers.Count -ne $spec.native_table_slides.Count) { throw "Native chart/table counts differ" }
                    $mainChart=@($charts | Where-Object slide -eq 2)[0]
                    for ($si=0;$si -lt $spec.chart_specs[0].series.Count;$si++) {
                        $expectedSeries=$spec.chart_specs[0].series[$si]
                        $actualSeries=$mainChart.series[$si]
                        if ($actualSeries.name -ne $expectedSeries.name -or $actualSeries.values.Count -ne 6) { throw "Native monthly chart series identity/count differs" }
                        for ($m=0;$m -lt 6;$m++) {
                            Assert-Numeric $actualSeries.values[$m] $expectedSeries.values[$m] "$($spec.id) slide2 series $si month $m"
                            if ([string]$actualSeries.categories[$m] -ne [string]$spec.chart_specs[0].categories[$m]) { throw "Native chart month labels differ" }
                        }
                    }
                    foreach ($chart in $charts) {
                        $rows=if ($chart.slide -eq 7) { @($spec.facts.by_group) } else { @($spec.facts.monthly) }
                        for ($si=0;$si -lt $chart.series.Count;$si++) {
                            $series=$chart.series[$si]
                            if ($series.values.Count -ne $rows.Count -or $series.categories.Count -ne $rows.Count) { throw "Native chart shape mismatch on slide $($chart.slide)" }
                            for ($i=0;$i -lt $rows.Count;$i++) {
                                $expected=if ($si -eq 0) { $rows[$i].primary } else { $rows[$i].secondary }
                                $category=if ($chart.slide -eq 7) { $rows[$i].group } else { $rows[$i].period }
                                Assert-Numeric $series.values[$i] $expected "$($spec.id) slide$($chart.slide) series$si point$i"
                                if ([string]$series.categories[$i] -ne [string]$category) { throw "Native chart category mismatch on slide $($chart.slide)" }
                            }
                        }
                    }
                    $expectedDefects=$variant -eq 'source' -and $spec.intentional_defects.Count -gt 0
                    $unintended=@($geometry | Where-Object { -not ($expectedDefects -and (($_.slide -eq 2 -and $_.kind -eq 'out_of_bounds') -or ($_.slide -eq 4 -and $_.kind -eq 'text_overflow'))) })
                    if ($unintended.Count -gt 0) { throw ('Unintended native geometry: '+($unintended | ConvertTo-Json -Compress)) }
                    if ($expectedDefects -and (@($geometry | Where-Object { $_.slide -eq 2 -and $_.kind -eq 'out_of_bounds' }).Count -ne 1 -or @($geometry | Where-Object { $_.slide -eq 4 -and $_.kind -eq 'text_overflow' }).Count -ne 1)) { throw "Intended repair geometry was not reproduced" }
                    foreach ($header in $headers) {
                        $expectedColor=if ($expectedDefects -and $header.slide -eq 3) { 'E91E63' } else { '4F81BD' }
                        if ($header.color -ne $expectedColor) { throw "Unexpected native table theme color $($header.color) on slide $($header.slide)" }
                    }
                    if (@($fonts | Where-Object { $_ -and $_ -ne 'Arial' }).Count -gt 0) { throw ('Unexpected native fonts: '+($fonts -join ',')) }
                    $presentation.Close()
                    Release-Com $presentation
                    $presentation=$null
                    if ($beforeHash -ne (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash) { throw "Source presentation changed" }
                    $results.Add([pscustomobject]@{id=$spec.id;kind='powerpoint';variant=$variant;passed=$true;slides=$spec.slide_count;charts=$charts;table_headers=$headers;fonts=@($fonts);geometry=$geometry;native_png_directory=$pngDir;source_sha256=$beforeHash})
                    Write-Host "$($spec.id) $variant native render/chart/theme/geometry checks passed"
                } catch {
                    $message="$($spec.id) $variant`: $($_.Exception.Message)"
                    $errors.Add($message)
                    $results.Add([pscustomobject]@{id=$spec.id;kind='powerpoint';variant=$variant;passed=$false;error=$message})
                    Write-Warning $message
                } finally {
                    if ($null -ne $presentation) { try { $presentation.Close() } catch {}; Release-Com $presentation }
                    Save-Progress
                }
            }
        }
    }
} finally {
    if ($null -ne $excel) { if ($quitExcel) { try { $excel.Quit() } catch {} }; Release-Com $excel }
    if ($null -ne $powerPoint) { if ($quitPowerPoint) { try { $powerPoint.Quit() } catch {} }; Release-Com $powerPoint }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Save-Progress
}
Write-Output $reportPath
if ($errors.Count -gt 0) { throw "Native corpus validation has $($errors.Count) failure(s); see $reportPath" }
