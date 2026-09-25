param([Parameter(Mandatory=$true)][string]$SourcePptx)

$ErrorActionPreference = 'Stop'
$app = $null
$source = $null
$draft = $null
$stage = 'start'
try {
    $app = New-Object -ComObject PowerPoint.Application
    $app.Visible = -1
    $stage = 'open_source'
    $source = $app.Presentations.Open($SourcePptx, -1, 0, 0)
    Write-Output "source saved before draft=$($source.Saved), path=$($source.Path)"
    $stage = 'new_draft'
    $draft = $app.Presentations.Add(-1)
    $draft.PageSetup.SlideWidth = $source.PageSetup.SlideWidth
    $draft.PageSetup.SlideHeight = $source.PageSetup.SlideHeight
    Write-Output "source saved after draft=$($source.Saved)"
    $page = $source.Slides.Item(2)
    $copy = $draft.Slides.Add(1, 12)
    for ($i = 1; $i -le $page.Shapes.Count; $i++) {
        $shape = $page.Shapes.Item($i)
        if ($shape.HasChart -ne 0) { continue }
        $stage = "copy_shape_$i"
        $shape.Copy()
        $null = $copy.Shapes.Paste()
    }
    $stage = 'verify_chartless_draft'
    if ($copy.Shapes.Count -ne ($page.Shapes.Count - 1)) {
        throw 'The new slide has an unexpected number of shapes.'
    }
    for ($i = 1; $i -le $copy.Shapes.Count; $i++) {
        $original = $page.Shapes.Item($i)
        $cloned = $copy.Shapes.Item($i)
        Write-Output ("shape {0}: {1}/{2} => {3}/{4}; x={5}/{6}; z={7}/{8}" -f
            $i, $original.Name, $original.Type, $cloned.Name, $cloned.Type,
            $original.Left, $cloned.Left,
            $original.ZOrderPosition, $cloned.ZOrderPosition)
    }
    Write-Output "CHARTLESS_SHAPE_COPY_OK shapes=$($copy.Shapes.Count)"
}
catch {
    Write-Output "CHARTLESS_SHAPE_COPY_FAILED at $stage : $_"
    exit 1
}
finally {
    if ($null -ne $draft) { try { $draft.Close() } catch { } }
    if ($null -ne $source) { try { $source.Close() } catch { } }
    if ($null -ne $app) { try { if ($app.Presentations.Count -eq 0) { $app.Quit() } } catch { } }
}
