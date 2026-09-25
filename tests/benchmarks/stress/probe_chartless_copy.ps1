param(
    [Parameter(Mandatory=$true)][string]$SourcePptx,
    [string]$AssemblyPath,
    [string]$CaptureDirectory
)

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
    $copy.FollowMasterBackground = $page.FollowMasterBackground
    if ($page.Background.Fill.Type -eq 1) {
        $copy.Background.Fill.Solid()
        $copy.Background.Fill.ForeColor.RGB = $page.Background.Fill.ForeColor.RGB
        $copy.Background.Fill.Transparency = $page.Background.Fill.Transparency
    }
    $copy.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text =
        $page.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text
    $copy.NotesPage.Shapes.Item(3).TextFrame.TextRange.Text =
        $page.NotesPage.Shapes.Item(3).TextFrame.TextRange.Text
    for ($i = 1; $i -le $page.Shapes.Count; $i++) {
        $shape = $page.Shapes.Item($i)
        if ($shape.HasChart -ne 0) { continue }
        $stage = "copy_shape_$i"
        $shape.Copy()
        $null = $copy.Shapes.Paste()
        $copy.Shapes.Item($copy.Shapes.Count).Name = $shape.Name
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
    foreach ($label in @('source', 'copy')) {
        $slide = if ($label -eq 'source') { $page } else { $copy }
        Write-Output "notes $label slide-number-visible=$($slide.NotesPage.HeadersFooters.SlideNumber.Visible)"
        for ($i = 1; $i -le $slide.NotesPage.Shapes.Count; $i++) {
            $notesShape = $slide.NotesPage.Shapes.Item($i)
            if ($notesShape.HasTextFrame -ne 0) {
                Write-Output ("notes {0} {1}: name={2}; text={3}" -f
                    $label, $i, $notesShape.Name,
                    $notesShape.TextFrame.TextRange.Text)
            }
        }
    }
    if ($AssemblyPath -and $CaptureDirectory) {
        Add-Type -AssemblyName System.Web.Extensions
        $assembly = [System.Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath).Path)
        $type = $assembly.GetType('Scribble.Office.PresentationInspection', $true)
        $flags = [System.Reflection.BindingFlags]::NonPublic -bor [System.Reflection.BindingFlags]::Static
        $method = $type.GetMethods($flags) | Where-Object {
            $_.Name -eq 'Capture' -and $_.GetParameters().Count -eq 2
        } | Select-Object -First 1
        $serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        $serializer.MaxJsonLength = [int]::MaxValue
        [System.IO.Directory]::CreateDirectory($CaptureDirectory) | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $CaptureDirectory 'source.json'),
            $serializer.Serialize($method.Invoke($null, @($page, $false))))
        [System.IO.File]::WriteAllText((Join-Path $CaptureDirectory 'copy.json'),
            $serializer.Serialize($method.Invoke($null, @($copy, $false))))
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
