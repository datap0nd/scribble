#requires -Version 5.1
param([string]$KitRoot=(Join-Path $PSScriptRoot '..\generated\scribble-test-kit-v1'))
$ErrorActionPreference='Stop'
$kit=(Resolve-Path -LiteralPath $KitRoot).Path
$qa=Join-Path (Split-Path $kit -Parent) 'qa'
New-Item -ItemType Directory -Path $qa -Force | Out-Null
$checks=@()
# Only fixture documents opened by this script are closed. No production documents are touched.
$excel=New-Object -ComObject Excel.Application
$excel.DisplayAlerts=$false
try {
    foreach($file in Get-ChildItem -LiteralPath (Join-Path $kit 'inputs\excel') -Filter *.xlsx) {
        $book=$excel.Workbooks.Open($file.FullName,0,$true)
        try { $checks+=@{kind='xlsx';file=$file.Name;sheets=$book.Worksheets.Count;opened=$true} } finally { $book.Close($false) }
    }
    $book=$excel.Workbooks.Open((Join-Path $kit 'evaluator-only\expected-analysis.xlsx'),0,$false)
    try {
        $excel.CalculateFullRebuild()
        $sheet=$book.Worksheets.Item('Scribble Draft')
        if ([double]$sheet.Range('C2').Value2 -ne 120000 -or [double]$sheet.Range('C4').Value2 -ne 46000) { throw 'Native formula recalculation mismatch.' }
        $checks+=@{kind='reference_xlsx';revenue=$sheet.Range('C2').Value2;gross_profit=$sheet.Range('C4').Value2;margin=$sheet.Range('C5').Value2;charts=$sheet.ChartObjects().Count;recalculated=$true}
        $book.Save()
    } finally { $book.Close($false) }
} finally { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
$ppt=New-Object -ComObject PowerPoint.Application
try {
    foreach($path in @((Join-Path $kit 'evaluator-only\reference-deck.pptx'),(Join-Path $kit 'inputs\powerpoint\Atlas-start.pptx'),(Join-Path $kit 'inputs\powerpoint\Atlas-crowded.pptx'))) {
        $deck=$ppt.Presentations.Open($path,$true,$false,$false)
        try {
            $checks+=@{kind='pptx';file=[IO.Path]::GetFileName($path);slides=$deck.Slides.Count;opened=$true}
            if($path.EndsWith('reference-deck.pptx')) {
                $deck.SaveCopyAs((Join-Path $kit 'evaluator-only\reference-deck.pdf'),32)
                for($i=1;$i -le $deck.Slides.Count;$i++) { $deck.Slides.Item($i).Export((Join-Path $kit "evaluator-only\reference-deck-slides\slide-$i.png"),'PNG',1280,720) }
            }
        } finally { $deck.Close() }
    }
} finally { if($ppt.Presentations.Count -eq 0) { $ppt.Quit() }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt) }
$word=New-Object -ComObject Word.Application
$word.Visible=$false
try {
    foreach($file in @((Get-ChildItem (Join-Path $kit 'inputs\word') -Filter *.docx)) + @((Get-Item (Join-Path $kit 'evaluator-only\reference-summary.docx')))) {
        $doc=$word.Documents.Open($file.FullName,$false,$true)
        try { $doc.ExportAsFixedFormat((Join-Path $qa ($file.BaseName+'.pdf')),17); $checks+=@{kind='docx';file=$file.Name;pages=$doc.ComputeStatistics(2);opened=$true} }
        finally { $doc.Close(0) }
    }
} finally { $word.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) }
$checks | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $qa 'native-fixture-checks.json') -Encoding utf8
$checks | ConvertTo-Json -Depth 5
