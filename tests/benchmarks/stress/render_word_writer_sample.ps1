param(
    [Parameter(Mandatory=$true)][string]$AssemblyPath,
    [Parameter(Mandatory=$true)][string]$OutputPdf
)

$ErrorActionPreference = 'Stop'
$word = $null
$draft = $null
try {
    $assembly = [System.Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $AssemblyPath).Path)
    $writer = $assembly.GetType('Scribble.Office.WordDraftWriter', $true)
    $flags = [System.Reflection.BindingFlags]::NonPublic -bor [System.Reflection.BindingFlags]::Static
    $method = $writer.GetMethods($flags) | Where-Object {
        $_.Name -eq 'WriteDraftDocument' -and $_.GetParameters().Count -eq 4
    } | Select-Object -First 1
    if ($null -eq $method) { throw 'Scribble Word draft writer was not found.' }

    $body = @'
## Executive summary
The source ledger was read into a separate analysis draft. Revenue and margin trends should be verified against the workbook before any business decision. The source remains unchanged.

## Review snapshot
| Check | Draft result |
| Source workbook | Read without writing |
| Analysis workbook | New draft created |
| Presentation | Four-slide draft created |
| Human review | Still required |

## Next actions
- Confirm the totals and date labels against the source ledger.
- Review every slide and page before sharing the output.

**Status:** Offline Scribble writer sample. No hosted model was called.
'@
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $status = $method.Invoke($null, @($word, 'Scribble 2.0 | Analysis brief', $body, 'new_document'))
    $draft = $word.ActiveDocument
    if ($null -eq $draft) { throw 'Scribble did not create a Word draft.' }
    $parent = Split-Path -Parent $OutputPdf
    [System.IO.Directory]::CreateDirectory($parent) | Out-Null
    $draft.ExportAsFixedFormat($OutputPdf, 17)
    Write-Output $status
    Write-Output $OutputPdf
}
finally {
    if ($null -ne $draft) {
        try { $draft.Close(0) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($draft) } catch { }
    }
    if ($null -ne $word) {
        try { $word.Quit(0) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) } catch { }
    }
}
