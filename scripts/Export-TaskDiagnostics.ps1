param([Parameter(Mandatory=$true)][string]$TaskId,
      [string]$AssemblyPath = (Join-Path $PSScriptRoot '..\src\Scribble\bin\Release\Scribble.dll'),
      [switch]$LocalReplay)
$ErrorActionPreference = 'Stop'
Add-Type -Path (Resolve-Path -LiteralPath $AssemblyPath)
$store = [Scribble.Chat.TaskCheckpointStore]::new($null)
$state = $store.Load($TaskId)
$trace = [Scribble.Chat.TaskDiagnostics]::new($store, $state)
if ($LocalReplay) {
    # Private local inspection only. Does not execute recorded tools or inference.
    $trace.ReadLocalReplay() | ConvertTo-Json -Depth 12
} else {
    # Preview this metadata-only report before sharing it.
    $trace.RedactedReport()
}
