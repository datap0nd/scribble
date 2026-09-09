#requires -Version 5.1
param([string]$AssemblyPath)
$ErrorActionPreference = 'Stop'
if (-not $AssemblyPath) {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Programs\Scribble\Scribble.dll'),
        (Join-Path $env:LOCALAPPDATA 'Scribble\Scribble.dll'),
        (Join-Path $PSScriptRoot '..\..\..\src\Scribble\bin\Release\Scribble.dll'))
    $AssemblyPath = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
}
if (-not $AssemblyPath -or -not (Test-Path -LiteralPath $AssemblyPath)) { throw 'Pass -AssemblyPath pointing to the installed Scribble.dll from the Test Lab build.' }
$resolvedAssembly = (Resolve-Path -LiteralPath $AssemblyPath).Path
Add-Type -Path $resolvedAssembly
$labType = 'Scribble.Testing.TestLab' -as [type]
if ($null -eq $labType) {
    $installedVersion = (Get-Item -LiteralPath $resolvedAssembly).VersionInfo.FileVersion
    throw "This Scribble DLL does not include Test Lab: $resolvedAssembly (version $installedVersion). Install Scribble 2.0.132.0 or newer from https://github.com/datap0nd/scribble/releases/latest/download/ScribbleSetup.exe, then rerun this script in a NEW Windows PowerShell process. You can select the updated DLL explicitly with -AssemblyPath."
}
if (-not [string]::Equals($labType.Assembly.Location, $resolvedAssembly, [StringComparison]::OrdinalIgnoreCase) -and
    (Get-FileHash -LiteralPath $labType.Assembly.Location -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $resolvedAssembly -Algorithm SHA256).Hash) {
    throw "A different Scribble DLL is already loaded: $($labType.Assembly.Location). Start a NEW Windows PowerShell process and rerun with -AssemblyPath '$resolvedAssembly'."
}
Write-Output "Loaded Test Lab from $($labType.Assembly.Location)"
