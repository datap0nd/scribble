# Signs one file with the certificate provided through the CI
# secrets SIGNING_PFX (base64 PFX) and SIGNING_PFX_PASSWORD. Used by
# the build workflow for the add-in DLL and the installer; does
# nothing useful outside CI. Fails loudly when signing fails so an
# expected-signed build can never ship unsigned.

param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

$ErrorActionPreference = "Stop"

if (-not $env:SIGNING_PFX) {
    throw "SIGNING_PFX is not set."
}
if (-not (Test-Path $Path)) {
    throw "File to sign not found: $Path"
}

$signtool = Get-ChildItem `
    "${env:ProgramFiles(x86)}\Windows Kits\10\bin\*\x64\signtool.exe" `
    -ErrorAction SilentlyContinue |
    Sort-Object FullName |
    Select-Object -Last 1
if (-not $signtool) {
    throw "signtool.exe was not found in the Windows SDK."
}

$pfxPath = Join-Path $env:RUNNER_TEMP "scribble-signing.pfx"
[IO.File]::WriteAllBytes(
    $pfxPath,
    [Convert]::FromBase64String($env:SIGNING_PFX))
try {
    & $signtool.FullName sign `
        /f $pfxPath `
        /p $env:SIGNING_PFX_PASSWORD `
        /fd SHA256 `
        /tr http://timestamp.digicert.com `
        /td SHA256 `
        $Path
    if ($LASTEXITCODE -ne 0) {
        throw "signtool failed for $Path."
    }
}
finally {
    Remove-Item $pfxPath -Force -ErrorAction SilentlyContinue
}

Write-Host "Signed $Path"
