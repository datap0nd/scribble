# Creates a personal self-signed code-signing certificate for Scribble
# and prepares everything needed to (a) let CI sign the DLL and the
# installer and (b) trust the signature on your own machines.
#
# Run once on any Windows machine:
#   powershell -ExecutionPolicy Bypass -File scripts\New-SigningCert.ps1
#
# It produces, in the current directory:
#   scribble-signing.cer          public certificate (safe to share)
#   scribble-signing.pfx          private key, protected by your password
#   scribble-signing.pfx.b64.txt  base64 of the PFX for the GitHub secret
#
# Then:
#   1. In the GitHub repo settings add two Actions secrets:
#        SIGNING_PFX           = contents of scribble-signing.pfx.b64.txt
#        SIGNING_PFX_PASSWORD  = the password you typed here
#   2. On every machine that runs Scribble, trust the certificate once
#      (elevated prompt not required for the current user):
#        certutil -user -addstore Root scribble-signing.cer
#        certutil -user -addstore TrustedPublisher scribble-signing.cer
#   3. Delete scribble-signing.pfx and the .b64.txt file after adding
#      the secrets - the .cer file is the only one to keep around.
#
# A self-signed certificate does not build SmartScreen reputation the
# way a paid certificate does, but once trusted on your machines the
# installer and DLL verify as signed by you, and any tampered build
# fails verification.

param(
    [string]$Subject = "CN=Scribble Personal Code Signing",
    [int]$Years = 5
)

$ErrorActionPreference = "Stop"

$password = Read-Host `
    -Prompt "Choose a password for the private key" `
    -AsSecureString

$cert = New-SelfSignedCertificate `
    -Type CodeSigningCert `
    -Subject $Subject `
    -KeyAlgorithm RSA `
    -KeyLength 3072 `
    -HashAlgorithm SHA256 `
    -KeyExportPolicy Exportable `
    -NotAfter (Get-Date).AddYears($Years) `
    -CertStoreLocation Cert:\CurrentUser\My

$cerPath = Join-Path (Get-Location) "scribble-signing.cer"
$pfxPath = Join-Path (Get-Location) "scribble-signing.pfx"
Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null
Export-PfxCertificate `
    -Cert $cert `
    -FilePath $pfxPath `
    -Password $password | Out-Null

$base64 = [Convert]::ToBase64String(
    [IO.File]::ReadAllBytes($pfxPath))
Set-Content `
    -Path "scribble-signing.pfx.b64.txt" `
    -Value $base64 `
    -Encoding ASCII

Write-Host ""
Write-Host "Created:"
Write-Host "  $cerPath"
Write-Host "  $pfxPath"
Write-Host "  scribble-signing.pfx.b64.txt"
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. GitHub repo -> Settings -> Secrets and variables -> Actions:"
Write-Host "       SIGNING_PFX          = contents of scribble-signing.pfx.b64.txt"
Write-Host "       SIGNING_PFX_PASSWORD = the password you just chose"
Write-Host "  2. On each machine that runs Scribble:"
Write-Host "       certutil -user -addstore Root scribble-signing.cer"
Write-Host "       certutil -user -addstore TrustedPublisher scribble-signing.cer"
Write-Host "  3. Delete the .pfx and .b64.txt files afterwards."
