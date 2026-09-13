#requires -Version 5.1
param(
 [string]$HelperPath=(Join-Path $PSScriptRoot '../../src/Scribble.Updater/bin/Release/ScribbleUpdater.exe'),
 [string]$InstallerPath,
 [string]$CandidatePath,
 [string]$InstalledDirectory
)
$ErrorActionPreference='Stop'
function Assert($value,$message){if(-not $value){throw $message}}
function Reject([scriptblock]$action,$message){$rejected=$false;try{& $action}catch{$rejected=$true};Assert $rejected $message}
$helper=(Resolve-Path $HelperPath).Path
$assembly=[Reflection.Assembly]::LoadFrom($helper)
Assert (-not @($assembly.GetReferencedAssemblies() | Where-Object { $_.Name -like 'Scribble*' }).Count) 'The updater would lock an installed Scribble assembly.'
Reject { [Scribble.Updater.UpdateEngine]::ParseCandidate('{}') } 'Incomplete manifest was accepted.'
$folder=Join-Path $env:TEMP ('scribble-updater-validation-'+[guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($folder) | Out-Null
try {
 $bytes=New-Object byte[] (256*1024);$bytes[0]=77;$bytes[1]=90
 $damaged=Join-Path $folder 'damaged.exe';[IO.File]::WriteAllBytes($damaged,$bytes)
 $candidate=[Scribble.Updater.UpdateEngine]::ParseCandidate((@{version='2.0.1.0';commit=('a'*40);installer_sha256=('b'*64)} | ConvertTo-Json))
 Reject { [Scribble.Updater.UpdateEngine]::ValidateInstaller($damaged,$candidate) } 'Corrupt installer passed checksum validation.'
 if($InstallerPath) {
  $installer=(Resolve-Path $InstallerPath).Path
  $candidate=if($CandidatePath){[Scribble.Updater.UpdateEngine]::ParseCandidate([IO.File]::ReadAllText((Resolve-Path $CandidatePath).Path))}else{
   [Scribble.Updater.UpdateEngine]::ParseCandidate((@{version=[Scribble.Updater.UpdateEngine]::VersionOf($installer);commit=('a'*40);installer_sha256=[Scribble.Updater.UpdateEngine]::Hash($installer)} | ConvertTo-Json))
  }
  [Scribble.Updater.UpdateEngine]::ValidateInstaller($installer,$candidate)
  $candidate.version='99.0.0.0'
  Reject { [Scribble.Updater.UpdateEngine]::ValidateInstaller($installer,$candidate) } 'Wrong-version installer was accepted.'
 }
 if($InstalledDirectory){
  $installed=Join-Path $InstalledDirectory 'ScribbleUpdater.exe'
  Assert ((Test-Path -LiteralPath $installed) -and [Scribble.Updater.UpdateEngine]::Hash($installed) -eq [Scribble.Updater.UpdateEngine]::Hash($helper)) 'The installer omitted or altered its independent update helper.'
 }
 Write-Output 'PASS: independent updater, malformed manifest, corrupted bytes, candidate identity and installed helper checks.'
} finally {
 # Only the exact two files created by this test are eligible for cleanup.
 if(Test-Path -LiteralPath (Join-Path $folder 'damaged.exe')){Remove-Item -LiteralPath (Join-Path $folder 'damaged.exe')}
 [IO.Directory]::Delete($folder)
}
