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
$olderPublic=[Scribble.Updater.UpdateEngine]::NoUpdateMessage('2.0.190.0','2.0.91.0')
Assert (-not [string]::IsNullOrEmpty($olderPublic)) 'The frozen public release would start a downgrade instead of returning a no-update result.'
Assert ($olderPublic.Contains('Installed: 2.0.190.0') -and $olderPublic.Contains('Public stable: 2.0.91.0')) 'The no-update result does not identify both installed and public versions.'
Assert ($olderPublic.Contains('GitHub Actions')) 'Development installations have no explanation of their separate update path.'
$equalPublic=[Scribble.Updater.UpdateEngine]::NoUpdateMessage('2.0.91.0','2.0.91.0')
Assert (-not [string]::IsNullOrEmpty($equalPublic) -and $equalPublic.Contains('already up to date')) 'An equal public version would download and reinstall unnecessarily.'
Assert ($equalPublic.Contains('Installed: 2.0.91.0') -and $equalPublic.Contains('Public stable: 2.0.91.0')) 'The equal-version result does not identify both installed and public versions.'
Assert ($null -eq [Scribble.Updater.UpdateEngine]::NoUpdateMessage('2.0.9.0','2.0.10.0')) 'A newer public version was suppressed or compared lexically.'
Reject { [Scribble.Updater.UpdateEngine]::NoUpdateMessage('invalid','2.0.91.0') } 'An invalid installed version was reported as up to date.'
Reject { [Scribble.Updater.UpdateEngine]::NoUpdateMessage('2.0.91.0','invalid') } 'An invalid public version was reported as up to date.'
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
 Write-Output 'PASS: independent updater, public update decisions, malformed manifest, corrupted bytes, candidate identity and installed helper checks.'
} finally {
 # Only the exact two files created by this test are eligible for cleanup.
 if(Test-Path -LiteralPath (Join-Path $folder 'damaged.exe')){Remove-Item -LiteralPath (Join-Path $folder 'damaged.exe')}
 [IO.Directory]::Delete($folder)
}
