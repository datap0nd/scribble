#requires -Version 5.1
param(
    [string]$Compiler = "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '../../artifacts/payload-retirement-tests')
)
$ErrorActionPreference = 'Stop'
function Assert($condition, [string]$message) { if (-not $condition) { throw $message } }
function Hash([string]$path) { (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() }
$repository = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$root = [IO.Path]::GetFullPath((Join-Path $OutputDirectory ([guid]::NewGuid().ToString('N'))))
[IO.Directory]::CreateDirectory($root) | Out-Null
$include = Join-Path $repository 'installer/PayloadRetirement.iss'
$source = Join-Path $root 'new.dll'
$oldSource = Join-Path ([Environment]::GetFolderPath('System')) 'version.dll'
Copy-Item -LiteralPath (Join-Path ([Environment]::GetFolderPath('System')) 'winmm.dll') -Destination $source
$oldHash = Hash $oldSource
$newHash = Hash $source
Assert ($oldHash -ne $newHash) 'The fixture needs different old and new native DLL bytes.'

# These are private copies of Windows DLLs under two real installer payload
# names. Mapping them reproduces the Office image-section lock without Office.
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class RetirementFixtureNative {
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern IntPtr LoadLibraryEx(string path, IntPtr file, uint flags);
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool FreeLibrary(IntPtr module);
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern bool DeleteFile(string path);
}
'@
$script = @'
[Setup]
AppName=Scribble payload retirement fixture
AppVersion=1.0
DefaultDirName={tmp}\ScribblePayloadFixture
PrivilegesRequired=lowest
CreateAppDir=yes
Uninstallable=no
CreateUninstallRegKey=no
DisableProgramGroupPage=yes
DisableReadyPage=yes
CloseApplications=no
RestartApplications=no
OutputDir=.
OutputBaseFilename=PayloadFixture
[Types]
Name: "full"; Description: "All fixture payloads"
Name: "custom"; Description: "Custom fixture payloads"; Flags: iscustom
[Components]
Name: "browser"; Description: "Browser payloads"; Types: full
[Files]
Source: "@SOURCE@"; DestDir: "{app}"; DestName: "Scribble.dll"; Flags: ignoreversion; Check: ShouldCopy; BeforeInstall: RetirePayload('Scribble.dll', '@HASH@'); AfterInstall: VerifyPayload('Scribble.dll', '@HASH@')
Source: "@SOURCE@"; DestDir: "{app}"; DestName: "ScribbleBrowserHost.exe.config"; Flags: ignoreversion; Check: ShouldCopy; BeforeInstall: RetirePayload('ScribbleBrowserHost.exe.config', '@HASH@'); AfterInstall: VerifyPayload('ScribbleBrowserHost.exe.config', '@HASH@')
Source: "@SOURCE@"; DestDir: "{app}"; DestName: "com.scribble.browser.json"; Flags: ignoreversion; Check: ShouldCopy; Components: browser; BeforeInstall: RetirePayload('com.scribble.browser.json', '@HASH@'); AfterInstall: VerifyPayload('com.scribble.browser.json', '@HASH@')
Source: "@SOURCE@"; DestDir: "{app}\BrowserExtension"; DestName: "manifest.json"; Flags: ignoreversion; Check: ShouldCopy; Components: browser; BeforeInstall: RetirePayload('BrowserExtension\manifest.json', '@HASH@'); AfterInstall: VerifyPayload('BrowserExtension\manifest.json', '@HASH@')
Source: "@SOURCE@"; DestDir: "{app}\BrowserExtension"; DestName: "background.js"; Flags: ignoreversion; Check: ShouldCopy; Components: browser; BeforeInstall: RetirePayload('BrowserExtension\background.js', '@HASH@'); AfterInstall: VerifyPayload('BrowserExtension\background.js', '@HASH@')
Source: "@SOURCE@"; DestDir: "{app}\BrowserExtension"; DestName: "sidepanel.html"; Flags: ignoreversion; Check: ShouldCopy; Components: browser; BeforeInstall: RetirePayload('BrowserExtension\sidepanel.html', '@HASH@'); AfterInstall: VerifyPayload('BrowserExtension\sidepanel.html', '@HASH@')
Source: "@SOURCE@"; DestDir: "{app}\BrowserExtension"; DestName: "sidepanel.css"; Flags: ignoreversion; Check: ShouldCopy; Components: browser; BeforeInstall: RetirePayload('BrowserExtension\sidepanel.css', '@HASH@'); AfterInstall: VerifyPayload('BrowserExtension\sidepanel.css', '@HASH@')
Source: "@SOURCE@"; DestDir: "{app}\BrowserExtension"; DestName: "sidepanel.js"; Flags: ignoreversion; Check: ShouldCopy; Components: browser; BeforeInstall: RetirePayload('BrowserExtension\sidepanel.js', '@HASH@'); AfterInstall: VerifyPayload('BrowserExtension\sidepanel.js', '@HASH@')
Source: "@SOURCE@"; DestDir: "{app}\BrowserExtension"; DestName: "README.md"; Flags: ignoreversion; Check: ShouldCopy; Components: browser; BeforeInstall: RetirePayload('BrowserExtension\README.md', '@HASH@'); AfterInstall: VerifyPayload('BrowserExtension\README.md', '@HASH@')
Source: "@SOURCE@"; DestDir: "{app}"; DestName: "ExcelDataReader.dll"; Flags: ignoreversion; Check: ShouldCopy; BeforeInstall: RetirePayload('ExcelDataReader.dll', '@HASH@'); AfterInstall: FinishSecondPayload
[Code]
#include "@INCLUDE@"
procedure FixtureExitProcess(Code: Cardinal);
  external 'ExitProcess@kernel32.dll stdcall';
function ShouldCopy: Boolean;
begin
  Result := ExpandConstant('{param:MODE|success}') <> 'recover';
end;
procedure FinishSecondPayload;
var Mode: String;
begin
  Mode := ExpandConstant('{param:MODE|success}');
  { Exit before AfterInstall verification: recovery must already know the new hash. }
  if Mode = 'crash' then FixtureExitProcess(23);
  VerifyPayload('ExcelDataReader.dll', '@HASH@');
  if Mode = 'rollback' then RaiseException('Injected failure after two payload copies.');
end;
'@
$script = $script.Replace('@SOURCE@', $source).Replace('@HASH@', $newHash).Replace('@INCLUDE@', $include)
$iss = Join-Path $root 'Fixture.iss'
[IO.File]::WriteAllText($iss, $script)
& $Compiler $iss | Out-File -LiteralPath (Join-Path $root 'compile.log')
Assert ($LASTEXITCODE -eq 0) "Installer fixture compilation failed; see $root\compile.log"
$installer = Join-Path $root 'PayloadFixture.exe'
$mappedNames = @('Scribble.dll', 'ExcelDataReader.dll')
$sharedNames = @('Scribble.dll', 'ScribbleBrowserHost.exe.config', 'ExcelDataReader.dll')
$browserNames = @(
    'com.scribble.browser.json',
    'BrowserExtension\manifest.json',
    'BrowserExtension\background.js',
    'BrowserExtension\sidepanel.html',
    'BrowserExtension\sidepanel.css',
    'BrowserExtension\sidepanel.js',
    'BrowserExtension\README.md'
)
$names = @($sharedNames) + @($browserNames)
$results = @()
function Run-Installer([string]$directory, [string]$mode, [string]$logName, [bool]$deselectBrowser = $false) {
    $arguments = @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/NOCLOSEAPPLICATIONS',
        ('/DIR="' + $directory + '"'), ('/MODE=' + $mode),
        ('/LOG="' + (Join-Path $directory $logName) + '"'))
    if ($deselectBrowser) { $arguments += @('/TYPE=custom', '/COMPONENTS=""') }
    else { $arguments += '/TYPE=full' }
    $process = Start-Process -FilePath $installer -ArgumentList $arguments -WindowStyle Hidden -PassThru
    try {
        if (-not $process.WaitForExit(60000)) {
            # This handle belongs only to the fixture process started just above.
            $process.Kill()
            throw 'The private installer fixture exceeded its one-minute deadline.'
        }
        return $process.ExitCode
    } finally { $process.Dispose() }
}
foreach ($case in @('success', 'rollback', 'crash-recover', 'foreign-bytes', 'deselect-success', 'deselect-rollback')) {
    $directory = Join-Path $root $case
    [IO.Directory]::CreateDirectory($directory) | Out-Null
    [IO.Directory]::CreateDirectory((Join-Path $directory 'BrowserExtension')) | Out-Null
    $handles = @()
    try {
        foreach ($name in $names) {
            $destination = Join-Path $directory $name
            Copy-Item -LiteralPath $oldSource -Destination $destination
        }
        foreach ($name in $mappedNames) {
            $destination = Join-Path $directory $name
            $module = [RetirementFixtureNative]::LoadLibraryEx($destination, [IntPtr]::Zero, 1)
            Assert ($module -ne [IntPtr]::Zero) ('Mapping failed: ' + [Runtime.InteropServices.Marshal]::GetLastWin32Error())
            $handles += $module
            Assert (-not [RetirementFixtureNative]::DeleteFile($destination)) 'The loaded fixture DLL did not reproduce a deletion lock.'
            Assert ((Hash $destination) -eq $oldHash) 'The lock probe changed its private DLL.'
        }
        $deselectBrowser = $case.StartsWith('deselect-')
        $mode = if ($case -in @('crash-recover', 'foreign-bytes')) { 'crash' }
            elseif ($case -eq 'deselect-success') { 'success' }
            elseif ($case -eq 'deselect-rollback') { 'rollback' }
            else { $case }
        $exit = Run-Installer $directory $mode 'first-install.log' $deselectBrowser
        if ($case -in @('success', 'deselect-success')) {
            Assert ($exit -eq 0) "Successful retirement installer returned $exit."
            foreach ($name in $sharedNames) { Assert ((Hash (Join-Path $directory $name)) -eq $newHash) 'New shared payload bytes are missing.' }
            if ($deselectBrowser) {
                foreach ($name in $browserNames) { Assert (-not (Test-Path -LiteralPath (Join-Path $directory $name))) 'Deselected browser payload remained at its live path.' }
            } else {
                foreach ($name in $browserNames) { Assert ((Hash (Join-Path $directory $name)) -eq $newHash) 'New browser payload bytes are missing.' }
            }
            $retired = @(Get-ChildItem -LiteralPath $directory -Recurse -Filter '*.scribble-retired-*')
            Assert ($retired.Count -eq $names.Count) 'Success did not retain every replaced or deselected payload.'
            foreach ($file in $retired) { Assert ((Hash $file.FullName) -eq $oldHash) 'Retired mapped bytes changed.' }
            Assert (([IO.File]::ReadAllText((Join-Path $directory 'Scribble.update-retirement.ini'))) -match 'committed=1') 'Success was not durably committed.'
        } elseif ($case -in @('rollback', 'deselect-rollback')) {
            Assert ($exit -ne 0) 'Injected installer failure reported success.'
            foreach ($name in $names) { Assert ((Hash (Join-Path $directory $name)) -eq $oldHash) 'Rollback did not restore the original bytes.' }
            Assert (-not (Test-Path -LiteralPath (Join-Path $directory 'Scribble.update-retirement.ini'))) 'Completed rollback left an active journal.'
        } else {
            Assert ($exit -eq 23) "The fixture did not exit at its injected crash point: $exit."
            $journal = Join-Path $directory 'Scribble.update-retirement.ini'
            Assert (Test-Path -LiteralPath $journal) 'Interrupted setup lost its recovery journal.'
            if ($case -eq 'foreign-bytes') {
                $foreign = Join-Path $directory 'Scribble.dll'
                [IO.File]::WriteAllText($foreign, 'A different actor wrote these bytes after setup stopped.')
                $foreignHash = Hash $foreign
            }
            $recoveryExit = Run-Installer $directory 'recover' 'recovery-install.log'
            if ($case -eq 'crash-recover') {
                Assert ($recoveryExit -eq 0) "Recovery-only installer returned $recoveryExit."
                foreach ($name in $names) { Assert ((Hash (Join-Path $directory $name)) -eq $oldHash) 'Interrupted install recovery did not restore the original bytes.' }
                Assert (-not (Test-Path -LiteralPath $journal)) 'Recovery left an active journal.'
            } else {
                Assert ($recoveryExit -ne 0) 'Recovery overwrote unexpected third-party bytes.'
                Assert ((Hash $foreign) -eq $foreignHash) 'Unexpected current bytes were altered.'
                Assert (Test-Path -LiteralPath $journal) 'Unresolved recovery journal was discarded.'
                $originals = @(Get-ChildItem -LiteralPath $directory -Filter 'Scribble.dll.scribble-retired-*')
                Assert ($originals.Count -eq 1 -and (Hash $originals[0].FullName) -eq $oldHash) 'Unknown-byte recovery discarded the original mapped image.'
            }
        }
        $results += [ordered]@{ case = $case; status = 'passed'; first_exit_code = $exit }
        Write-Output "PASS: mapped DLL retirement / $case"
    } finally {
        foreach ($module in $handles) { [RetirementFixtureNative]::FreeLibrary($module) | Out-Null }
    }
}
[ordered]@{ schema = 1; include_sha256 = Hash $include; old_sha256 = $oldHash;
    new_sha256 = $newHash; results = $results } | ConvertTo-Json -Depth 5 |
    Set-Content -LiteralPath (Join-Path $root 'results.json') -Encoding UTF8
Write-Output "Retained fixture evidence: $root"
