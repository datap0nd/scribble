// Included inside [Code]. Only explicitly allowlisted shipped payloads participate.
// A loaded Windows image may be renamed even when DeleteFile is denied.
// Keep its old bytes available while installing the new image at its usual path.
type
  TPayloadRecord = record
    Name: String;
    OldHash: String;
    NewHash: String;
  end;
var
  PayloadRecords: array of TPayloadRecord;
  PayloadTransaction: String;
  PayloadCommitted: Boolean;
  PayloadStarted: Boolean;

function PayloadMoveFileEx(ExistingName, NewName: String; Flags: Cardinal): Boolean;
  external 'MoveFileExW@kernel32.dll stdcall';
function PayloadFileAttributes(Name: String): Cardinal;
  external 'GetFileAttributesW@kernel32.dll stdcall';

function AllowedPayload(Name: String): Boolean;
begin
  if Pos('|', Name) <> 0 then begin Result := False; Exit; end;
  Result := Pos('|' + Lowercase(Name) + '|',
    '|scribble.dll|scribbleupdater.exe|exceldatareader.dll|' +
    'pdfsharp-gdi.dll|pdfsharp.system.dll|pdfsharp.cryptography.dll|pdfsharp.shared.dll|' +
    'microsoft.extensions.logging.abstractions.dll|system.security.cryptography.pkcs.dll|' +
    'microsoft.extensions.dependencyinjection.abstractions.dll|microsoft.bcl.asyncinterfaces.dll|' +
    'system.buffers.dll|system.memory.dll|system.numerics.vectors.dll|' +
    'system.runtime.compilerservices.unsafe.dll|system.threading.tasks.extensions.dll|system.valuetuple.dll|' +
    'microsoft.web.webview2.core.dll|microsoft.web.webview2.winforms.dll|microsoft.web.webview2.wpf.dll|' +
    'runtimes\win-x86\native\webview2loader.dll|runtimes\win-x64\native\webview2loader.dll|' +
    'runtimes\win-arm64\native\webview2loader.dll|webview2loader.dll|scribblebrowserhost.exe|' +
    'scribblebrowserhost.exe.config|com.scribble.browser.json|' +
    'browserextension\manifest.json|browserextension\background.js|' +
    'browserextension\sidepanel.html|browserextension\sidepanel.css|' +
    'browserextension\sidepanel.js|browserextension\readme.md|') > 0;
end;

function PayloadHex(Value: String; ExpectedLength: Integer): Boolean;
var I: Integer;
begin
  Result := False;
  if Length(Value) <> ExpectedLength then Exit;
  for I := 1 to ExpectedLength do
    if Pos(Value[I], '0123456789abcdef') = 0 then Exit;
  Result := True;
end;

procedure CheckPayloadPath(Path: String);
var Cursor, Root: String; Attributes: Cardinal;
begin
  Root := RemoveBackslashUnlessRoot(ExpandConstant('{app}'));
  Cursor := Path;
  while Length(Cursor) >= Length(Root) do begin
    Attributes := PayloadFileAttributes(Cursor);
    if (Attributes <> $FFFFFFFF) and ((Attributes and $400) <> 0) then
      RaiseException('Links are not allowed in payload recovery paths: ' + Cursor);
    if CompareText(Cursor, Root) = 0 then Exit;
    Cursor := ExtractFileDir(Cursor);
  end;
  RaiseException('Payload recovery path escapes the installation folder.');
end;

function PayloadPath(Name: String): String;
begin
  if not AllowedPayload(Name) then RaiseException('Unknown installer payload: ' + Name);
  Result := AddBackslash(ExpandConstant('{app}')) + Name;
  CheckPayloadPath(Result);
end;

function PayloadJournal: String;
begin
  Result := ExpandConstant('{app}\Scribble.update-retirement.ini');
  CheckPayloadPath(Result);
end;

function PayloadToken: String;
begin
  Result := Copy(Lowercase(GetSHA256OfUnicodeString(
    GetDateTimeString('yyyymmddhhnnsszzz', '', '') + IntToStr(Random(2147483647)))), 1, 32);
end;

function RetiredPath(Name: String): String;
begin
  Result := PayloadPath(Name) + '.scribble-retired-' + PayloadTransaction;
  CheckPayloadPath(Result);
end;

procedure SavePayloadJournal;
var I: Integer; Data, Temporary: String;
begin
  Data := '[transaction]' + #13#10 + 'schema=1' + #13#10 +
    'id=' + PayloadTransaction + #13#10 + 'count=' + IntToStr(GetArrayLength(PayloadRecords)) + #13#10;
  if PayloadCommitted then Data := Data + 'committed=1' + #13#10
  else Data := Data + 'committed=0' + #13#10;
  for I := 0 to GetArrayLength(PayloadRecords) - 1 do
    Data := Data + '[file' + IntToStr(I) + ']' + #13#10 +
      'name=' + PayloadRecords[I].Name + #13#10 +
      'old=' + PayloadRecords[I].OldHash + #13#10 +
      'new=' + PayloadRecords[I].NewHash + #13#10;
  Temporary := PayloadJournal + '.tmp-' + PayloadToken;
  if FileExists(Temporary) then RaiseException('The installer journal temporary file already exists.');
  if not SaveStringToFile(Temporary, AnsiString(Data), False) then
    RaiseException('Could not write the payload recovery journal.');
  if not PayloadMoveFileEx(Temporary, PayloadJournal, 9) then
    RaiseException('Could not publish the payload recovery journal. No further payload was changed.');
end;

procedure LoadPayloadJournal;
var I, J, Count: Integer; Section, Value: String;
begin
  if GetIniString('transaction', 'schema', '', PayloadJournal) <> '1' then
    RaiseException('The payload recovery journal is not recognized. Preserve it for repair.');
  PayloadTransaction := GetIniString('transaction', 'id', '', PayloadJournal);
  if not PayloadHex(PayloadTransaction, 32) then RaiseException('Invalid payload transaction identity.');
  Value := GetIniString('transaction', 'committed', '', PayloadJournal);
  if (Value <> '0') and (Value <> '1') then RaiseException('Invalid payload transaction state.');
  PayloadCommitted := Value = '1';
  Count := StrToIntDef(GetIniString('transaction', 'count', '', PayloadJournal), -1);
  if (Count < 0) or (Count > 64) then RaiseException('Invalid payload recovery record count.');
  SetArrayLength(PayloadRecords, Count);
  for I := 0 to Count - 1 do begin
    Section := 'file' + IntToStr(I);
    PayloadRecords[I].Name := GetIniString(Section, 'name', '', PayloadJournal);
    PayloadRecords[I].OldHash := GetIniString(Section, 'old', '', PayloadJournal);
    PayloadRecords[I].NewHash := GetIniString(Section, 'new', '', PayloadJournal);
    if not AllowedPayload(PayloadRecords[I].Name) or
       ((PayloadRecords[I].NewHash <> '') and not PayloadHex(PayloadRecords[I].NewHash, 64)) or
       ((PayloadRecords[I].OldHash <> '') and not PayloadHex(PayloadRecords[I].OldHash, 64)) then
      RaiseException('Invalid payload recovery boundary. Existing files were preserved.');
    for J := 0 to I - 1 do
      if CompareText(PayloadRecords[J].Name, PayloadRecords[I].Name) = 0 then
        RaiseException('Duplicate payload recovery path.');
  end;
end;

procedure RestorePayloads;
var I: Integer; Destination, Retired, Rejected, CurrentHash: String;
begin
  for I := GetArrayLength(PayloadRecords) - 1 downto 0 do begin
    Destination := PayloadPath(PayloadRecords[I].Name);
    Retired := RetiredPath(PayloadRecords[I].Name);
    if FileExists(Retired) then begin
      if (PayloadRecords[I].OldHash = '') or
         (Lowercase(GetSHA256OfFile(Retired)) <> PayloadRecords[I].OldHash) then
        RaiseException('Retired payload changed; both copies were preserved: ' + Retired);
    end;
    if FileExists(Destination) then begin
      CurrentHash := Lowercase(GetSHA256OfFile(Destination));
      if (PayloadRecords[I].OldHash <> '') and (CurrentHash = PayloadRecords[I].OldHash) then
        Continue;
      if CurrentHash <> PayloadRecords[I].NewHash then
        RaiseException('Unexpected payload bytes; both copies were preserved: ' + Destination);
      if (PayloadRecords[I].OldHash <> '') and not FileExists(Retired) then
        RaiseException('The original payload is missing; the current copy was preserved: ' + Destination);
      Rejected := Destination + '.scribble-rejected-' + PayloadTransaction;
      CheckPayloadPath(Rejected);
      if FileExists(Rejected) or not RenameFile(Destination, Rejected) then
        RaiseException('Could not preserve the incomplete new payload: ' + Destination);
    end;
    if PayloadRecords[I].OldHash <> '' then begin
      if not FileExists(Retired) or not RenameFile(Retired, Destination) then
        RaiseException('Could not restore the original payload: ' + Destination);
      if Lowercase(GetSHA256OfFile(Destination)) <> PayloadRecords[I].OldHash then
        RaiseException('Restored payload verification failed: ' + Destination);
      Log('Restored original payload: ' + Destination);
    end;
  end;
end;

procedure ArchivePayloadJournal;
var Archive: String;
begin
  Archive := PayloadJournal + '.history-' + PayloadTransaction;
  if FileExists(Archive) or not RenameFile(PayloadJournal, Archive) then
    RaiseException('Could not retain the completed payload recovery journal.');
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  if FileExists(PayloadJournal) then begin
    try
      LoadPayloadJournal;
      if not PayloadCommitted then RestorePayloads;
      if FileExists(PayloadJournal) then ArchivePayloadJournal;
      SetArrayLength(PayloadRecords, 0);
      PayloadTransaction := '';
      PayloadCommitted := False;
    except
      Result := 'The previous update needs recovery. ' + GetExceptionMessage;
    end;
  end;
end;

procedure BeginPayloadRetirement(Name, NewHash: String; VerifyInstallerFile: Boolean);
var I, Index: Integer; Destination, Retired: String;
begin
  Destination := PayloadPath(Name);
  NewHash := Lowercase(NewHash);
  if (NewHash <> '') and not PayloadHex(NewHash, 64) then
    RaiseException('The installer payload hash is invalid.');
  if VerifyInstallerFile then
    if (NewHash = '') or (CompareText(CurrentFileName, Destination) <> 0) then
      RaiseException('The installer payload does not match its declared destination.');
  if not PayloadStarted then begin
    if FileExists(PayloadJournal) then RaiseException('An unresolved update journal exists.');
    PayloadTransaction := PayloadToken;
    PayloadCommitted := False;
    PayloadStarted := True;
  end;
  for I := 0 to GetArrayLength(PayloadRecords) - 1 do
    if CompareText(PayloadRecords[I].Name, Name) = 0 then RaiseException('Payload was installed twice.');
  Index := GetArrayLength(PayloadRecords);
  SetArrayLength(PayloadRecords, Index + 1);
  PayloadRecords[Index].Name := Name;
  PayloadRecords[Index].NewHash := NewHash;
  if FileExists(Destination) then PayloadRecords[Index].OldHash := Lowercase(GetSHA256OfFile(Destination));
  SavePayloadJournal;
  if PayloadRecords[Index].OldHash <> '' then begin
    Retired := RetiredPath(Name);
    if FileExists(Retired) or not RenameFile(Destination, Retired) then
      RaiseException('Could not retain the loaded payload. Close its app and retry: ' + Destination);
    if Lowercase(GetSHA256OfFile(Retired)) <> PayloadRecords[Index].OldHash then
      RaiseException('Retired payload verification failed: ' + Retired);
    Log('Retained previous payload: ' + Retired);
  end;
end;

procedure RetirePayload(Name, NewHash: String);
begin
  BeginPayloadRetirement(Name, NewHash, True);
end;

procedure RetireRemovedPayload(Name: String);
begin
  if FileExists(PayloadPath(Name)) then BeginPayloadRetirement(Name, '', False);
end;

procedure RetireDeselectedBrowserPayloads;
begin
  if WizardIsComponentSelected('browser') then Exit;
  RetireRemovedPayload('com.scribble.browser.json');
  RetireRemovedPayload('BrowserExtension\manifest.json');
  RetireRemovedPayload('BrowserExtension\background.js');
  RetireRemovedPayload('BrowserExtension\sidepanel.html');
  RetireRemovedPayload('BrowserExtension\sidepanel.css');
  RetireRemovedPayload('BrowserExtension\sidepanel.js');
  RetireRemovedPayload('BrowserExtension\README.md');
end;

procedure VerifyPayload(Name, ExpectedHash: String);
begin
  if Lowercase(GetSHA256OfFile(PayloadPath(Name))) <> Lowercase(ExpectedHash) then
    RaiseException('Installed payload verification failed: ' + Name);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then RetireDeselectedBrowserPayloads;
  if (CurStep = ssPostInstall) and PayloadStarted then begin
    PayloadCommitted := True;
    try SavePayloadJournal;
    except PayloadCommitted := False; RaiseException(GetExceptionMessage); end;
    Log('Payload update committed; previous mapped images and journal were retained.');
  end;
end;

procedure DeinitializeSetup;
begin
  if PayloadStarted and not PayloadCommitted then begin
    try
      RestorePayloads;
      ArchivePayloadJournal;
      Log('Payload update rolled back and original bytes verified.');
    except
      Log('PAYLOAD RECOVERY REQUIRED: ' + GetExceptionMessage + ' Journal: ' + PayloadJournal);
    end;
  end;
end;
