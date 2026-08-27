#define AppName "AI365"
#ifndef AppVersion
  #define AppVersion "2.0.0"
#endif
#define AppPublisher "AI365 contributors"
#define AppProgId "OutlookLocalAIChat.AddIn"
#define AppClsid "{{0D6E56F9-BE2D-4B94-B5E4-4C2DB0FD13E7}"
#define PaneProgId "OutlookLocalAIChat.ChatPane"
#define PaneClsid "{{14D24FA1-4342-442F-B68B-B68D7372794C}"
#define ExcelProgId "AI365.ExcelAddIn"
#define ExcelClsid "{{C0ABFA36-9854-434D-A542-DD834938737F}"
#define PptProgId "AI365.PowerPointAddIn"
#define PptClsid "{{69FAE812-274F-43F8-8F45-1B4EB22B5248}"
#define WordProgId "AI365.WordAddIn"
#define WordClsid "{{B49E9DB7-0C40-46A8-80A3-547626FE5331}"
#define OfficePaneProgId "AI365.OfficePane"
#define OfficePaneClsid "{{BC9047E7-9AFE-4F75-BBBC-27241B1DE2FA}"
#define ManagedCategory "{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
#define ControlCategory "{{40FC6ED4-2438-11CF-A3DB-080036F12502}"
#define LockbackInterface "{{000C0601-0000-0000-C000-000000000046}"
#define AssemblyName "OutlookLocalAIChat, Version=1.1.0.0, Culture=neutral, PublicKeyToken=f51b005bfa6d7cc3"
#define BrowserNativeHostName "com.ai365.browser"

[Setup]
AppId={{6BA7BCA9-F17E-4B50-8734-242063264160}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={localappdata}\Programs\OutlookLocalAIChat
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x86 x64compatible
OutputDir=..\artifacts
OutputBaseFilename=AI365Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
CloseApplicationsFilter=outlook.exe,excel.exe,powerpnt.exe,winword.exe,AI365BrowserHost.exe
RestartApplications=no
UninstallDisplayName={#AppName}
VersionInfoVersion={#AppVersion}
VersionInfoDescription=AI assistant suite for Outlook, Excel, PowerPoint, Word, Edge, and Chrome
VersionInfoProductName={#AppName}
VersionInfoCompany={#AppPublisher}

[Types]
Name: "full"; Description: "All apps (recommended)"
Name: "custom"; Description: "Choose which apps get AI365"; Flags: iscustom

[Components]
Name: "outlook"; Description: "AI365 for Outlook (mailbox chat and email drafts)"; Types: full
Name: "excel"; Description: "AI365 for Excel (workbook chat and draft sheets)"; Types: full
Name: "powerpoint"; Description: "AI365 for PowerPoint (presentation chat and draft slides)"; Types: full
Name: "word"; Description: "AI365 for Word (document chat and draft documents)"; Types: full
Name: "browser"; Description: "AI365 for Edge and Chrome (one-time browser approval required)"; Types: full

[Files]
Source: "..\src\OutlookLocalAIChat\bin\Release\OutlookLocalAIChat.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\src\OutlookLocalAIChat\bin\Release\Microsoft.Web.WebView2.Core.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\src\OutlookLocalAIChat\bin\Release\Microsoft.Web.WebView2.WinForms.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\src\OutlookLocalAIChat\bin\Release\Microsoft.Web.WebView2.Wpf.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\src\OutlookLocalAIChat\bin\Release\runtimes\win-x86\native\WebView2Loader.dll"; DestDir: "{app}\runtimes\win-x86\native"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\src\OutlookLocalAIChat\bin\Release\runtimes\win-x64\native\WebView2Loader.dll"; DestDir: "{app}\runtimes\win-x64\native"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\src\OutlookLocalAIChat\bin\Release\runtimes\win-arm64\native\WebView2Loader.dll"; DestDir: "{app}\runtimes\win-arm64\native"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\src\OutlookLocalAIChat\bin\Release\WebView2Loader.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\src\AI365.BrowserHost\bin\Release\AI365BrowserHost.exe"; DestDir: "{app}"; Flags: ignoreversion; Components: browser
Source: "..\src\AI365.BrowserHost\bin\Release\AI365BrowserHost.exe.config"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist; Components: browser
Source: "..\src\AI365.BrowserHost\com.ai365.browser.json"; DestDir: "{app}"; Flags: ignoreversion; Components: browser
Source: "..\src\AI365.BrowserExtension\manifest.json"; DestDir: "{app}\BrowserExtension"; Flags: ignoreversion; Components: browser
Source: "..\src\AI365.BrowserExtension\background.js"; DestDir: "{app}\BrowserExtension"; Flags: ignoreversion; Components: browser
Source: "..\src\AI365.BrowserExtension\sidepanel.html"; DestDir: "{app}\BrowserExtension"; Flags: ignoreversion; Components: browser
Source: "..\src\AI365.BrowserExtension\sidepanel.css"; DestDir: "{app}\BrowserExtension"; Flags: ignoreversion; Components: browser
Source: "..\src\AI365.BrowserExtension\sidepanel.js"; DestDir: "{app}\BrowserExtension"; Flags: ignoreversion; Components: browser
Source: "..\src\AI365.BrowserExtension\README.md"; DestDir: "{app}\BrowserExtension"; Flags: ignoreversion; Components: browser
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion

[InstallDelete]
; These paths contain shipped browser files only. Clearing them before
; reinstall also removes a browser component that the user deselected.
Type: filesandordirs; Name: "{app}\BrowserExtension"
Type: files; Name: "{app}\AI365BrowserHost.exe"
Type: files; Name: "{app}\AI365BrowserHost.exe.config"
Type: files; Name: "{app}\com.ai365.browser.json"

[UninstallDelete]
Type: filesandordirs; Name: "{app}\BrowserExtension"
Type: files; Name: "{app}\AI365BrowserHost.exe"
Type: files; Name: "{app}\AI365BrowserHost.exe.config"
Type: files; Name: "{app}\com.ai365.browser.json"

[Registry]
; Deselected components are cleanly unregistered on reinstall or
; change: Office never loads an add-in the user opted out of, and
; no stale COM class stays behind.
Root: HKCU32; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: none; Flags: deletekey; Components: not outlook
Root: HKCU64; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: none; Flags: deletekey; Components: not outlook; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}"; ValueType: none; Flags: deletekey; Components: not outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}"; ValueType: none; Flags: deletekey; Components: not outlook; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\{#AppProgId}"; ValueType: none; Flags: deletekey; Components: not outlook
Root: HKCU64; Subkey: "Software\Classes\{#AppProgId}"; ValueType: none; Flags: deletekey; Components: not outlook; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}"; ValueType: none; Flags: deletekey; Components: not outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}"; ValueType: none; Flags: deletekey; Components: not outlook; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\{#PaneProgId}"; ValueType: none; Flags: deletekey; Components: not outlook
Root: HKCU64; Subkey: "Software\Classes\{#PaneProgId}"; ValueType: none; Flags: deletekey; Components: not outlook; Check: IsWin64
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: none; Flags: deletekey; Components: not excel
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: none; Flags: deletekey; Components: not excel; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}"; ValueType: none; Flags: deletekey; Components: not excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}"; ValueType: none; Flags: deletekey; Components: not excel; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\{#ExcelProgId}"; ValueType: none; Flags: deletekey; Components: not excel
Root: HKCU64; Subkey: "Software\Classes\{#ExcelProgId}"; ValueType: none; Flags: deletekey; Components: not excel; Check: IsWin64
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: none; Flags: deletekey; Components: not powerpoint
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: none; Flags: deletekey; Components: not powerpoint; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}"; ValueType: none; Flags: deletekey; Components: not powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}"; ValueType: none; Flags: deletekey; Components: not powerpoint; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\{#PptProgId}"; ValueType: none; Flags: deletekey; Components: not powerpoint
Root: HKCU64; Subkey: "Software\Classes\{#PptProgId}"; ValueType: none; Flags: deletekey; Components: not powerpoint; Check: IsWin64
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: none; Flags: deletekey; Components: not word
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: none; Flags: deletekey; Components: not word; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}"; ValueType: none; Flags: deletekey; Components: not word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}"; ValueType: none; Flags: deletekey; Components: not word; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\{#WordProgId}"; ValueType: none; Flags: deletekey; Components: not word
Root: HKCU64; Subkey: "Software\Classes\{#WordProgId}"; ValueType: none; Flags: deletekey; Components: not word; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}"; ValueType: none; Flags: deletekey; Components: not excel and not powerpoint and not word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}"; ValueType: none; Flags: deletekey; Components: not excel and not powerpoint and not word; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\{#OfficePaneProgId}"; ValueType: none; Flags: deletekey; Components: not excel and not powerpoint and not word
Root: HKCU64; Subkey: "Software\Classes\{#OfficePaneProgId}"; ValueType: none; Flags: deletekey; Components: not excel and not powerpoint and not word; Check: IsWin64
Root: HKCU32; Subkey: "Software\Microsoft\Edge\NativeMessagingHosts\{#BrowserNativeHostName}"; ValueType: none; Flags: deletekey; Components: not browser
Root: HKCU32; Subkey: "Software\Google\Chrome\NativeMessagingHosts\{#BrowserNativeHostName}"; ValueType: none; Flags: deletekey; Components: not browser

; The browser bridge is private to this Windows account. Setup stages the
; unpacked extension but Edge or Chrome still requires the user's one-time
; approval; no policy, force-install, or browser-profile keys are written.
Root: HKCU32; Subkey: "Software\Microsoft\Edge\NativeMessagingHosts\{#BrowserNativeHostName}"; ValueType: string; ValueName: ""; ValueData: "{app}\com.ai365.browser.json"; Flags: uninsdeletekey; Components: browser
Root: HKCU32; Subkey: "Software\Google\Chrome\NativeMessagingHosts\{#BrowserNativeHostName}"; ValueType: string; ValueName: ""; ValueData: "{app}\com.ai365.browser.json"; Flags: uninsdeletekey; Components: browser

; 32-bit COM registration. Required for 32-bit Office, including on 64-bit Windows.
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.AddIn"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.AddIn"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#AppProgId}"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\{#AppProgId}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\{#AppProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#AppClsid}"; Components: outlook
Root: HKCU32; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey; Components: outlook
Root: HKCU32; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Local mailbox AI chat with one linked unsent draft."; Components: outlook
Root: HKCU32; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Components: outlook
Root: HKCU32; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Components: outlook

; Managed ActiveX control hosted by Office as the native Outlook sidebar.
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}"; ValueType: string; ValueName: ""; ValueData: "{#AppName} Sidebar"; Flags: uninsdeletekey; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.UI.ChatPane"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.UI.ChatPane"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#PaneProgId}"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Implemented Categories\{#ControlCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Programmable"; ValueType: string; ValueName: ""; ValueData: ""; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\{#PaneProgId}"; ValueType: string; ValueName: ""; ValueData: "{#AppName} Sidebar"; Flags: uninsdeletekey; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\{#PaneProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#PaneClsid}"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\Interface\{#LockbackInterface}"; ValueType: string; ValueName: ""; ValueData: "Office .NET Framework Lockback Bypass Key"

; 64-bit COM registration. Written only on 64-bit Windows.
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.AddIn"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.AddIn"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#AppProgId}"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\{#AppProgId}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\{#AppProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#AppClsid}"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Local mailbox AI chat with one linked unsent draft."; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#AppProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Check: IsWin64; Components: outlook

Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}"; ValueType: string; ValueName: ""; ValueData: "{#AppName} Sidebar"; Flags: uninsdeletekey; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.UI.ChatPane"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.UI.ChatPane"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#PaneProgId}"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Implemented Categories\{#ControlCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Programmable"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\{#PaneProgId}"; ValueType: string; ValueName: ""; ValueData: "{#AppName} Sidebar"; Flags: uninsdeletekey; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\{#PaneProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#PaneClsid}"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\Interface\{#LockbackInterface}"; ValueType: string; ValueName: ""; ValueData: "Office .NET Framework Lockback Bypass Key"; Check: IsWin64


; AI365 for Excel add-in (32-bit).
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}"; ValueType: string; ValueName: ""; ValueData: "AI365 for Excel"; Flags: uninsdeletekey; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.ExcelAddIn"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.ExcelAddIn"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#ExcelProgId}"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Components: excel
Root: HKCU32; Subkey: "Software\Classes\{#ExcelProgId}"; ValueType: string; ValueName: ""; ValueData: "AI365 for Excel"; Flags: uninsdeletekey; Components: excel
Root: HKCU32; Subkey: "Software\Classes\{#ExcelProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#ExcelClsid}"; Components: excel
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "AI365"; Flags: uninsdeletekey; Components: excel
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Chat with your workbook. AI365 never saves, deletes, or sends."; Components: excel
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Components: excel
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Components: excel

; AI365 for PowerPoint add-in (32-bit).
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}"; ValueType: string; ValueName: ""; ValueData: "AI365 for PowerPoint"; Flags: uninsdeletekey; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.PowerPointAddIn"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.PowerPointAddIn"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#PptProgId}"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\{#PptProgId}"; ValueType: string; ValueName: ""; ValueData: "AI365 for PowerPoint"; Flags: uninsdeletekey; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\{#PptProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#PptClsid}"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "AI365"; Flags: uninsdeletekey; Components: powerpoint
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Chat with your presentation. AI365 never saves, deletes, or sends."; Components: powerpoint
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Components: powerpoint

; Managed ActiveX control hosted as the Excel/PowerPoint sidebar (32-bit).
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}"; ValueType: string; ValueName: ""; ValueData: "AI365 Sidebar"; Flags: uninsdeletekey; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.UI.OfficeChatPane"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.UI.OfficeChatPane"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#OfficePaneProgId}"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Implemented Categories\{#ControlCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Programmable"; ValueType: string; ValueName: ""; ValueData: ""; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\{#OfficePaneProgId}"; ValueType: string; ValueName: ""; ValueData: "AI365 Sidebar"; Flags: uninsdeletekey; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\{#OfficePaneProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#OfficePaneClsid}"; Components: excel powerpoint word

; AI365 for Excel add-in (64-bit).
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}"; ValueType: string; ValueName: ""; ValueData: "AI365 for Excel"; Flags: uninsdeletekey; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.ExcelAddIn"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.ExcelAddIn"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#ExcelProgId}"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\{#ExcelProgId}"; ValueType: string; ValueName: ""; ValueData: "AI365 for Excel"; Flags: uninsdeletekey; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\{#ExcelProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#ExcelClsid}"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "AI365"; Flags: uninsdeletekey; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Chat with your workbook. AI365 never saves, deletes, or sends."; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Check: IsWin64; Components: excel

; AI365 for PowerPoint add-in (64-bit).
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}"; ValueType: string; ValueName: ""; ValueData: "AI365 for PowerPoint"; Flags: uninsdeletekey; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.PowerPointAddIn"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.PowerPointAddIn"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#PptProgId}"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\{#PptProgId}"; ValueType: string; ValueName: ""; ValueData: "AI365 for PowerPoint"; Flags: uninsdeletekey; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\{#PptProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#PptClsid}"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "AI365"; Flags: uninsdeletekey; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Chat with your presentation. AI365 never saves, deletes, or sends."; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Check: IsWin64; Components: powerpoint

; Managed ActiveX control hosted as the Excel/PowerPoint sidebar (64-bit).
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}"; ValueType: string; ValueName: ""; ValueData: "AI365 Sidebar"; Flags: uninsdeletekey; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.UI.OfficeChatPane"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.UI.OfficeChatPane"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#OfficePaneProgId}"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Implemented Categories\{#ControlCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Programmable"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\{#OfficePaneProgId}"; ValueType: string; ValueName: ""; ValueData: "AI365 Sidebar"; Flags: uninsdeletekey; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\{#OfficePaneProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#OfficePaneClsid}"; Check: IsWin64; Components: excel powerpoint word


; AI365 for Word add-in (32-bit).
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}"; ValueType: string; ValueName: ""; ValueData: "AI365 for Word"; Flags: uninsdeletekey; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.WordAddIn"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.WordAddIn"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#WordProgId}"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Components: word
Root: HKCU32; Subkey: "Software\Classes\{#WordProgId}"; ValueType: string; ValueName: ""; ValueData: "AI365 for Word"; Flags: uninsdeletekey; Components: word
Root: HKCU32; Subkey: "Software\Classes\{#WordProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#WordClsid}"; Components: word
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "AI365"; Flags: uninsdeletekey; Components: word
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Chat with your document. AI365 never saves, deletes, or sends."; Components: word
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Components: word
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Components: word

; AI365 for Word add-in (64-bit).
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}"; ValueType: string; ValueName: ""; ValueData: "AI365 for Word"; Flags: uninsdeletekey; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.WordAddIn"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "OutlookLocalAIChat.WordAddIn"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#WordProgId}"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\{#WordProgId}"; ValueType: string; ValueName: ""; ValueData: "AI365 for Word"; Flags: uninsdeletekey; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\{#WordProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#WordClsid}"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "AI365"; Flags: uninsdeletekey; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Chat with your document. AI365 never saves, deletes, or sends."; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Check: IsWin64; Components: word

[Icons]
Name: "{group}\Set up AI365 in Microsoft Edge"; Filename: "{app}\AI365BrowserHost.exe"; Parameters: "--setup edge"; Components: browser
Name: "{group}\Set up AI365 in Google Chrome"; Filename: "{app}\AI365BrowserHost.exe"; Parameters: "--setup chrome"; Components: browser

[Run]
Filename: "{app}\AI365BrowserHost.exe"; Parameters: "--setup auto"; Description: "Finish setting up AI365 in Edge or Chrome"; Flags: nowait postinstall skipifsilent; Components: browser

[Code]
function GetAssemblyCodeBase(Param: String): String;
var
  Path: String;
begin
  Path := ExpandConstant('{app}\OutlookLocalAIChat.dll');
  StringChangeEx(Path, '\', '/', True);
  Result := 'file:///' + Path;
end;
