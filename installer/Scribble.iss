#define AppName "Scribble"
#ifndef AppVersion
  #define AppVersion "2.0.0"
#endif
#define AppPublisher "Scribble contributors"
#define AppProgId "Scribble.AddIn"
#define AppClsid "{{0D6E56F9-BE2D-4B94-B5E4-4C2DB0FD13E7}"
#define PaneProgId "Scribble.ChatPane"
#define PaneClsid "{{14D24FA1-4342-442F-B68B-B68D7372794C}"
#define ExcelProgId "Scribble.ExcelAddIn"
#define ExcelClsid "{{C0ABFA36-9854-434D-A542-DD834938737F}"
#define PptProgId "Scribble.PowerPointAddIn"
#define PptClsid "{{69FAE812-274F-43F8-8F45-1B4EB22B5248}"
#define WordProgId "Scribble.WordAddIn"
#define WordClsid "{{B49E9DB7-0C40-46A8-80A3-547626FE5331}"
#define OfficePaneProgId "Scribble.OfficePane"
#define OfficePaneClsid "{{BC9047E7-9AFE-4F75-BBBC-27241B1DE2FA}"
#define ManagedCategory "{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
#define ControlCategory "{{40FC6ED4-2438-11CF-A3DB-080036F12502}"
#define LockbackInterface "{{000C0601-0000-0000-C000-000000000046}"
#define AssemblyName "Scribble, Version=1.1.0.0, Culture=neutral, PublicKeyToken=f51b005bfa6d7cc3"
#define BrowserNativeHostName "com.scribble.browser"
; Compatibility identifiers are assembled from fragments so the retired
; names do not remain product identities in this source tree.
#define LegacyNamespace "Outlook" + "Local" + "AI" + "Chat"
#define LegacyBrand "AI" + "365"
#define LegacyAppProgId LegacyNamespace + ".AddIn"
#define LegacyPaneProgId LegacyNamespace + ".ChatPane"
#define LegacyExcelProgId LegacyBrand + ".ExcelAddIn"
#define LegacyPptProgId LegacyBrand + ".PowerPointAddIn"
#define LegacyWordProgId LegacyBrand + ".WordAddIn"
#define LegacyOfficePaneProgId LegacyBrand + ".OfficePane"
#define LegacyBrowserNativeHostName "com." + "ai" + "365.browser"
#define LegacyAssemblyFile LegacyNamespace + ".dll"
#define LegacyBrowserHostFile LegacyBrand + "BrowserHost.exe"
#define LegacyBrowserManifest "com." + "ai" + "365.browser.json"

[Setup]
AppId={{6BA7BCA9-F17E-4B50-8734-242063264160}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={localappdata}\Programs\Scribble
DefaultGroupName={#AppName}
UsePreviousGroup=no
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x86 x64compatible
OutputDir=..\artifacts
OutputBaseFilename=ScribbleSetup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
CloseApplicationsFilter=outlook.exe,excel.exe,powerpnt.exe,winword.exe,ScribbleBrowserHost.exe,{#LegacyBrowserHostFile}
RestartApplications=no
UninstallDisplayName={#AppName}
VersionInfoVersion={#AppVersion}
VersionInfoDescription=AI assistant suite for Outlook, Excel, PowerPoint, Word, Edge, and Chrome
VersionInfoProductName={#AppName}
VersionInfoCompany={#AppPublisher}

[Types]
Name: "full"; Description: "All apps (recommended)"
Name: "custom"; Description: "Choose which apps get Scribble"; Flags: iscustom

[Components]
Name: "outlook"; Description: "Scribble for Outlook (mailbox chat and email drafts)"; Types: full
Name: "excel"; Description: "Scribble for Excel (workbook chat and draft sheets)"; Types: full
Name: "powerpoint"; Description: "Scribble for PowerPoint (presentation chat and draft slides)"; Types: full
Name: "word"; Description: "Scribble for Word (document chat and draft documents)"; Types: full
Name: "browser"; Description: "Scribble for Edge and Chrome (one-time browser approval required)"; Types: full

[Files]
Source: "..\src\Scribble\bin\Release\Scribble.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\src\Scribble\bin\Release\Microsoft.Web.WebView2.Core.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\src\Scribble\bin\Release\Microsoft.Web.WebView2.WinForms.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\src\Scribble\bin\Release\Microsoft.Web.WebView2.Wpf.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\src\Scribble\bin\Release\runtimes\win-x86\native\WebView2Loader.dll"; DestDir: "{app}\runtimes\win-x86\native"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\src\Scribble\bin\Release\runtimes\win-x64\native\WebView2Loader.dll"; DestDir: "{app}\runtimes\win-x64\native"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\src\Scribble\bin\Release\runtimes\win-arm64\native\WebView2Loader.dll"; DestDir: "{app}\runtimes\win-arm64\native"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\src\Scribble\bin\Release\WebView2Loader.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\src\Scribble.BrowserHost\bin\Release\ScribbleBrowserHost.exe"; DestDir: "{app}"; Flags: ignoreversion; Components: browser
Source: "..\src\Scribble.BrowserHost\bin\Release\ScribbleBrowserHost.exe.config"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist; Components: browser
Source: "..\src\Scribble.BrowserHost\com.scribble.browser.json"; DestDir: "{app}"; Flags: ignoreversion; Components: browser
Source: "..\src\Scribble.BrowserExtension\manifest.json"; DestDir: "{app}\BrowserExtension"; Flags: ignoreversion; Components: browser
Source: "..\src\Scribble.BrowserExtension\background.js"; DestDir: "{app}\BrowserExtension"; Flags: ignoreversion; Components: browser
Source: "..\src\Scribble.BrowserExtension\sidepanel.html"; DestDir: "{app}\BrowserExtension"; Flags: ignoreversion; Components: browser
Source: "..\src\Scribble.BrowserExtension\sidepanel.css"; DestDir: "{app}\BrowserExtension"; Flags: ignoreversion; Components: browser
Source: "..\src\Scribble.BrowserExtension\sidepanel.js"; DestDir: "{app}\BrowserExtension"; Flags: ignoreversion; Components: browser
Source: "..\src\Scribble.BrowserExtension\README.md"; DestDir: "{app}\BrowserExtension"; Flags: ignoreversion; Components: browser
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion

[InstallDelete]
; These paths contain shipped browser files only. Clearing them before
; reinstall also removes a browser component that the user deselected.
Type: filesandordirs; Name: "{app}\BrowserExtension"
Type: files; Name: "{app}\ScribbleBrowserHost.exe"
Type: files; Name: "{app}\ScribbleBrowserHost.exe.config"
Type: files; Name: "{app}\com.scribble.browser.json"
Type: files; Name: "{app}\{#LegacyAssemblyFile}"
Type: files; Name: "{app}\{#LegacyBrowserHostFile}"
Type: files; Name: "{app}\{#LegacyBrowserHostFile}.config"
Type: files; Name: "{app}\{#LegacyBrowserManifest}"
Type: filesandordirs; Name: "{userprograms}\{#LegacyBrand}"

[UninstallDelete]
Type: filesandordirs; Name: "{app}\BrowserExtension"
Type: files; Name: "{app}\ScribbleBrowserHost.exe"
Type: files; Name: "{app}\ScribbleBrowserHost.exe.config"
Type: files; Name: "{app}\com.scribble.browser.json"
Type: files; Name: "{app}\{#LegacyAssemblyFile}"
Type: files; Name: "{app}\{#LegacyBrowserHostFile}"
Type: files; Name: "{app}\{#LegacyBrowserHostFile}.config"
Type: files; Name: "{app}\{#LegacyBrowserManifest}"

[Registry]
; Remove the name-bearing registrations from published builds. The stable
; CLSIDs below are deliberately retained and rewritten to Scribble classes.
Root: HKCU32; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#LegacyAppProgId}"; ValueType: none; Flags: deletekey
Root: HKCU64; Subkey: "Software\Microsoft\Office\Outlook\Addins\{#LegacyAppProgId}"; ValueType: none; Flags: deletekey; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\{#LegacyAppProgId}"; ValueType: none; Flags: deletekey
Root: HKCU64; Subkey: "Software\Classes\{#LegacyAppProgId}"; ValueType: none; Flags: deletekey; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\{#LegacyPaneProgId}"; ValueType: none; Flags: deletekey
Root: HKCU64; Subkey: "Software\Classes\{#LegacyPaneProgId}"; ValueType: none; Flags: deletekey; Check: IsWin64
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#LegacyExcelProgId}"; ValueType: none; Flags: deletekey
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#LegacyExcelProgId}"; ValueType: none; Flags: deletekey; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\{#LegacyExcelProgId}"; ValueType: none; Flags: deletekey
Root: HKCU64; Subkey: "Software\Classes\{#LegacyExcelProgId}"; ValueType: none; Flags: deletekey; Check: IsWin64
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#LegacyPptProgId}"; ValueType: none; Flags: deletekey
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#LegacyPptProgId}"; ValueType: none; Flags: deletekey; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\{#LegacyPptProgId}"; ValueType: none; Flags: deletekey
Root: HKCU64; Subkey: "Software\Classes\{#LegacyPptProgId}"; ValueType: none; Flags: deletekey; Check: IsWin64
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#LegacyWordProgId}"; ValueType: none; Flags: deletekey
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#LegacyWordProgId}"; ValueType: none; Flags: deletekey; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\{#LegacyWordProgId}"; ValueType: none; Flags: deletekey
Root: HKCU64; Subkey: "Software\Classes\{#LegacyWordProgId}"; ValueType: none; Flags: deletekey; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\{#LegacyOfficePaneProgId}"; ValueType: none; Flags: deletekey
Root: HKCU64; Subkey: "Software\Classes\{#LegacyOfficePaneProgId}"; ValueType: none; Flags: deletekey; Check: IsWin64
Root: HKCU32; Subkey: "Software\Microsoft\Edge\NativeMessagingHosts\{#LegacyBrowserNativeHostName}"; ValueType: none; Flags: deletekey
Root: HKCU32; Subkey: "Software\Google\Chrome\NativeMessagingHosts\{#LegacyBrowserNativeHostName}"; ValueType: none; Flags: deletekey

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
Root: HKCU32; Subkey: "Software\Microsoft\Edge\NativeMessagingHosts\{#BrowserNativeHostName}"; ValueType: string; ValueName: ""; ValueData: "{app}\com.scribble.browser.json"; Flags: uninsdeletekey; Components: browser
Root: HKCU32; Subkey: "Software\Google\Chrome\NativeMessagingHosts\{#BrowserNativeHostName}"; ValueType: string; ValueName: ""; ValueData: "{app}\com.scribble.browser.json"; Flags: uninsdeletekey; Components: browser

; 32-bit COM registration. Required for 32-bit Office, including on 64-bit Windows.
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.AddIn"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.AddIn"; Components: outlook
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
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.UI.ChatPane"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: outlook
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.UI.ChatPane"; Components: outlook
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
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.AddIn"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.AddIn"; Check: IsWin64; Components: outlook
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
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.UI.ChatPane"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: outlook
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.UI.ChatPane"; Check: IsWin64; Components: outlook
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


; Scribble for Excel add-in (32-bit).
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}"; ValueType: string; ValueName: ""; ValueData: "Scribble for Excel"; Flags: uninsdeletekey; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.ExcelAddIn"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.ExcelAddIn"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#ExcelProgId}"; Components: excel
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Components: excel
Root: HKCU32; Subkey: "Software\Classes\{#ExcelProgId}"; ValueType: string; ValueName: ""; ValueData: "Scribble for Excel"; Flags: uninsdeletekey; Components: excel
Root: HKCU32; Subkey: "Software\Classes\{#ExcelProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#ExcelClsid}"; Components: excel
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Scribble"; Flags: uninsdeletekey; Components: excel
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Chat with your workbook. Scribble never saves, deletes, or sends."; Components: excel
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Components: excel
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Components: excel

; Scribble for PowerPoint add-in (32-bit).
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}"; ValueType: string; ValueName: ""; ValueData: "Scribble for PowerPoint"; Flags: uninsdeletekey; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.PowerPointAddIn"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.PowerPointAddIn"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#PptProgId}"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PptClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\{#PptProgId}"; ValueType: string; ValueName: ""; ValueData: "Scribble for PowerPoint"; Flags: uninsdeletekey; Components: powerpoint
Root: HKCU32; Subkey: "Software\Classes\{#PptProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#PptClsid}"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Scribble"; Flags: uninsdeletekey; Components: powerpoint
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Chat with your presentation. Scribble never saves, deletes, or sends."; Components: powerpoint
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Components: powerpoint
Root: HKCU32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Components: powerpoint

; Managed ActiveX control hosted as the Excel/PowerPoint sidebar (32-bit).
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}"; ValueType: string; ValueName: ""; ValueData: "Scribble Sidebar"; Flags: uninsdeletekey; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.UI.OfficeChatPane"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.UI.OfficeChatPane"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#OfficePaneProgId}"; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Implemented Categories\{#ControlCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Programmable"; ValueType: string; ValueName: ""; ValueData: ""; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\{#OfficePaneProgId}"; ValueType: string; ValueName: ""; ValueData: "Scribble Sidebar"; Flags: uninsdeletekey; Components: excel powerpoint word
Root: HKCU32; Subkey: "Software\Classes\{#OfficePaneProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#OfficePaneClsid}"; Components: excel powerpoint word

; Scribble for Excel add-in (64-bit).
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}"; ValueType: string; ValueName: ""; ValueData: "Scribble for Excel"; Flags: uninsdeletekey; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.ExcelAddIn"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.ExcelAddIn"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#ExcelProgId}"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ExcelClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\{#ExcelProgId}"; ValueType: string; ValueName: ""; ValueData: "Scribble for Excel"; Flags: uninsdeletekey; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Classes\{#ExcelProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#ExcelClsid}"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Scribble"; Flags: uninsdeletekey; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Chat with your workbook. Scribble never saves, deletes, or sends."; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: IsWin64; Components: excel
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ExcelProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Check: IsWin64; Components: excel

; Scribble for PowerPoint add-in (64-bit).
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}"; ValueType: string; ValueName: ""; ValueData: "Scribble for PowerPoint"; Flags: uninsdeletekey; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.PowerPointAddIn"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.PowerPointAddIn"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#PptProgId}"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PptClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\{#PptProgId}"; ValueType: string; ValueName: ""; ValueData: "Scribble for PowerPoint"; Flags: uninsdeletekey; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Classes\{#PptProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#PptClsid}"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Scribble"; Flags: uninsdeletekey; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Chat with your presentation. Scribble never saves, deletes, or sends."; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: IsWin64; Components: powerpoint
Root: HKCU64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#PptProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Check: IsWin64; Components: powerpoint

; Managed ActiveX control hosted as the Excel/PowerPoint sidebar (64-bit).
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}"; ValueType: string; ValueName: ""; ValueData: "Scribble Sidebar"; Flags: uninsdeletekey; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.UI.OfficeChatPane"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.UI.OfficeChatPane"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#OfficePaneProgId}"; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Implemented Categories\{#ControlCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#OfficePaneClsid}\Programmable"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\{#OfficePaneProgId}"; ValueType: string; ValueName: ""; ValueData: "Scribble Sidebar"; Flags: uninsdeletekey; Check: IsWin64; Components: excel powerpoint word
Root: HKCU64; Subkey: "Software\Classes\{#OfficePaneProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#OfficePaneClsid}"; Check: IsWin64; Components: excel powerpoint word


; Scribble for Word add-in (32-bit).
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}"; ValueType: string; ValueName: ""; ValueData: "Scribble for Word"; Flags: uninsdeletekey; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.WordAddIn"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.WordAddIn"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#WordProgId}"; Components: word
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#WordClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Components: word
Root: HKCU32; Subkey: "Software\Classes\{#WordProgId}"; ValueType: string; ValueName: ""; ValueData: "Scribble for Word"; Flags: uninsdeletekey; Components: word
Root: HKCU32; Subkey: "Software\Classes\{#WordProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#WordClsid}"; Components: word
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Scribble"; Flags: uninsdeletekey; Components: word
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Chat with your document. Scribble never saves, deletes, or sends."; Components: word
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Components: word
Root: HKCU32; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Components: word

; Scribble for Word add-in (64-bit).
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}"; ValueType: string; ValueName: ""; ValueData: "Scribble for Word"; Flags: uninsdeletekey; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.WordAddIn"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Class"; ValueData: "Scribble.WordAddIn"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\InprocServer32\1.1.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#WordProgId}"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#WordClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\{#WordProgId}"; ValueType: string; ValueName: ""; ValueData: "Scribble for Word"; Flags: uninsdeletekey; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Classes\{#WordProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#WordClsid}"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Scribble"; Flags: uninsdeletekey; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Chat with your document. Scribble never saves, deletes, or sends."; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: IsWin64; Components: word
Root: HKCU64; Subkey: "Software\Microsoft\Office\Word\Addins\{#WordProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Check: IsWin64; Components: word

[Icons]
Name: "{group}\Set up Scribble in Microsoft Edge"; Filename: "{app}\ScribbleBrowserHost.exe"; Parameters: "--setup edge"; Components: browser
Name: "{group}\Set up Scribble in Google Chrome"; Filename: "{app}\ScribbleBrowserHost.exe"; Parameters: "--setup chrome"; Components: browser

[Run]
Filename: "{app}\ScribbleBrowserHost.exe"; Parameters: "--setup auto"; Description: "Finish setting up Scribble in Edge or Chrome"; Flags: nowait postinstall skipifsilent; Components: browser

[Code]
function GetAssemblyCodeBase(Param: String): String;
var
  Path: String;
begin
  Path := ExpandConstant('{app}\Scribble.dll');
  StringChangeEx(Path, '\', '/', True);
  Result := 'file:///' + Path;
end;
