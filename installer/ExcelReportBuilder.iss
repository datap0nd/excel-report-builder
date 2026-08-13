#define AppName "PivotTable+"
#ifndef AppVersion
  #define AppVersion "0.2.2"
#endif
#define AppPublisher "PivotTable+ contributors"
#define AppProgId "ExcelReportBuilder.AddIn"
#define AppClsid "{{F953480C-A73C-4121-9E21-18676EC34CE8}"
#define PaneProgId "ExcelReportBuilder.TaskPaneHost"
#define PaneClsid "{{A3F4E10D-0DD1-420E-8B6F-E0A654BBEA16}"
#define ManagedCategory "{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
#define ControlCategory "{{40FC6ED4-2438-11CF-A3DB-080036F12502}"
#define AssemblyVersion "0.2.2.0"
#define AssemblyName "ExcelReportBuilder.AddIn, Version=0.2.2.0, Culture=neutral, PublicKeyToken=null"

[Setup]
AppId={{6A0B5710-1CD6-4F13-BE63-0E05B6860547}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion} (unsigned prototype)
AppPublisher={#AppPublisher}
DefaultDirName={localappdata}\Programs\ExcelReportBuilder
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x86os x64os
MinVersion=10.0.10240
OutputDir=..\artifacts
OutputBaseFilename=ExcelReportBuilderSetup-{#AppVersion}-unsigned
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
CloseApplicationsFilter=excel.exe
RestartApplications=no
UninstallDisplayName={#AppName} {#AppVersion} (unsigned prototype)
VersionInfoVersion={#AppVersion}
VersionInfoDescription=Unsigned prototype for enhancing native Excel PivotTables
VersionInfoProductName={#AppName}
VersionInfoCompany={#AppPublisher}
AppComments=Unsigned prototype. Verify the published SHA-256 checksum before installation.
SetupLogging=yes

[Files]
Source: "..\artifacts\sbom-payload\addin\*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\artifacts\sbom-payload\addin\*.config"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\artifacts\sbom-payload\worker-x64\*"; DestDir: "{app}\worker"; Flags: ignoreversion recursesubdirs createallsubdirs; Check: IsWin64
Source: "..\artifacts\sbom-payload\worker-x86\*"; DestDir: "{app}\worker"; Flags: ignoreversion recursesubdirs createallsubdirs; Check: not IsWin64
Source: "..\artifacts\sbom-payload\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\artifacts\sbom-payload\LICENSE"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; Excel add-in and managed COM classes for 32-bit Office.
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "ExcelReportBuilder.AddIn.Com.ExcelReportBuilderAddIn"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "Class"; ValueData: "ExcelReportBuilder.AddIn.Com.ExcelReportBuilderAddIn"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#AppProgId}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AppClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""
Root: HKCU32; Subkey: "Software\Classes\{#AppProgId}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\{#AppProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#AppClsid}"
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#AppProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "PivotTable+"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#AppProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Enhance a selected native Excel PivotTable."
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#AppProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"
Root: HKCU32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#AppProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"

Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}"; ValueType: string; ValueName: ""; ValueData: "{#AppName} task pane"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "ExcelReportBuilder.AddIn.Hosting.TaskPaneHost"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "Class"; ValueData: "ExcelReportBuilder.AddIn.Hosting.TaskPaneHost"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#PaneProgId}"
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Implemented Categories\{#ControlCategory}"; ValueType: string; ValueName: ""; ValueData: ""
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Programmable"; ValueType: string; ValueName: ""; ValueData: ""
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Control"; ValueType: string; ValueName: ""; ValueData: ""
Root: HKCU32; Subkey: "Software\Classes\{#PaneProgId}"; ValueType: string; ValueName: ""; ValueData: "{#AppName} task pane"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\{#PaneProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#PaneClsid}"

; Matching 64-bit registration for 64-bit Office.
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "ExcelReportBuilder.AddIn.Com.ExcelReportBuilderAddIn"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "Class"; ValueData: "ExcelReportBuilder.AddIn.Com.ExcelReportBuilderAddIn"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#AppProgId}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AppClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\{#AppProgId}"; ValueType: string; ValueName: ""; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\{#AppProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#AppClsid}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#AppProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "PivotTable+"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#AppProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Enhance a selected native Excel PivotTable."; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#AppProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#AppProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Check: IsWin64

Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}"; ValueType: string; ValueName: ""; ValueData: "{#AppName} task pane"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "ExcelReportBuilder.AddIn.Hosting.TaskPaneHost"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "Class"; ValueData: "ExcelReportBuilder.AddIn.Hosting.TaskPaneHost"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AssemblyName}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\InprocServer32\{#AssemblyVersion}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetAssemblyCodeBase}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#PaneProgId}"; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Implemented Categories\{#ManagedCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Implemented Categories\{#ControlCategory}"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Programmable"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#PaneClsid}\Control"; ValueType: string; ValueName: ""; ValueData: ""; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\{#PaneProgId}"; ValueType: string; ValueName: ""; ValueData: "{#AppName} task pane"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\{#PaneProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#PaneClsid}"; Check: IsWin64

[Code]
const
  DotNet48Release = 528040;
  InternetMaxUrlLength = 2083;
  S_OK = 0;

function UrlCreateFromPathW(pszPath, pszUrl: string; var pcchUrl: DWORD;
  dwFlags: DWORD): HResult;
  external 'UrlCreateFromPathW@shlwapi.dll stdcall';

function IsDotNet48Installed: Boolean;
var
  Release: Cardinal;
begin
  Release := 0;
  if IsWin64 then
    Result := RegQueryDWordValue(HKLM64,
      'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full',
      'Release', Release)
  else
    Result := RegQueryDWordValue(HKLM32,
      'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full',
      'Release', Release);

  Result := Result and (Release >= DotNet48Release);
end;

function InitializeSetup: Boolean;
begin
  Result := IsDotNet48Installed;
  if not Result then
    MsgBox('.NET Framework 4.8 or newer is required. Install it through Windows Update, then run setup again.',
      mbError, MB_OK);
end;

function GetAssemblyCodeBase(Param: String): String;
var
  Path: String;
  Url: String;
  CharacterCount: DWORD;
begin
  Path := ExpandConstant('{app}\ExcelReportBuilder.AddIn.dll');
  CharacterCount := InternetMaxUrlLength;
  SetLength(Url, CharacterCount);
  if UrlCreateFromPathW(Path, Url, CharacterCount, 0) <> S_OK then
    RaiseException('Could not create the managed assembly CodeBase URL.');

  SetLength(Url, CharacterCount);
  Result := Url;
end;
