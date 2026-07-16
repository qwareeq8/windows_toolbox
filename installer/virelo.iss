#define MyAppName "Virelo"
#define MyAppPublisher "Yusuf Qwareeq"
#define MyAppURL "https://github.com/qwareeq8/virelo"
#define MyAppExeName "Virelo.exe"
#ifndef MyAppVersion
  #define MyAppVersion "0.0.0-dev"
#endif

[Setup]
AppId={{7B84D572-0B2B-4C1B-9B45-0F2CFAB58A31}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}/releases
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=no
OutputDir=dist
OutputBaseFilename={#MyAppName}Setup
SetupIconFile={#SourcePath}\..\icon.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
WizardImageFile={#SourcePath}\..\branding\installer-wizard.bmp,{#SourcePath}\..\branding\installer-wizard_2x.bmp
WizardSmallImageFile={#SourcePath}\..\branding\installer-header.bmp,{#SourcePath}\..\branding\installer-header_2x.bmp
WizardImageAlphaFormat=none
; Qt 6.11 supports Windows 10 version 1809 (build 17763) or later.
MinVersion=10.0.17763
ArchitecturesAllowed=x64os
ArchitecturesInstallIn64BitMode=x64os
PrivilegesRequired=admin
AppMutex=Local\Virelo_Mutex,Global\Virelo_Mutex
CloseApplications=yes
RestartApplications=no
UninstallDisplayIcon={app}\{#MyAppExeName}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} Setup
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}
VersionInfoVersion={#MyAppVersion}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "{#SourcePath}\.\..\dist\Virelo\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[InstallDelete]
; Replace the complete PyInstaller-owned payload so an upgrade cannot retain
; removed modules or native libraries from an earlier one-folder build.
Type: filesandordirs; Name: "{app}\_internal"
Type: files; Name: "{app}\Virelo.exe"
Type: files; Name: "{app}\.release.json"
Type: files; Name: "{app}\bundle-files.sha256"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
