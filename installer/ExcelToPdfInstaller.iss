; ---------------------------------------------------------------
; Inno Setup script for CheapNFriendly Excel to PDF Converter
; Builds a single-file Windows installer (setup.exe) with the
; produced PyInstaller executable and application icon.
; ---------------------------------------------------------------

#define MyAppName "CheapNFriendly Excel to PDF Converter"
#define MyAppShortName "Excel to PDF Converter"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "CheapNFriendly"
#define MyAppURL "https://cheapnfriendly.example"
#define MyAppExeName "ExcelToPdfConverter.exe"

[Setup]
AppId={{7C1F2A94-5B3D-4E8A-9C21-6D4B8F0E1A23}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\CheapNFriendly\ExcelToPdfConverter
DefaultGroupName=CheapNFriendly
DisableProgramGroupPage=yes
; Installer output
OutputDir=.
OutputBaseFilename=CheapNFriendly-ExcelToPdf-Setup-{#MyAppVersion}
; Icon for the installer itself and for Add/Remove Programs
SetupIconFile=..\icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
; Require admin so we can install into Program Files
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; The PyInstaller one-file executable
Source: "..\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; Ship the icon alongside for reference/shortcuts
Source: "..\icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppShortName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\icon.ico"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppShortName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Remove persisted queue data and app caches
Type: filesandordirs; Name: "{localappdata}\CheapNFriendly"
