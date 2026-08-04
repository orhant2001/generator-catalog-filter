; Inno Setup script for GeneratorCatalogFilter
; Packages the PyInstaller onedir output into a single setup.exe.

#define MyAppName "Generator Catalog Filter"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "orhant2001"
#define MyAppExeName "GeneratorCatalogFilter.exe"

[Setup]
AppId={{A7F3C2E1-9B4D-4E6A-8C1F-2D5B7E9A0C34}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\GeneratorCatalogFilter
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename=GeneratorCatalogFilter-Setup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
; Installs per-user without requiring admin rights.
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "turkish"; MessagesFile: "compiler:Languages\Turkish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Take the entire PyInstaller onedir folder (exe + _internal) and install it.
Source: "dist\GeneratorCatalogFilter\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent
