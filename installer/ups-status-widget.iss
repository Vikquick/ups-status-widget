#ifndef MyAppVersion
#define MyAppVersion "2.0.0"
#endif

[Setup]
AppId={{A58E2B50-824E-4F89-9714-8B76AB95C0D6}
AppName=UPS Status Widget
AppVersion={#MyAppVersion}
AppPublisher=UPS Status Widget Contributors
DefaultDirName={autopf}\UPS Status Widget
DefaultGroupName=UPS Status Widget
DisableDirPage=no
DisableProgramGroupPage=no
OutputDir=..\publish\installer
OutputBaseFilename=UPS-Status-Widget-Setup-{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\UPS-Status-Widget.exe

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "..\publish\x64\UPS-Status-Widget.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\UPS Status Widget"; Filename: "{app}\UPS-Status-Widget.exe"
Name: "{autodesktop}\UPS Status Widget"; Filename: "{app}\UPS-Status-Widget.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\UPS-Status-Widget.exe"; Description: "Launch UPS Status Widget"; Flags: nowait postinstall skipifsilent
