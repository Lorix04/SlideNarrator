#define MyAppName "SlideNarrator"
#define MyAppVersion "2.9.0"
#define MyAppPublisher "Lorix04"
#define MyAppURL "https://github.com/Lorix04/SlideNarrator"
#define MyAppExeName "SlideNarrator.exe"

[Setup]
AppId={{A5182182-EB31-4CDB-A74B-88CE1C556471}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=output
OutputBaseFilename=SlideNarrator_Setup_{#MyAppVersion}
SetupIconFile=..\assets\slide_narrator_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
VersionInfoVersion={#MyAppVersion}.0
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription=Installer ufficiale di SlideNarrator
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}.0

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Crea un collegamento sul desktop"; GroupDescription: "Collegamenti aggiuntivi:"

[Files]
Source: "..\dist\SlideNarrator\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\SlideNarrator"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\SlideNarrator"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Avvia SlideNarrator"; Flags: nowait postinstall skipifsilent
