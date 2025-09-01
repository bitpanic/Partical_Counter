; Inno Setup script for SPS30 Tray Logger
; Requires Inno Setup (ISCC.exe) to be installed and on PATH

[Setup]
AppId={{B5E56B4F-5E5B-4A9B-9C4B-3E2C4C7E0821}
AppName=SPS30 Tray Logger
AppVersion=1.0.0
AppPublisher=Your Organization
DefaultDirName={autopf}\\SPS30 Tray Logger
DefaultGroupName=SPS30 Tray Logger
DisableDirPage=no
DisableProgramGroupPage=no
OutputDir=Output
OutputBaseFilename=SPS30TrayLoggerSetup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
; Built executable from PyInstaller
Source: "..\\dist\\sps30-tray-logger.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\\SPS30 Tray Logger"; Filename: "{app}\\sps30-tray-logger.exe"
Name: "{commondesktop}\\SPS30 Tray Logger"; Filename: "{app}\\sps30-tray-logger.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\\sps30-tray-logger.exe"; Description: "Launch SPS30 Tray Logger"; Flags: nowait postinstall skipifsilent


