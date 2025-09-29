; Inno Setup script for PartsMaster

[Setup]
AppName=Parts Master
AppVersion=0.0.2
AppPublisher=In.Sist d.o.o.
DefaultDirName={pf}\Parts Master
DefaultGroupName=Parts Master
OutputDir=dist
OutputBaseFilename=PartsMasterInstaller
Compression=lzma
SolidCompression=yes
SetupIconFile=assets\official-logo.ico

[Files]
; Main executable (renamed without space)
Source: "dist\PartsMaster.exe"; DestDir: "{app}"; Flags: ignoreversion

; (Optional) external assets (only if you didn’t bundle with PyInstaller)
; Source: "assets\*"; DestDir: "{app}\assets"; Flags: recursesubdirs createallsubdirs ignoreversion
; Source: "app_info.json"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Parts Master"; Filename: "{app}\PartsMaster.exe"
Name: "{commondesktop}\Parts Master"; Filename: "{app}\PartsMaster.exe"

[Run]
Filename: "{app}\PartsMaster.exe"; Description: "Launch Parts Master"; Flags: nowait postinstall skipifsilent
