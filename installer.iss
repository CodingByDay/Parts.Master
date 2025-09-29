; Inno Setup script for Parts Master

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

; Optional: create desktop and start menu shortcuts
SetupIconFile=assets\official-logo.ico

[Files]
; Main executable
Source: "dist\Parts Master.exe"; DestDir: "{app}"; Flags: ignoreversion

; Include assets folder
Source: "assets\*"; DestDir: "{app}\assets"; Flags: recursesubdirs createallsubdirs ignoreversion

; Include app_info.json
Source: "app_info.json"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Parts Master"; Filename: "{app}\Parts Master.exe"
Name: "{commondesktop}\Parts Master"; Filename: "{app}\Parts Master.exe"

[Run]
Filename: "{app}\Parts Master.exe"; Description: "Launch Parts Master"; Flags: nowait postinstall skipifsilent
