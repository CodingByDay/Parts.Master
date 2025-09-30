; Inno Setup script for Parts Master

[Setup]
AppName=Parts Master
AppVersion=0.0.7
AppPublisher=In.Sist d.o.o.
DefaultDirName={pf}\Parts Master
DefaultGroupName=Parts Master
OutputDir=dist
OutputBaseFilename=PartsMasterInstaller
Compression=lzma
SolidCompression=yes
SetupIconFile=assets\favicon.ico

; 👇 Ensure Add/Remove Programs shows correct name and icon
UninstallDisplayName=Parts Master
UninstallDisplayIcon={app}\PartsMaster.exe

[Files]
; Main executable
Source: "dist\PartsMaster.exe"; DestDir: "{app}"; Flags: ignoreversion

; Optional: external assets (only needed if not bundled into exe)
; Source: "assets\*"; DestDir: "{app}\assets"; Flags: recursesubdirs createallsubdirs ignoreversion
; Source: "app_info.json"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu shortcut
Name: "{group}\Parts Master"; Filename: "{app}\PartsMaster.exe"
; Desktop shortcut
Name: "{commondesktop}\Parts Master"; Filename: "{app}\PartsMaster.exe"

