; Inno Setup script — produces an installer and optional desktop shortcut
[Setup]
AppName=FisiereContorPagini
AppVersion=1.0
DefaultDirName={pf}\FisiereContorPagini
DefaultGroupName=FisiereContorPagini
OutputBaseFilename=FisiereContorPaginiSetup
Compression=lzma
SolidCompression=yes

[Files]
; adjust path to your built EXE
Source: "bin\Release\FisiereContorPagini.exe"; DestDir: "{app}"; Flags: ignoreversion

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; Flags: checkablealone

[Icons]
Name: "{group}\FisiereContorPagini"; Filename: "{app}\FisiereContorPagini.exe"; WorkingDir: "{app}"
Name: "{userdesktop}\FisiereContorPagini"; Filename: "{app}\FisiereContorPagini.exe"; Tasks: desktopicon; WorkingDir: "{app}"