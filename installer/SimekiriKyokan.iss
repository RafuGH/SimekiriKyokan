[Setup]
AppId={{8F6A3D52-9E21-4C1E-ABCD-1234567890AB}}
AppName=締切教官 (SimekiriKyokan)
AppVersion=2.0
AppPublisher=Rafu
DefaultDirName={autopf}\SimekiriKyokan
DefaultGroupName=締切教官
OutputBaseFilename=SimekiriKyokan_Setup_v2.0
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; Flags: unchecked

[Files]
Source: "dist\simekiri_gui\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

Source: "data\Tasks.xlsx"; DestDir: "{app}"; Flags: ignoreversion

Source: "data\SimekiriKyokan_Manual.pdf"; DestDir: "{app}"; Flags: ignoreversion

Source: "data\icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\締切教官"; Filename: "{app}\simekiri_gui.exe"; IconFilename: "{app}\icon.ico"
Name: "{autodesktop}\締切教官"; Filename: "{app}\simekiri_gui.exe"; Tasks: desktopicon; IconFilename: "{app}\icon.ico"

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  AppDataDir, SourceFile, DestFile: string;
begin
  if CurStep = ssPostInstall then
  begin
    AppDataDir := ExpandConstant('{userappdata}\SimekiriKyokan');
    if not DirExists(AppDataDir) then
      CreateDir(AppDataDir);

    SourceFile := ExpandConstant('{app}\Tasks.xlsx');
    DestFile := AppDataDir + '\Tasks.xlsx';

    if not FileExists(DestFile) then
      CopyFile(SourceFile, DestFile, False);
  end;
end;

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
