[Setup]
AppId=TimetableScheduler
AppName=空课与排班
AppVersion=1.4.1
SetupIconFile=resources\app.ico
DefaultDirName={localappdata}\Programs\TimetableScheduler
DisableDirPage=no
PrivilegesRequired=lowest
OutputDir=.
OutputBaseFilename=TimetableScheduler-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\TimetableScheduler.exe

[Files]
Source: "TimetableScheduler.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "resources\*"; DestDir: "{app}\resources"; Flags: ignoreversion recursesubdirs
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion

[InstallDelete]
Type: files; Name: "{app}\installed.flag"

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    SaveStringToFile(ExpandConstant('{app}\installed.flag'), 'installed', False);
end;

[Icons]
Name: "{userprograms}\空课与排班"; Filename: "{app}\TimetableScheduler.exe"
Name: "{userdesktop}\空课与排班"; Filename: "{app}\TimetableScheduler.exe"

[Run]
Filename: "{app}\TimetableScheduler.exe"; Description: "启动空课与排班"; Flags: nowait postinstall skipifsilent
