; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "pdfUtils"
#define MyAppVersion "0.0"
#define MyAppPublisher "MTMSD"
#define MyAppExeName "pdfUtils.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{AE8E96FB-5056-46B0-8E61-59730A1AB7CF}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={pf}\{#MyAppName}
DisableProgramGroupPage=yes
ChangesAssociations=yes
OutputDir=D:\TRABAJO\DEV_PDF\INSTALLER
OutputBaseFilename=pdfUtils Installer
;SetupIconFile=D:\TRABAJO\DEV_RIMSA\rimsa.ico
Compression=lzma
SolidCompression=yes
;UninstallDisplayIcon={app}\RIMSA.exe

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
;Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Dirs]
Name: "{app}"; Permissions: users-full
;Name: "{app}\catalogo"; Permissions: users-full
;Name: "{app}\icons"; Permissions: users-full

[Files]
Source: "D:\TRABAJO\DEV_PDF\dist\pdfUtils\*"; DestDir: "{app}"; Flags: ignoreversion
Source: "D:\TRABAJO\DEV_PDF\dist\pdfUtils\Include\*"; DestDir: "{app}\Include"; Flags: recursesubdirs createallsubdirs
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{commonprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
;Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Registry]
Root: HKCR; Subkey: "*\shell\PDF Utils"; Flags: uninsdeletekey
Root: HKCR; Subkey: "*\shell\PDF Utils"; ValueType: string; ValueName: "MUIVerb"; ValueData: "PDF Utils"; Flags: uninsdeletevalue
Root: HKCR; Subkey: "*\shell\PDF Utils"; ValueType: string; ValueName: "SubCommands"; ValueData: "PDF.Merge;PDF.Split;PDF.Splitpng"; Flags: uninsdeletevalue
Root: HKLM64; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\PDF.Split"; ValueType: string; ValueName: ""; ValueData: "Split PDF"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\PDF.Merge"; ValueType: string; ValueName: ""; ValueData: "Merge PDF"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\PDF.Splitpng"; ValueType: string; ValueName: ""; ValueData: "Split PDF to PNG"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\PDF.Split\command"; ValueType: string; ValueName: ""; ValueData: """{app}\pdfUtils.exe"" ""split"""; Flags: uninsdeletekey
Root: HKLM64; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\PDF.Merge\command"; ValueType: string; ValueName: ""; ValueData: """{app}\pdfUtils.exe"" ""merge"""; Flags: uninsdeletekey
Root: HKLM64; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\shell\PDF.Splitpng\command"; ValueType: string; ValueName: ""; ValueData: """{app}\pdfUtils.exe"" ""splitpng"""; Flags: uninsdeletekey

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent


