[Setup]
AppName=PPK Batch Processor
AppVersion=1.0.2
AppVerName=PPK Batch Processor 1.0.2
DefaultDirName={autopf}\PPK Batch Processor
DefaultGroupName=PPK Batch Processor
OutputDir=Output
OutputBaseFilename=PPK_Batch_Processor_Setup_1_0_2
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
UninstallDisplayIcon={app}\icon.ico
SetupIconFile=assets\icon.ico
CloseApplications=force
RestartApplications=no

[Files]
Source: "dist\PPK_Batch_Processor.exe"; DestDir: "{app}"; Flags: ignoreversion replacesameversion
Source: "config.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\icon.ico"; DestDir: "{app}"; Flags: ignoreversion replacesameversion

[Icons]
Name: "{group}\PPK Batch Processor"; Filename: "{app}\PPK_Batch_Processor.exe"; IconFilename: "{app}\icon.ico"; Flags: uninsnevernouninstall
Name: "{commondesktop}\PPK Batch Processor"; Filename: "{app}\PPK_Batch_Processor.exe"; IconFilename: "{app}\icon.ico"; Flags: uninsnevernouninstall

[Run]
Filename: "{app}\PPK_Batch_Processor.exe"; Description: "Lancer PPK Batch Processor"; Flags: postinstall nowait

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Tasks]
Name: "desktopicon"; Description: "Créer un raccourci sur le bureau"; GroupDescription: "Raccourcis:"

[Code]
function DirExists(const Name: string): Boolean;
begin
  Result := DirExists(ExpandConstant('{#SourcePath}\' + Name));
end;