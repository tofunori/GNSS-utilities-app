[Setup]
AppName=PPK Batch Processor
AppVersion=1.1
WizardStyle=modern
DefaultDirName={autopf}\PPK Batch Processor
DefaultGroupName=PPK Batch Processor
OutputDir=Output
OutputBaseFilename=PPK_Batch_Processor_Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
AppPublisher=Votre Entreprise
AppPublisherURL=https://votre-site.com
AppSupportURL=https://votre-site.com/support
AppUpdatesURL=https://votre-site.com/updates
VersionInfoCompany=Votre Entreprise
VersionInfoCopyright=© 2024 Votre Entreprise
VersionInfoDescription=PPK Batch Processor Installation

[Files]
Source: "dist\PPK_Batch_Processor\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs
Source: "README.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\PPK Batch Processor"; Filename: "{app}\PPK_Batch_Processor.exe"
Name: "{commondesktop}\PPK Batch Processor"; Filename: "{app}\PPK_Batch_Processor.exe"

[Run]
Filename: "{app}\PPK_Batch_Processor.exe"; Description: "Lancer PPK Batch Processor"; Flags: nowait postinstall skipifsilent
Filename: "notepad.exe"; Parameters: "{app}\README.txt"; Description: "Lire les instructions d'installation"; Flags: postinstall nowait skipifsilent

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Tasks]
Name: "desktopicon"; Description: "Créer un raccourci sur le bureau"; GroupDescription: "Raccourcis:"

[Code]
var
  ResultCode: Integer;

function DirExists(const Name: string): Boolean;
begin
  Result := DirExists(ExpandConstant('{#SourcePath}\' + Name));
end;

procedure AddWindowsDefenderExclusion();
var
  ExecPath: string;
begin
  ExecPath := ExpandConstant('{app}');
  Exec('powershell.exe', '-Command Add-MpPreference -ExclusionPath "' + ExecPath + '"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    if MsgBox('Voulez-vous ajouter le dossier d''installation aux exclusions de Windows Defender ?', 
      mbConfirmation, MB_YESNO) = IDYES then
    begin
      AddWindowsDefenderExclusion();
    end;
  end;
end;