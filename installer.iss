#define MyAppName "RECA Inclusion Laboral"
#ifndef MyAppVersion
  #define MyAppVersion "0.1.0"
#endif
#define MyAppPublisher "RECA"
#define MyAppExeName "RECA_INCLUSION_LABORAL.exe"

#include "installer_config.iss"

[Setup]
AppId={{8D9DB4D8-98CA-41E5-BC6A-B8F5167CFCA2}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=installer
OutputBaseFilename=RECA_INCLUSION_LABORAL_Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64
WizardStyle=modern

[Files]
Source: "dist\RECA_INCLUSION_LABORAL\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Crear icono en el escritorio"; GroupDescription: "Accesos directos:"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName}"; Flags: nowait postinstall

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  EnvPath: string;
  EnvContent: string;
begin
  if CurStep = ssInstall then
  begin
    EnvPath := ExpandConstant('{app}\.env');
    EnvContent := 'SUPABASE_URL={#SupabaseUrl}' + #13#10 +
                  'SUPABASE_KEY={#SupabaseKey}' + #13#10 +
                  'GITHUB_REPO_OWNER={#GithubRepoOwner}' + #13#10 +
                  'GITHUB_REPO_NAME={#GithubRepoName}' + #13#10 +
                  'INSTALLER_ASSET_NAME={#InstallerAssetName}' + #13#10;
    SaveStringToFile(EnvPath, EnvContent, False);
  end;
end;
