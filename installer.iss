#define MyAppName "RECA Inclusion Laboral"
#ifndef MyAppVersion
  #define MyAppVersion "0.1.0"
#endif
#define MyAppPublisher "RECA"
#define MyAppExeName "RECA_INCLUSION_LABORAL.exe"

#ifnexist "installer_config.local.iss"
  #error "installer_config.local.iss no encontrado. Ejecuta build.ps1 antes de compilar el instalador."
#endif
#include "installer_config.local.iss"

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
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern

[Files]
Source: "dist\RECA_INCLUSION_LABORAL\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#GoogleServiceAccountSourcePath}"; DestDir: "{userappdata}\RECA Inclusion Laboral"; DestName: "{#GoogleServiceAccountFileName}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Crear icono en el escritorio"; GroupDescription: "Accesos directos:"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName}"; Flags: nowait postinstall skipifsilent
Filename: "{app}\{#MyAppExeName}"; Flags: nowait runasoriginaluser; Check: WizardSilent

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  RoamingPath: string;
  EnvContent: string;
begin
  if CurStep = ssPostInstall then
  begin
    EnvContent := 'SUPABASE_URL={#SupabaseUrl}' + #13#10 +
                  'SUPABASE_KEY={#SupabaseKey}' + #13#10 +
                  'INSTALLER_ASSET_NAME={#InstallerAssetName}' + #13#10 +
                  'GOOGLE_SERVICE_ACCOUNT_FILE={#GoogleServiceAccountFileName}' + #13#10;

    RoamingPath := ExpandConstant('{userappdata}\RECA Inclusion Laboral\.env');
    ForceDirectories(ExtractFileDir(RoamingPath));
    SaveStringToFile(RoamingPath, EnvContent, False);
  end;
end;
