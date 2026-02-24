Instalador y releases (Windows)

Requisitos en la maquina de build
- Python 3.10+
- Inno Setup 6
- GitHub CLI (`gh`) autenticado

Preparar `.env`
- Debe incluir:
  - `SUPABASE_URL=...`
  - `SUPABASE_KEY=...`
  - `GITHUB_REPO_OWNER=...`
  - `GITHUB_REPO_NAME=...`
  - `INSTALLER_ASSET_NAME=RECA_INCLUSION_LABORAL_Setup.exe` (opcional)

Build local
1) `powershell -ExecutionPolicy Bypass -File build.ps1`
2) Ejecutable generado en: `dist\RECA_INCLUSION_LABORAL\RECA_INCLUSION_LABORAL.exe`

Installer
1) Compilar `installer.iss` con Inno Setup
2) Instalador generado en: `installer\RECA_INCLUSION_LABORAL_Setup.exe`

Release automatizado
1) `powershell -ExecutionPolicy Bypass -File release.ps1 vX.Y.Z`
2) Publica en GitHub Release:
   - `RECA_INCLUSION_LABORAL_Setup.exe`
   - `RECA_INCLUSION_LABORAL_Setup.exe.sha256`

Actualizacion desde la app
- En el Hub: boton `Actualizar aplicación`.
- Consulta `releases/latest` del repo configurado.
- Si hay nueva version:
  - descarga instalador,
  - valida SHA256 (si existe asset `.sha256`),
  - instala en silencio,
  - reinicia la app.
