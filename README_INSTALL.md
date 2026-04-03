Instalador y releases (Windows)

Requisitos en la maquina de build
- Python 3.10+
- Inno Setup 6
- GitHub CLI (`gh`) autenticado

Preparar `.env`
- Debe incluir:
  - `SUPABASE_URL=...`
  - `SUPABASE_KEY=...`
  - `GOOGLE_SERVICE_ACCOUNT_FILE=...`
  - `INSTALLER_ASSET_NAME=RECA_INCLUSION_LABORAL_Setup.exe` (opcional)

Preparar `config.json`
- Debe existir al momento del build si la app usa IDs de Drive/Sheets definidos por configuracion.
- Desde este cambio se empaqueta dentro del release y queda dentro del bundle instalado.
- Para reparar una instalacion ya desplegada sin reinstalar, copia `config.json` a una de estas rutas:
  - `%APPDATA%\\RECA Inclusion Laboral\\config.json`
  - o la carpeta donde quede `RECA_INCLUSION_LABORAL.exe`

Notas
- `GOOGLE_SERVICE_ACCOUNT_FILE` puede ser ruta absoluta o relativa al directorio del `.env`.
- La ubicacion recomendada es `%APPDATA%\RECA Inclusion Laboral\.env`.
- El build del instalador toma `service-account.json` desde `GOOGLE_SERVICE_ACCOUNT_FILE` y lo copia a `%APPDATA%\RECA Inclusion Laboral\` durante la instalacion.
- Si se requiere una rotacion de credenciales, recompila y redistribuye el instalador para reemplazar ese archivo.

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
