param(
    [switch]$Clean,
    [switch]$ForceDependencyInstall
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$venvPath = Join-Path $root ".venv"
$venvCreated = $false
if (!(Test-Path $venvPath)) {
    python -m venv $venvPath
    $venvCreated = $true
}

$python = Join-Path $venvPath "Scripts\python.exe"
$requirementsPath = Join-Path $root "requirements.txt"
$requirementsHashPath = Join-Path $venvPath ".requirements.sha256"

function Get-FileSha256 {
    param([Parameter(Mandatory = $true)][string]$Path)

    return (Get-FileHash -Algorithm SHA256 $Path).Hash.ToLowerInvariant()
}

function Test-PythonModule {
    param([Parameter(Mandatory = $true)][string]$ModuleName)

    & $python -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('$ModuleName') else 1)"
    return ($LASTEXITCODE -eq 0)
}

function Write-Utf8NoBomFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Value
    )

    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Value, $encoding)
}

function Repair-TkcalendarSyntaxWarning {
    param([Parameter(Mandatory = $true)][string]$VenvPath)

    $calendarPath = Join-Path $VenvPath "Lib\site-packages\tkcalendar\calendar_.py"
    if (!(Test-Path $calendarPath)) {
        return
    }

    $original = Get-Content $calendarPath -Raw
    $needle = '"Liberation\ Sans 9"'
    $replacement = '"Liberation\\ Sans 9"'
    if ($original.Contains($needle)) {
        Write-Utf8NoBomFile -Path $calendarPath -Value ($original.Replace($needle, $replacement))
        Write-Host "tkcalendar calendar_.py normalizado para evitar SyntaxWarning."
    }
}

$requirementsHash = Get-FileSha256 $requirementsPath
$cachedRequirementsHash = ""
if (Test-Path $requirementsHashPath) {
    $cachedRequirementsHash = (Get-Content $requirementsHashPath | Select-Object -First 1).Trim().ToLowerInvariant()
}

$needsDependencyInstall = $ForceDependencyInstall `
    -or $venvCreated `
    -or ($cachedRequirementsHash -ne $requirementsHash) `
    -or -not (Test-PythonModule "PyInstaller")

if ($needsDependencyInstall) {
    & $python -m pip install --upgrade pip
    & $python -m pip install -r requirements.txt
    & $python -m pip install pyinstaller
    Set-Content -Path $requirementsHashPath -Value "$requirementsHash`n" -Encoding utf8
} else {
    Write-Host "Dependencias sin cambios; se reutiliza la .venv."
}

Repair-TkcalendarSyntaxWarning -VenvPath $venvPath

function Get-EnvMap {
    param([string[]]$Lines)

    $values = @{}
    foreach ($raw in $Lines) {
        if ([string]::IsNullOrWhiteSpace($raw)) {
            continue
        }
        $line = $raw.Trim()
        if ($line.StartsWith("#")) {
            continue
        }
        $separator = $line.IndexOf("=")
        if ($separator -lt 1) {
            continue
        }
        $key = $line.Substring(0, $separator).Trim()
        $value = $line.Substring($separator + 1).Trim()
        $values[$key] = $value
    }
    return $values
}

function Resolve-EnvPath {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [string]$EnvName = ".env"
    )

    $candidates = @()
    $roamingAppData = [Environment]::GetFolderPath("ApplicationData")
    if (-not [string]::IsNullOrWhiteSpace($roamingAppData)) {
        $candidates += Join-Path (Join-Path $roamingAppData "RECA Inclusion Laboral") $EnvName
    }
    $candidates += Join-Path $Root $EnvName

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    return ""
}

$envPath = Resolve-EnvPath -Root $root
if (!(Test-Path $envPath)) {
    throw ".env no encontrado. Colocalo en %APPDATA%\\RECA Inclusion Laboral\\.env o junto a build.ps1."
}
$envValues = Get-EnvMap (Get-Content $envPath)
$supabaseUrl = [string]($envValues["SUPABASE_URL"])
$supabaseKey = [string]($envValues["SUPABASE_KEY"])
$installerAsset = [string]($envValues["INSTALLER_ASSET_NAME"])
if (-not $installerAsset) { $installerAsset = "RECA_INCLUSION_LABORAL_Setup.exe" }

$googleServiceAccount = [string]($envValues["GOOGLE_SERVICE_ACCOUNT_FILE"])
if (-not $googleServiceAccount) {
    $googleServiceAccount = [string]($envValues["GOOGLE_DRIVE_SA_JSON"])
}
if (-not $googleServiceAccount) {
    $googleServiceAccount = "service-account.json"
}
if (-not [System.IO.Path]::IsPathRooted($googleServiceAccount)) {
    $googleServiceAccount = Join-Path (Split-Path -Parent $envPath) $googleServiceAccount
}
$googleServiceAccountPath = (Resolve-Path $googleServiceAccount).Path
$googleServiceAccountFileName = [System.IO.Path]::GetFileName($googleServiceAccountPath)
$googleServiceAccountSourcePath = $googleServiceAccountPath.Replace('"', '""')

$installerConfig = @"
#define SupabaseUrl "$supabaseUrl"
#define SupabaseKey "$supabaseKey"
#define InstallerAssetName "$installerAsset"
#define GoogleServiceAccountFileName "$googleServiceAccountFileName"
#define GoogleServiceAccountSourcePath "$googleServiceAccountSourcePath"
"@
Set-Content -Path (Join-Path $root "installer_config.local.iss") -Value $installerConfig -Encoding UTF8

$requiredBuildInputs = @(
    "config.json",
    "Diccionario.txt",
    "VERSION",
    "templates"
)

foreach ($relativePath in $requiredBuildInputs) {
    $candidate = Join-Path $root $relativePath
    if (!(Test-Path $candidate)) {
        throw "Falta recurso requerido para el build: $relativePath"
    }
}

# Validar claves obligatorias dentro de config.json
$configPath = Join-Path $root "config.json"
$configContent = Get-Content -Path $configPath -Raw | ConvertFrom-Json
$requiredConfigKeys = @(
    "google_drive_folder_id",
    "google_sheets_master_template_id",
    "pdf_export_folder_id"
)
foreach ($key in $requiredConfigKeys) {
    if (-not ($configContent.PSObject.Properties.Name -contains $key) -or
        [string]::IsNullOrWhiteSpace($configContent.$key)) {
        throw "config.json falta la clave obligatoria: '$key'. Consulta config.example.json."
    }
}
Write-Host "config.json validado: todas las claves obligatorias presentes."

$pyiArgs = @(
    "--noconfirm",
    "--windowed",
    "--name", "RECA_INCLUSION_LABORAL",
    "--additional-hooks-dir", "pyinstaller_hooks",
    "--add-data", "config.json;.",
    "--add-data", "templates;templates",
    "--add-data", "Diccionario.txt;.",
    "--add-data", "VERSION;.",
    "--hidden-import", "win32com",
    "--hidden-import", "win32com.client",
    "--hidden-import", "pythoncom",
    "--hidden-import", "pywintypes",
    "--hidden-import", "win32timezone",
    "app.py"
)

if ($Clean) {
    $pyiArgs = @("--clean") + $pyiArgs
}

& $python -m PyInstaller @pyiArgs

$distRoot = Join-Path $root "dist\RECA_INCLUSION_LABORAL"
$runtimeContentRoot = Join-Path $distRoot "_internal"
if (!(Test-Path $runtimeContentRoot)) {
    $runtimeContentRoot = $distRoot
}

$requiredRuntimePaths = @(
    "config.json",
    "Diccionario.txt",
    "VERSION",
    "templates"
)

foreach ($relativePath in $requiredRuntimePaths) {
    $candidate = Join-Path $runtimeContentRoot $relativePath
    if (!(Test-Path $candidate)) {
        throw "Build incompleto: falta $relativePath en $runtimeContentRoot"
    }
}

Write-Host "Build verificado: recursos runtime presentes en $runtimeContentRoot."
