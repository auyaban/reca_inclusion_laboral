$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$venvPath = Join-Path $root ".venv"
if (!(Test-Path $venvPath)) {
    python -m venv $venvPath
}

$python = Join-Path $venvPath "Scripts\python.exe"
& $python -m pip install --upgrade pip
& $python -m pip install -r requirements.txt
& $python -m pip install pyinstaller

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

$configPath = Join-Path $root "config.json"
$config = $null
if (Test-Path $configPath) {
    try {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
    } catch {
        throw "config.json invÃ¡lido: $($_.Exception.Message)"
    }
}

function Get-ConfigValue {
    param(
        $Config,
        [string]$Key
    )

    if ($null -eq $Config) {
        return ""
    }
    $property = $Config.PSObject.Properties[$Key]
    if ($null -eq $property) {
        return ""
    }
    return [string]$property.Value
}

$envPath = Join-Path $root ".env"
if (!(Test-Path $envPath)) {
    throw ".env no encontrado"
}
$envValues = Get-EnvMap (Get-Content $envPath)
$supabaseUrl = [string]($envValues["SUPABASE_URL"])
$supabaseKey = [string]($envValues["SUPABASE_KEY"])
$repoOwner = [string]($envValues["GITHUB_REPO_OWNER"])
$repoName = [string]($envValues["GITHUB_REPO_NAME"])
$installerAsset = [string]($envValues["INSTALLER_ASSET_NAME"])
if (-not $installerAsset) { $installerAsset = "RECA_INCLUSION_LABORAL_Setup.exe" }

$googleServiceAccount = [string]($envValues["GOOGLE_SERVICE_ACCOUNT_FILE"])
if (-not $googleServiceAccount) {
    $googleServiceAccount = Get-ConfigValue $config "google_service_account_file"
}
if (-not $googleServiceAccount) {
    $googleServiceAccount = Get-ConfigValue $config "google_sheets_sa_json"
}
if (-not $googleServiceAccount) {
    $googleServiceAccount = Get-ConfigValue $config "google_drive_sa_json"
}
if (-not $googleServiceAccount) {
    throw "Falta GOOGLE_SERVICE_ACCOUNT_FILE o config.json con google_service_account_file/google_sheets_sa_json/google_drive_sa_json."
}
if (-not [System.IO.Path]::IsPathRooted($googleServiceAccount)) {
    $googleServiceAccount = Join-Path $root $googleServiceAccount
}
$googleServiceAccountPath = (Resolve-Path $googleServiceAccount).Path
$googleServiceAccountFileName = [System.IO.Path]::GetFileName($googleServiceAccountPath)

$installerConfig = @"
#define SupabaseUrl "$supabaseUrl"
#define SupabaseKey "$supabaseKey"
#define GithubRepoOwner "$repoOwner"
#define GithubRepoName "$repoName"
#define InstallerAssetName "$installerAsset"
#define GoogleServiceAccountFileName "$googleServiceAccountFileName"
"@
Set-Content -Path (Join-Path $root "installer_config.local.iss") -Value $installerConfig -Encoding UTF8

$pyiArgs = @(
    "--noconfirm",
    "--clean",
    "--windowed",
    "--name", "RECA_INCLUSION_LABORAL",
    "--add-data", "templates;templates",
    "--add-data", "Diccionario.txt;.",
    "--add-data", "VERSION;.",
    "--add-data", "config.json;.",
    "--add-data", "$googleServiceAccountPath;.",
    "--hidden-import", "win32com",
    "--hidden-import", "win32com.client",
    "--hidden-import", "pythoncom",
    "--hidden-import", "pywintypes",
    "--hidden-import", "win32timezone",
    "app.py"
)

& $python -m PyInstaller @pyiArgs
