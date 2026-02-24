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

$envPath = Join-Path $root ".env"
if (!(Test-Path $envPath)) {
    throw ".env no encontrado"
}
$envLines = Get-Content $envPath
$supabaseUrl = (($envLines | Where-Object { $_ -match '^SUPABASE_URL=' }) -replace '^SUPABASE_URL=', '').Trim()
$supabaseKey = (($envLines | Where-Object { $_ -match '^SUPABASE_KEY=' }) -replace '^SUPABASE_KEY=', '').Trim()
$repoOwner = (($envLines | Where-Object { $_ -match '^GITHUB_REPO_OWNER=' }) -replace '^GITHUB_REPO_OWNER=', '').Trim()
$repoName = (($envLines | Where-Object { $_ -match '^GITHUB_REPO_NAME=' }) -replace '^GITHUB_REPO_NAME=', '').Trim()
$installerAsset = (($envLines | Where-Object { $_ -match '^INSTALLER_ASSET_NAME=' }) -replace '^INSTALLER_ASSET_NAME=', '').Trim()
if (-not $installerAsset) { $installerAsset = "RECA_INCLUSION_LABORAL_Setup.exe" }

$installerConfig = @"
#define SupabaseUrl "$supabaseUrl"
#define SupabaseKey "$supabaseKey"
#define GithubRepoOwner "$repoOwner"
#define GithubRepoName "$repoName"
#define InstallerAssetName "$installerAsset"
"@
Set-Content -Path (Join-Path $root "installer_config.iss") -Value $installerConfig -Encoding UTF8

$pyiArgs = @(
    "--noconfirm",
    "--clean",
    "--windowed",
    "--name", "RECA_INCLUSION_LABORAL",
    "--add-data", "templates;templates",
    "--add-data", "Diccionario.txt;.",
    "--add-data", "VERSION;.",
    "--add-data", "config.json;.",
    "--hidden-import", "win32com",
    "--hidden-import", "win32com.client",
    "--hidden-import", "pythoncom",
    "--hidden-import", "pywintypes",
    "app.py"
)

& $python -m PyInstaller @pyiArgs
