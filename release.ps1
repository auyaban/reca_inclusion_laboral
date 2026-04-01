param(
    [Parameter(Position = 0)]
    [string]$Version = "",
    [switch]$CleanBuild,
    [switch]$ForceDependencyInstall
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

function Write-Utf8NoBomFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Value
    )

    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Value, $encoding)
}

$versionPath = Join-Path $root "VERSION"
if ($Version) {
    $version = $Version.TrimStart("v")
    Write-Utf8NoBomFile -Path $versionPath -Value "$version`n"
} else {
    if (!(Test-Path $versionPath)) {
        Write-Host "VERSION no encontrado. Usa: .\release.ps1 vX.Y.Z"
        exit 1
    }
    $version = (Get-Content $versionPath | Select-Object -First 1).Trim()
    if (-not $version) {
        Write-Host "VERSION esta vacio. Usa: .\release.ps1 vX.Y.Z"
        exit 1
    }
}

$gh = "gh"
if (-not (Get-Command $gh -ErrorAction SilentlyContinue)) {
    $ghPath = "C:\Program Files\GitHub CLI\gh.exe"
    if (Test-Path $ghPath) {
        $gh = $ghPath
    } else {
        Write-Host "GitHub CLI no encontrado."
        exit 1
    }
}

& $gh auth status -h github.com | Out-Null

$buildArgs = @(
    "-ExecutionPolicy", "Bypass",
    "-File", "build.ps1"
)
if ($CleanBuild) {
    $buildArgs += "-Clean"
}
if ($ForceDependencyInstall) {
    $buildArgs += "-ForceDependencyInstall"
}
powershell @buildArgs

& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DMyAppVersion=$version installer.iss

$installerPath = Join-Path $root "installer\RECA_INCLUSION_LABORAL_Setup.exe"
$hash = (Get-FileHash -Algorithm SHA256 $installerPath).Hash.ToLower()
Write-Utf8NoBomFile `
  -Path (Join-Path $root "installer\RECA_INCLUSION_LABORAL_Setup.exe.sha256") `
  -Value "$hash  RECA_INCLUSION_LABORAL_Setup.exe`n"

$releaseTag = "v$version"
$exists = $true
try {
    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    & $gh release view $releaseTag 2>$null | Out-Null
    $exitCode = $LASTEXITCODE
    $ErrorActionPreference = $prev
} catch {
    $ErrorActionPreference = $prev
    $exitCode = 1
}
if ($exitCode -ne 0) {
    $exists = $false
}

if ($exists) {
    Write-Host "Release $releaseTag encontrado. Subiendo assets..."
    & $gh release upload $releaseTag $installerPath (Join-Path $root "installer\RECA_INCLUSION_LABORAL_Setup.exe.sha256") --clobber
} else {
    Write-Host "Release $releaseTag no existe. Creandolo..."
    & $gh release create $releaseTag $installerPath (Join-Path $root "installer\RECA_INCLUSION_LABORAL_Setup.exe.sha256") `
      --title "RECA Inclusion Laboral v$version" `
      --notes "Release v$version"
}
