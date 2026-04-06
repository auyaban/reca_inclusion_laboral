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

function Get-MainWorktreeRoot {
    param(
        [Parameter(Mandatory = $true)][string]$CurrentRoot
    )

    $gitCmd = Get-Command git -ErrorAction SilentlyContinue
    if (-not $gitCmd) {
        return $null
    }

    try {
        $lines = & $gitCmd.Source worktree list --porcelain 2>$null
        if ($LASTEXITCODE -ne 0 -or -not $lines) {
            return $null
        }
    } catch {
        return $null
    }

    $entries = @()
    $current = @{}
    foreach ($line in $lines) {
        if ($line -like "worktree *") {
            if ($current.Count -gt 0) {
                $entries += [pscustomobject]$current
            }
            $current = @{ worktree = $line.Substring(9).Trim() }
            continue
        }
        if (-not $line) {
            continue
        }
        $spaceIndex = $line.IndexOf(' ')
        if ($spaceIndex -lt 0) {
            continue
        }
        $key = $line.Substring(0, $spaceIndex)
        $value = $line.Substring($spaceIndex + 1).Trim()
        $current[$key] = $value
    }
    if ($current.Count -gt 0) {
        $entries += [pscustomobject]$current
    }

    $mainEntry = $entries | Where-Object { $_.branch -eq "refs/heads/main" } | Select-Object -First 1
    if (-not $mainEntry) {
        return $null
    }

    $mainRoot = [System.IO.Path]::GetFullPath([string]$mainEntry.worktree)
    $currentFull = [System.IO.Path]::GetFullPath($CurrentRoot)
    if ($mainRoot -eq $currentFull) {
        return $null
    }
    return $mainRoot
}

function Sync-InstallerArtifactsToMainWorktree {
    param(
        [Parameter(Mandatory = $true)][string]$CurrentRoot,
        [Parameter(Mandatory = $true)][string[]]$ArtifactPaths
    )

    $mainRoot = Get-MainWorktreeRoot -CurrentRoot $CurrentRoot
    if (-not $mainRoot) {
        return
    }

    $targetInstallerDir = Join-Path $mainRoot "installer"
    New-Item -ItemType Directory -Path $targetInstallerDir -Force | Out-Null

    foreach ($artifact in $ArtifactPaths) {
        if (-not (Test-Path $artifact)) {
            continue
        }
        $targetPath = Join-Path $targetInstallerDir (Split-Path $artifact -Leaf)
        Copy-Item -Path $artifact -Destination $targetPath -Force
        Write-Host "Artefacto sincronizado a worktree principal: $targetPath"
    }
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
$installerHashPath = Join-Path $root "installer\RECA_INCLUSION_LABORAL_Setup.exe.sha256"
$hash = (Get-FileHash -Algorithm SHA256 $installerPath).Hash.ToLower()
Write-Utf8NoBomFile `
  -Path $installerHashPath `
  -Value "$hash  RECA_INCLUSION_LABORAL_Setup.exe`n"

Sync-InstallerArtifactsToMainWorktree -CurrentRoot $root -ArtifactPaths @(
    $installerPath,
    $installerHashPath
)

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
    & $gh release upload $releaseTag $installerPath $installerHashPath --clobber
} else {
    Write-Host "Release $releaseTag no existe. Creandolo..."
    & $gh release create $releaseTag $installerPath $installerHashPath `
      --title "RECA Inclusion Laboral v$version" `
      --notes "Release v$version"
}
