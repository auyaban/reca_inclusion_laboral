param(
    [Parameter(Mandatory = $true)]
    [string]$Version,
    [string]$Notes = "",
    [switch]$AutoConfirm
)

$ErrorActionPreference = "Stop"

function Read-YesNo([string]$Prompt) {
    while ($true) {
        $answer = Read-Host "$Prompt (s/n)"
        if (-not $answer) { continue }
        switch ($answer.Trim().ToLower()) {
            "s" { return $true }
            "n" { return $false }
        }
    }
}

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$ver = $Version.TrimStart("v")
if (-not $ver) {
    Write-Host "Version invalida."
    exit 1
}

Set-Content -Path (Join-Path $root "VERSION") -Value $ver -Encoding UTF8

Write-Host "Publicando version $ver..."
if (-not $AutoConfirm) {
    if (-not (Read-YesNo "Continuar con commit, push y release")) {
        Write-Host "Cancelado."
        exit 0
    }
}

git add VERSION
git add -A
git commit -m "Release $ver"
git push

powershell -ExecutionPolicy Bypass -File release.ps1 "v$ver"
