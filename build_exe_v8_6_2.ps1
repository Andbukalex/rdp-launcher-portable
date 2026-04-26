# Build-Script fuer RDP Launcher 8.6.2
# Erwartet im gleichen Ordner: launcher_v8_6_2.ps1
# Optional: rdp-launcher.ico fuer das EXE-Icon

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$Source = Join-Path $ScriptDir "launcher_v8_6_2.ps1"
$Icon = Join-Path $ScriptDir "rdp-launcher.ico"
$OutDir = Join-Path $ScriptDir "dist"
$Target = Join-Path $OutDir "RDP-Launcher.exe"

if (-not (Test-Path $Source)) {
    Write-Host "FEHLER: launcher_v8_6_2.ps1 wurde nicht gefunden." -ForegroundColor Red
    Write-Host "Lege diese Datei in denselben Ordner wie build_exe_v8_6_2.ps1." -ForegroundColor Yellow
    pause
    exit 1
}

if (-not (Get-Command Invoke-ps2exe -ErrorAction SilentlyContinue)) {
    Write-Host "PS2EXE nicht gefunden. Installation wird versucht..." -ForegroundColor Yellow
    Install-Module ps2exe -Scope CurrentUser -Force
}

if (-not (Test-Path $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir | Out-Null
}

Write-Host "Baue RDP-Launcher.exe ..." -ForegroundColor Cyan

$ps2exeParams = @{
    InputFile   = $Source
    OutputFile  = $Target
    NoConsole   = $true
    Title       = "RDP Launcher"
    Description = "Portable RDP Launcher"
    Company     = "Andreas Husemann"
    Product     = "RDP Launcher"
    Copyright   = "Andreas Husemann"
    Version     = "8.6.2.0"
}

if (Test-Path $Icon) {
    $ps2exeParams.IconFile = $Icon
    Write-Host "Icon wird eingebunden: $Icon" -ForegroundColor Cyan
} else {
    Write-Host "Hinweis: rdp-launcher.ico nicht gefunden. Build erfolgt ohne Icon." -ForegroundColor Yellow
}

Invoke-ps2exe @ps2exeParams

Write-Host "Fertig:" -ForegroundColor Green
Write-Host $Target -ForegroundColor Green
pause
