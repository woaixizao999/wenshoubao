param(
    [string]$Python = "C:\Users\bangl\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir
$AppName = "win"

& $Python -m PyInstaller `
  --noconfirm `
  --clean `
  --onefile `
  --windowed `
  --icon ".\assets\win.ico" `
  --add-data ".\assets\win.ico;assets" `
  --name $AppName `
  ".\wenshoubao.py"

Write-Host "Build complete: $ScriptDir\dist\$AppName.exe"
