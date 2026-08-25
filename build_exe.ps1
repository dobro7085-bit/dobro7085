$ErrorActionPreference = "Stop"

Set-Location -LiteralPath $PSScriptRoot

$pythonExe = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $pythonExe)) {
    Write-Host "Python virtual environment was not found." -ForegroundColor Red
    Write-Host "Run run_app.bat once first, then run this file again."
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Installing or checking PyInstaller..."
& $pythonExe -m pip install pyinstaller

Write-Host "Building distribution folder..."
& $pythonExe -m PyInstaller --noconfirm --clean "AGM_Reduction_Analysis.spec"

$distDir = Join-Path $PSScriptRoot "dist\AGM_Reduction_Analysis"
$standardDir = Join-Path $distDir "app\data"
New-Item -ItemType Directory -Force -Path $standardDir | Out-Null
Copy-Item -LiteralPath (Join-Path $PSScriptRoot "app\data\standard.xlsx") -Destination (Join-Path $standardDir "standard.xlsx") -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot "README_팀원용.txt") -Destination (Join-Path $distDir "README_팀원용.txt") -Force

Write-Host ""
Write-Host "Build completed." -ForegroundColor Green
Write-Host "Distribution folder:"
Write-Host $distDir
