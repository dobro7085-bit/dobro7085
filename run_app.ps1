$ErrorActionPreference = "Stop"

Set-Location -LiteralPath $PSScriptRoot

$venvPython = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
$bundledPython = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$pythonExe = $null

if (Test-Path -LiteralPath $venvPython) {
    $pythonExe = $venvPython
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $pythonExe = "python"
} elseif (Get-Command py -ErrorAction SilentlyContinue) {
    $pythonExe = "py"
} elseif (Test-Path -LiteralPath $bundledPython) {
    $pythonExe = $bundledPython
}

if (-not $pythonExe) {
    Write-Host "Python을 찾을 수 없습니다." -ForegroundColor Red
    Write-Host "Python 3.10 이상을 설치한 뒤 다시 실행해 주세요."
    Write-Host "https://www.python.org/downloads/"
    Read-Host "Enter 키를 누르면 종료합니다"
    exit 1
}

if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host "가상환경을 생성합니다..."
    & $pythonExe -m venv ".venv"
}

Write-Host "필요한 패키지를 설치/확인합니다..."
& $venvPython -m pip install -r requirements.txt

Write-Host "AGM 감액량 분석 대시보드를 실행합니다..."
& $venvPython -m streamlit run app.py
