@echo off
setlocal

cd /d "%~dp0"

set "VENV_PY=%CD%\.venv\Scripts\python.exe"
set "BUNDLED_PY=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if not exist "%VENV_PY%" (
    echo Creating virtual environment...
    if exist "%BUNDLED_PY%" (
        call "%BUNDLED_PY%" -m venv ".venv"
    ) else (
        python -m venv ".venv"
    )
)

if not exist "%VENV_PY%" (
    echo Failed to create Python virtual environment.
    echo Please install Python 3.10 or later and run this file again.
    echo https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Installing or checking required packages...
call "%VENV_PY%" -m pip install -r requirements.txt
if errorlevel 1 (
    echo Package installation failed.
    echo Please check your internet connection or company security policy.
    pause
    exit /b 1
)

echo Starting AGM dashboard...
echo Browser URL: http://localhost:8501
call "%VENV_PY%" -m streamlit run app.py --server.address localhost --server.port 8501 --server.fileWatcherType none --server.runOnSave false --server.enableWebsocketCompression false --server.websocketPingInterval 10 --server.disconnectedSessionTTL 3600 --browser.gatherUsageStats false

pause
