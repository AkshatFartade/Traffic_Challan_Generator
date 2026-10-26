@echo off
setlocal
cd /d "%~dp0"

echo =====================================================================
echo  Traffic Notice Generator — Local Official Web Interface
echo =====================================================================

if not exist ".venv\Scripts\activate.bat" (
    echo [ERROR] Virtual environment not found in .venv.
    echo Please create it first: python -m venv .venv
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat

if "%CHALLAN_SUPERVISOR_PASSWORD%"=="" (
    echo [NOTE] CHALLAN_SUPERVISOR_PASSWORD environment variable is not set.
    echo To enable login, set it before running, e.g.:
    echo   set CHALLAN_SUPERVISOR_PASSWORD=MySecretPassphrase
    echo.
)

echo Starting server with waitress on http://127.0.0.1:5000 ...
python -m waitress --listen=127.0.0.1:5000 --call "app:create_app"

endlocal
