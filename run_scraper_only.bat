@echo off
setlocal enabledelayedexpansion

REM ==== Paths ====
set "PROJECT_DIR=C:\Projects\MaslahaScheduler"
set "PYTHON=%PROJECT_DIR%\.venv\Scripts\python.exe"

REM ==== Debug info ====
echo PROJECT_DIR = %PROJECT_DIR%
echo PYTHON      = %PYTHON%
echo Current user: %USERNAME%
echo Current dir before cd: %CD%
cd /d "%PROJECT_DIR%"
echo Current dir after cd: %CD%
echo.

REM ==== Check python path ====
if exist "%PYTHON%" (
    echo [OK] Found Python at %PYTHON%
) else (
    echo [ERROR] Python not found at %PYTHON%
    exit /b 1
)

REM ==== Run data_scraper ====
echo Running data_scraper.py ...
"%PYTHON%" "%PROJECT_DIR%\data_scraper.py"
echo Exit code = %ERRORLEVEL%
echo.

endlocal
