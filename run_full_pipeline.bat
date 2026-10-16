@echo off
setlocal enabledelayedexpansion

REM ==== Paths ====
set "PROJECT_DIR=%~dp0."
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
echo Running scrapping_tool.py ...
"%PYTHON%" "%PROJECT_DIR%\scrapping_tool.py"
echo Exit code = %ERRORLEVEL%
echo.

REM ==== Run json_parser ====
echo Running json_extractor.py ...
"%PYTHON%" "%PROJECT_DIR%\json_extractor.py"
echo Exit code = %ERRORLEVEL%
echo.

REM ==== Run send_email ====
echo Running send_email.py ...
"%PYTHON%" "%PROJECT_DIR%\send_email.py"
echo Exit code = %ERRORLEVEL%
echo.

endlocal
