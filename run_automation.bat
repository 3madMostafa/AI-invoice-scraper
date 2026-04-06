@echo off
setlocal enabledelayedexpansion

REM ==== Paths ====
set "PROJECT_DIR=C:\Projects\MaslahaScheduler"
set "PYTHON=%PROJECT_DIR%\.venv\Scripts\python.exe"
set "LOG_DIR=%PROJECT_DIR%\logs"

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

REM ==== Timestamp for log file ====
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HH-mm-ss"') do set "RUN_TS=%%i"
set "LOG_FILE=%LOG_DIR%\pipeline_%RUN_TS%.log"

echo ========================================= >> "%LOG_FILE%"
echo INVOICE AUTOMATION PIPELINE >> "%LOG_FILE%"
echo ========================================= >> "%LOG_FILE%"
echo Started at: %DATE% %TIME% >> "%LOG_FILE%"
echo PROJECT_DIR = %PROJECT_DIR% >> "%LOG_FILE%"
echo PYTHON      = %PYTHON% >> "%LOG_FILE%"
echo Current user: %USERNAME% >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

echo =========================================
echo INVOICE AUTOMATION PIPELINE
echo =========================================
echo Log file: %LOG_FILE%
echo.

REM ==== Check project dir ====
if not exist "%PROJECT_DIR%" (
    echo [ERROR] Project directory not found: %PROJECT_DIR%
    echo [ERROR] Project directory not found: %PROJECT_DIR% >> "%LOG_FILE%"
    exit /b 1
)

cd /d "%PROJECT_DIR%"
echo Working dir: %CD%
echo Working dir: %CD% >> "%LOG_FILE%"

REM ==== Check python ====
if exist "%PYTHON%" (
    echo [OK] Python found
    echo [OK] Python found >> "%LOG_FILE%"
) else (
    echo [ERROR] Python not found at %PYTHON%
    echo [ERROR] Python not found at %PYTHON% >> "%LOG_FILE%"
    exit /b 1
)

echo. >> "%LOG_FILE%"
echo ========================================= >> "%LOG_FILE%"
echo STEP 1 - DOWNLOAD INVOICES FROM ETA API >> "%LOG_FILE%"
echo ========================================= >> "%LOG_FILE%"

echo =========================================
echo STEP 1 - DOWNLOAD INVOICES FROM ETA API
echo =========================================

"%PYTHON%" "%PROJECT_DIR%\get_invoices_using_api.py" >> "%LOG_FILE%" 2>&1
set "STEP1_EXIT=%ERRORLEVEL%"
echo Exit code = %STEP1_EXIT%
echo Exit code = %STEP1_EXIT% >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

if not "%STEP1_EXIT%"=="0" (
    echo [ERROR] STEP 1 failed. Stopping pipeline.
    echo [ERROR] STEP 1 failed. Stopping pipeline. >> "%LOG_FILE%"
    exit /b %STEP1_EXIT%
)

echo ========================================= >> "%LOG_FILE%"
echo STEP 2 - PROCESS JSON TO EXCEL >> "%LOG_FILE%"
echo ========================================= >> "%LOG_FILE%"

echo =========================================
echo STEP 2 - PROCESS JSON TO EXCEL
echo =========================================

"%PYTHON%" "%PROJECT_DIR%\invoices_preprocessing.py" >> "%LOG_FILE%" 2>&1
set "STEP2_EXIT=%ERRORLEVEL%"
echo Exit code = %STEP2_EXIT%
echo Exit code = %STEP2_EXIT% >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

if not "%STEP2_EXIT%"=="0" (
    echo [ERROR] STEP 2 failed. Stopping pipeline.
    echo [ERROR] STEP 2 failed. Stopping pipeline. >> "%LOG_FILE%"
    exit /b %STEP2_EXIT%
)

echo ========================================= >> "%LOG_FILE%"
echo STEP 3 - ZIP AND SEND EMAIL >> "%LOG_FILE%"
echo ========================================= >> "%LOG_FILE%"

echo =========================================
echo STEP 3 - ZIP AND SEND EMAIL
echo =========================================

"%PYTHON%" "%PROJECT_DIR%\send_email.py" >> "%LOG_FILE%" 2>&1
set "STEP3_EXIT=%ERRORLEVEL%"
echo Exit code = %STEP3_EXIT%
echo Exit code = %STEP3_EXIT% >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

if not "%STEP3_EXIT%"=="0" (
    echo [ERROR] STEP 3 failed. Stopping pipeline.
    echo [ERROR] STEP 3 failed. Stopping pipeline. >> "%LOG_FILE%"
    exit /b %STEP3_EXIT%
)

echo =========================================
echo PIPELINE FINISHED SUCCESSFULLY
echo =========================================

echo ========================================= >> "%LOG_FILE%"
echo PIPELINE FINISHED SUCCESSFULLY >> "%LOG_FILE%"
echo ========================================= >> "%LOG_FILE%"
echo Finished at: %DATE% %TIME% >> "%LOG_FILE%"

endlocal
exit /b 0