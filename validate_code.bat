@echo off
REM VBA Code Validation Script for Windows
REM Double-click this file to validate your latest BWS code

echo.
echo ============================================================
echo          BWS VBA Code Validator
echo ============================================================
echo.

REM Find the latest BWS version file by modification date
set LATEST_FILE=
for /f "delims=" %%f in ('dir /b /o-d BWS_v1.7.*.bas 2^>nul') do (
    if not defined LATEST_FILE set LATEST_FILE=%%f
)

if "%LATEST_FILE%"=="" (
    echo ERROR: No BWS_v1.7.*.bas files found in current directory
    echo.
    echo Please make sure you're running this from the VBA folder
    echo Press any key to exit...
    pause >nul
    exit /b 1
)

echo Checking most recently modified file: %LATEST_FILE%
echo.

REM Run the Python validation script
python check_vba_code.py %LATEST_FILE%

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ============================================================
    echo   SUCCESS! Your code passed all validation checks
    echo ============================================================
) else (
    echo.
    echo ============================================================
    echo   ATTENTION: Please fix the issues listed above
    echo ============================================================
)

echo.
echo Press any key to exit...
pause >nul
