@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PY_CMD="

where py >nul 2>nul
if not errorlevel 1 set "PY_CMD=py -3"

if not defined PY_CMD (
    where python >nul 2>nul
    if not errorlevel 1 set "PY_CMD=python"
)

if not defined PY_CMD (
    echo.
    echo ERROR: Python was not found.
    echo Install Python and enable Add Python to PATH.
    echo.
    pause
    exit /b 1
)

echo Installing required packages...
%PY_CMD% -m pip install --upgrade pip
if errorlevel 1 goto :fail

%PY_CMD% -m pip install pandas matplotlib openpyxl
if errorlevel 1 goto :fail

echo.
echo Installation completed successfully.
pause
exit /b 0

:fail
echo.
echo Package installation failed.
pause
exit /b 1
