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

if "%~1"=="" (
    echo.
    echo Drag and drop an XODR file onto this BAT file.
    echo Or type the full XODR path below.
    echo.
    set /p "XODR_PATH=XODR path: "
) else (
    set "XODR_PATH=%~f1"
)

if not exist "%XODR_PATH%" (
    echo.
    echo ERROR: File not found:
    echo "%XODR_PATH%"
    echo.
    pause
    exit /b 1
)

echo.
echo Starting XODR analysis...
echo Input: "%XODR_PATH%"
echo.

%PY_CMD% "%~dp0xodr_analyzer.py" "%XODR_PATH%"
set "RESULT=%ERRORLEVEL%"

echo.
if not "%RESULT%"=="0" (
    echo Analysis failed. Error code: %RESULT%
    echo Check error.log in the result folder.
    pause
    exit /b %RESULT%
)

echo Analysis completed successfully.
echo Check the *_analysis_result folder next to the XODR file.
echo.
pause
