@echo off
setlocal EnableExtensions
cd /d "%~dp0"

where pyw >nul 2>nul
if not errorlevel 1 (
    start "" pyw -3 "%~dp0xodr_analyzer_gui.pyw"
    exit /b 0
)

where pythonw >nul 2>nul
if not errorlevel 1 (
    start "" pythonw "%~dp0xodr_analyzer_gui.pyw"
    exit /b 0
)

echo ERROR: Python was not found.
pause
exit /b 1
