@echo off
setlocal
cd /d "%~dp0"

set "PYTHON=%~dp0.venv\Scripts\python.exe"

if not exist "%PYTHON%" (
    where py >nul 2>nul
    if errorlevel 1 (
        echo Python 3 is required. Install Python 3.10 or newer, then run this file again.
        pause
        exit /b 1
    )
    echo Creating the local Python environment...
    py -3 -m venv "%~dp0.venv"
    if errorlevel 1 goto :error
)

echo Installing app and packaging dependencies...
"%PYTHON%" -m pip install --index-url https://pypi.org/simple -e ".[build]"
if errorlevel 1 goto :error

echo Building the standalone Windows application...
"%PYTHON%" -m PyInstaller --noconfirm --clean --onefile --windowed --name "KPI Pivot Chart Creator" --distpath "%~dp0dist\current" --paths "%~dp0src" --add-data "%~dp0make_pivot.ps1;." "%~dp0main.py"
if errorlevel 1 goto :error

echo Build complete. Opening the app...
start "" "%~dp0dist\current\KPI Pivot Chart Creator.exe"
exit /b 0

:error
echo.
echo The build did not complete. Review the error above and try again.
pause
exit /b 1
