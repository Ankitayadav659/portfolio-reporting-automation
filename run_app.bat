@echo off
rem ============================================================
rem  Portfolio Report Maker - double-click to open the window.
rem  First run (on any PC):
rem    - installs missing packages (openpyxl, pdfplumber, tkinterdnd2)
rem    - puts a "Portfolio Report Maker" shortcut on the Desktop
rem  After that the window opens with no black console left behind.
rem ============================================================
if not "%~1"=="min" (
    start "" /min cmd /c ""%~f0" min"
    exit /b
)
cd /d "%~dp0"

rem ---- is Python there? ("py" launcher, or plain "python") ----
set "PY=py"
%PY% -c "import sys" >nul 2>&1
if errorlevel 1 (
    set "PY=python"
    python -c "import sys" >nul 2>&1
    if errorlevel 1 (
        start "Python not found" cmd /k "echo Python was not found on this PC. & echo Install it from python.org - tick 'Add python.exe to PATH' - then run this again. & echo. & pause & exit"
        exit /b
    )
)

rem ---- install missing packages (visible window only if needed) ----
%PY% -c "import openpyxl, pdfplumber, tkinterdnd2" >nul 2>&1
if errorlevel 1 (
    start "Installing packages - one time only" /wait cmd /c "echo Installing openpyxl, pdfplumber, tkinterdnd2 (one time only)... & echo. & %PY% -m pip install --user openpyxl pdfplumber tkinterdnd2 & if errorlevel 1 (echo. & echo Install FAILED - check the internet connection and try again. & pause)"
)

rem ---- one time: Desktop shortcut (skipped if it already exists) ----
set "APP_BAT=%~f0"
set "APP_DIR=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$lnk = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Portfolio Report Maker.lnk'; if (-not (Test-Path -LiteralPath $lnk)) { $s = (New-Object -ComObject WScript.Shell).CreateShortcut($lnk); $s.TargetPath = $env:APP_BAT; $s.WorkingDirectory = $env:APP_DIR; $s.WindowStyle = 7; $s.IconLocation = 'imageres.dll,109'; $s.Save() }" >nul 2>&1

rem ---- start the window with pythonw.exe (no console) ----
set "PYW="
for /f "delims=" %%p in ('%PY% -c "import sys,os;print(os.path.join(os.path.dirname(sys.executable),'pythonw.exe'))"') do set "PYW=%%p"
if exist "%PYW%" (
    start "" "%PYW%" "%~dp0gui_app.py"
) else (
    start "" pyw "%~dp0gui_app.py"
)
exit /b
