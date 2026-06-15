@echo off
setlocal enabledelayedexpansion
title NJN Comparator Setup

echo.
echo   ╔══════════════════════════════════╗
echo   ║      NJN COMPARATOR SETUP        ║
echo   ╚══════════════════════════════════╝
echo.

:: ── 1. Check for Python ──────────────────────────────────────────────────────
python --version >nul 2>&1
if !errorlevel! neq 0 (
    echo   Python not found. Downloading Python 3.12 installer...
    echo   This may take a minute. Please wait.
    echo.
    curl -L --progress-bar -o "%TEMP%\python_setup.exe" ^
        "https://www.python.org/ftp/python/3.12.4/python-3.12.4-amd64.exe"
    echo   Installing Python 3.12 (silent)...
    "%TEMP%\python_setup.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1
    del "%TEMP%\python_setup.exe"
    :: Reload PATH so python is found
    call refreshenv >nul 2>&1
    python --version >nul 2>&1
    if !errorlevel! neq 0 (
        echo.
        echo   [!] Python installed but PATH not updated yet.
        echo   [!] Please CLOSE this window and double-click install_windows.bat again.
        pause
        exit /b 1
    )
)

for /f "tokens=*" %%v in ('python --version') do set PYVER=%%v
echo   ✓ %PYVER%

:: ── 2. Install packages ───────────────────────────────────────────────────────
echo   Installing required packages (first time only)...
python -m pip install --quiet --upgrade customtkinter openpyxl
echo   ✓ Packages ready

:: ── 3. Launch ─────────────────────────────────────────────────────────────────
echo   Launching NJN Comparator...
echo.
cd /d "%~dp0"
python app.py
