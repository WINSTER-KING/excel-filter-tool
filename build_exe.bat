@echo off
setlocal
title Build Excel Filter Tool
echo ============================================
echo   Excel Filter Tool - one-click build
echo ============================================
echo.

REM ---------- 0. Switch to script folder ----------
cd /d "%~dp0"

REM ---------- 1. Check Python ----------
where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python not found.
    echo.
    echo Please install Python 3.9+ from https://www.python.org
    echo and make sure to check "Add Python to PATH" during install.
    echo.
    pause
    exit /b 1
)

echo [1/6] Python detected:
python --version
echo.

REM ---------- 2. Install dependencies ----------
echo [2/6] Installing openpyxl and pyinstaller ...
python -m pip install --upgrade openpyxl pyinstaller
if errorlevel 1 (
    echo [ERROR] Dependency install failed. Check your network and retry.
    pause
    exit /b 1
)
echo.

REM ---------- 3. Check source files ----------
if not exist "excel_filter_tool.py" (
    echo [ERROR] excel_filter_tool.py not found in this folder.
    echo Please put excel_filter_tool.py next to this script and retry.
    pause
    exit /b 1
)
echo [3/6] Source file OK.
echo.

REM ---------- 4. Build ----------
echo [4/6] Building, please wait ...
if exist "app.ico" (
    echo       Using icon: app.ico
    python -m PyInstaller -n ExcelFilterTool -F -w --clean --icon app.ico excel_filter_tool.py
) else (
    echo       [WARN] app.ico not found, building without an icon.
    python -m PyInstaller -n ExcelFilterTool -F -w --clean excel_filter_tool.py
)
if errorlevel 1 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)
echo.

REM ---------- 5. Copy docs (any .html / .md files) ----------
echo [5/6] Copying docs ...
copy /y *.html "dist\" >nul 2>nul
copy /y *.md "dist\" >nul 2>nul
echo.

REM ---------- 6. Done ----------
echo [6/6] Build complete!
echo.
echo The executable is located at:
echo     %~dp0dist\ExcelFilterTool.exe
echo Docs (if present) were copied into the dist folder too.
echo.
echo You can copy this exe anywhere and double-click to run it.
echo No Python installation is needed on the target machine.
echo.
pause
