@echo off
title QSA Audit Search MES EXE Build
cd /d "%~dp0"

echo ==========================================
echo QSA Audit Search MES - EXE Build
echo ==========================================
echo.

echo [1/3] Installing required Python packages...
py -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ERROR: Package install failed.
    echo Check your internet connection and Python installation.
    echo.
    pause
    exit /b 1
)

echo.
echo [2/3] Building EXE file...
py -m PyInstaller --onefile --noconsole --name "QSA Audit Search MES" main.py
if errorlevel 1 (
    echo.
    echo ERROR: EXE build failed.
    echo.
    pause
    exit /b 1
)

echo.
echo [3/3] Done.
echo.
echo EXE file:
echo %~dp0dist\QSA Audit Search MES.exe
echo.
pause
