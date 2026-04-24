@echo off
setlocal

cd /d "%~dp0"

echo.
echo [1/4] Python version check
python --version
if errorlevel 1 (
  echo.
  echo Python is not installed or PATH is not set.
  echo Please read 먼저읽기_수정방법.txt and install Python first.
  pause
  exit /b 1
)

echo.
echo [2/4] Installing PyInstaller
python -m pip install --upgrade pip
python -m pip install pyinstaller
if errorlevel 1 (
  echo.
  echo PyInstaller install failed.
  echo If company internet is blocked, build the EXE on another PC.
  pause
  exit /b 1
)

echo.
echo [3/4] Building MES_v1.exe
cd /d "%~dp0dev"
python -m PyInstaller --onefile --windowed --name MES_v1 MES_v1.py
if errorlevel 1 (
  echo.
  echo EXE build failed.
  pause
  exit /b 1
)

echo.
echo [4/4] Copying new EXE to exe folder
cd /d "%~dp0"
copy /Y "%~dp0dev\dist\MES_v1.exe" "%~dp0exe\MES_v1.exe"
if not exist "%~dp0exe\quality_history.db" (
  copy /Y "%~dp0data\quality_history.db" "%~dp0exe\quality_history.db"
)

echo.
echo Done.
echo New file: exe\MES_v1.exe
echo Keep exe\quality_history.db in the same folder.
echo.
pause

