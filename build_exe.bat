@echo off
echo ============================================
echo  Building WO Flow Planning App (.exe)
echo ============================================

REM Activate venv if present
if exist venv\Scripts\activate.bat (
    call venv\Scripts\activate.bat
)

REM Install / upgrade PyInstaller
pip install --quiet --upgrade pyinstaller

REM Clean previous build
if exist dist\WO_Planning.exe del /f /q dist\WO_Planning.exe
if exist build rmdir /s /q build

REM Build single-file exe
pyinstaller WO_Planning.spec

echo.
if exist dist\WO_Planning.exe (
    echo  SUCCESS — exe is at:  dist\WO_Planning.exe
    echo.
    echo  IMPORTANT: Copy BOTH files to share with others:
    echo    1. dist\WO_Planning.exe
    echo    2. WO_Planning_App.xlsx   (station config - must stay next to the exe)
    echo.
) else (
    echo  BUILD FAILED — check the output above for errors.
)

pause
