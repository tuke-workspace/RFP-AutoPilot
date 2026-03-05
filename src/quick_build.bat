@echo off
chcp 65001 > nul

:: Check if running as administrator
openfiles >nul 2>&1
if %errorlevel%==0 (
    echo Error: Please do not run this script as administrator!
    echo    Right-click and select "Run as normal user" or double-click to run
    echo.
    pause
    exit /b 1
)

:: Switch to script directory
cd /d "%~dp0"

:: Check if in correct directory
if not exist "main_word_processor.py" (
    echo Error: main_word_processor.py not found!
    echo    Please make sure to run this script in the correct directory
    echo    Current directory: %CD%
    echo.
    pause
    exit /b 1
)

echo ========================================
echo   Word Smart Processor Build Script
echo ========================================
echo Current directory: %CD%
echo.

echo Installing dependencies...
python -m pip install PyQt5 python-docx pywin32 comtypes pyinstaller --no-warn-script-location

echo.
echo Starting build (this may take a few minutes)...

:: Set command parameters
set NAME=RFP-AutoPilot
set DATA1=document_processor.py;.
set DATA2=clause_utils.py;.

:: Run PyInstaller
python -m PyInstaller ^
  --onefile ^
  --windowed ^
  --name="%NAME%" ^
  --version-file="..\version.txt" ^
  --add-data="%DATA1%" ^
  --add-data="%DATA2%" ^
  --hidden-import="win32com.client" ^
  --hidden-import="comtypes.client" ^
  --hidden-import="PyQt5.QtCore" ^
  --hidden-import="PyQt5.QtGui" ^
  --hidden-import="PyQt5.QtWidgets" ^
  main_word_processor.py

if exist "dist\%NAME%.exe" (
    echo.
    echo Build successful! EXE file is in the dist folder
    echo Location: %CD%\dist\%NAME%.exe
    explorer dist
) else (
    echo.
    echo Build failed, please check error messages
    echo.
    echo Trying simplified version...
    python -m PyInstaller --onefile --windowed main_word_processor.py
    if exist "dist\RFP-AutoPilot.exe" (
        echo Simplified version build successful!
        explorer dist
    )
)

pause
