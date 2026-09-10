@echo off
setlocal enableextensions
cd /d "%~dp0"

echo ========================================================
echo   HOPEPHARMA ROBUST WINDOWS BUILDER
echo ========================================================
echo.

:: 1. Check for Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install the "python-3.11.9-amd64.exe" file in this folder.
    echo IMPORTANT: Check the box "Add Python to PATH" during installation.
    pause
    exit /b 1
)

:: 2. Check for PIP
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] PIP is missing. Reinstall Python.
    pause
    exit /b 1
)

echo [1/4] Installing/Upgrading build tools...
pip install --upgrade pip --quiet
pip install pyinstaller reportlab pillow --quiet
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies. Check your internet connection.
    pause
    exit /b 1
)

echo [2/4] Cleaning old build files...
if exist build rmdir /s /q build
if exist dist\HopePharma.exe del /f /q dist\HopePharma.exe
if exist dist\HopePharma rmdir /s /q dist\HopePharma
if exist *.spec del /f /q *.spec

echo [3/4] Building Windows Executable (Please wait ~1 minute)...
:: Build command with strict windowed mode and all assets
pyinstaller --clean --noconfirm --onefile --windowed ^
    --name "HopePharma" ^
    --icon "logo.ico" ^
    --hidden-import "invoice_app" ^
    --add-data "AssistemLogo.png;." ^
    --add-data "logo.png;." ^
    --add-data "sign1.png;." ^
    --add-data "sign2.png;." ^
    --collect-all "reportlab" ^
    --collect-all "PIL" ^
    "hope_pharma_complete.py"

if %errorlevel% neq 0 (
    echo.
    echo [FATAL ERROR] Build failed.
    echo Please take a photo of the error message above and show it to support.
    pause
    exit /b 1
)

echo.
echo [4/4] Verifying build...
if exist "dist\HopePharma.exe" (
    echo.
    echo ========================================================
    echo   SUCCESS! APP CREATED:
    echo   %~dp0dist\HopePharma.exe
    echo ========================================================
    echo.
    echo You can now copy "HopePharma.exe" anywhere and run it.
    echo.
    pause
) else (
    echo [ERROR] The build finished but HopePharma.exe is missing.
    echo Antivirus might have deleted it. Check your protection history.
    pause
)

endlocal
