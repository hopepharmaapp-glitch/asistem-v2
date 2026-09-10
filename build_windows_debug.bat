@echo off
setlocal enableextensions
cd /d "%~dp0"

echo ========================================================
echo   HOPEPHARMA DEBUG BUILDER (LOGGING ENABLED)
echo ========================================================
echo.
echo This script will save all output to "build_log.txt".
echo If it fails, please send me the content of "build_log.txt".
echo.

(
    echo STARTING BUILD AT %TIME% on %DATE%
    echo.
    echo [INFO] Python Version:
    python --version
    echo.
    echo [INFO] PIP Version:
    pip --version
    echo.
    
    echo [INFO] Installing build tools...
    pip install --upgrade pip
    pip install pyinstaller reportlab pillow
    
    echo.
    echo [INFO] Starting PyInstaller...
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
        --log-level DEBUG ^
        "hope_pharma_complete.py"
        
    echo.
    echo FINISHED BUILD AT %TIME%
) > build_log.txt 2>&1

echo.
if exist "dist\HopePharma.exe" (
    echo [SUCCESS] Build completed. App is in 'dist' folder.
) else (
    echo [FAILURE] Build failed.
    echo.
    echo ========================================================
    echo   ERROR LOG (LAST 20 LINES):
    echo ========================================================
    powershell -Command "Get-Content build_log.txt -Tail 20"
    echo.
    echo ========================================================
    echo FULL LOG SAVED TO: %~dp0build_log.txt
    echo Please open build_log.txt to see the exact error.
    echo ========================================================
)
pause
