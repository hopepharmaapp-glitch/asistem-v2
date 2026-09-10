@echo off
setlocal enableextensions
cd /d "%~dp0"

echo ========================================================
echo   HOPEPHARMA FINAL BUILDER (PATH FIX)
echo ========================================================
echo.

:: Check Python availability
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not recognized.
    echo Please install Python 3.11 from the file "python-3.11.9-amd64.exe"
    echo AND CHECK THE BOX "Add Python to PATH" on the first screen.
    pause
    exit /b 1
)

echo [1/3] Installing PyInstaller...
python -m pip install --upgrade pip --quiet
python -m pip install pyinstaller reportlab pillow --quiet

echo [2/3] Building App (Using python -m PyInstaller)...
:: Using 'python -m PyInstaller' bypasses the PATH issue
python -m PyInstaller --clean --noconfirm --onefile --windowed ^
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
    echo [FATAL ERROR] Build failed even with python -m.
    pause
    exit /b 1
)

echo.
echo [3/3] Success!
if exist "dist\HopePharma.exe" (
    echo App created at: dist\HopePharma.exe
    pause
) else (
    echo Error: dist\HopePharma.exe not found.
    pause
)
